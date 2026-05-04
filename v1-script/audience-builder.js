/**
 * Google Ads Audience Builder
 * Built by Ewald van Kampen (Results Driven Marketing) - https://resultsdriven.nl/
 * MIT License
 *
 * V1 public release.
 *
 * SECURITY CHECKLIST:
 * - Never paste your Gemini API key into a shared or public Google Sheet.
 * - Never commit Google Ads credentials, OAuth tokens, API keys, account IDs,
 *   customer data, or private spreadsheet URLs.
 * - Keep this script in your own Google Ads account unless you have removed
 *   private configuration values.
 *
 * LIVE VALIDATION: Audience creation uses AdsApp.mutate() with Google Ads API
 * mutate payloads. Keep DRY_RUN=true until this script is tested in the target
 * Google Ads account.
 */

var SPREADSHEET_URL = 'TODO: paste your private Google Sheets URL here';
var GEMINI_API_KEY = 'TODO: paste your private Gemini API key here';

var RUN_MODE = 'generate_variants'; // generate_variants or execute
var DRY_RUN = true;

var SCRIPT_VERSION = 'v1.0.0';
var GEMINI_MODEL = 'gemini-2.5-flash';
var GEMINI_API_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta/models/';
var DEFAULT_SPREADSHEET_NAME = 'Google Ads Audience Builder - V1 Template - By Results Driven Marketing';
var DEFAULT_MEMBERSHIP_DURATIONS = [7, 30, 90];
var DEFAULT_TEMPLATE_VALIDATION_ROWS = 500;
var HEADER_BACKGROUND_COLOR = '#fce4d6';
var STATUS_CREATED_BACKGROUND_COLOR = '#d9ead3';
var STATUS_EXISTS_BACKGROUND_COLOR = '#fff2cc';
var STATUS_ERROR_BACKGROUND_COLOR = '#f4cccc';
var MIN_MEMBERSHIP_DURATION = 1;
var MAX_MEMBERSHIP_DURATION = 540;
var YES_VALUE = 'yes';
var NO_VALUE = 'no';

var SHEET_NAMES = {
  HOW_TO_USE: 'HOW TO USE',
  AUDIENCES: 'AUDIENCES',
  VARIANTS: 'VARIANTS'
};

var AUDIENCE_HEADER_ROW = 1;
var AUDIENCE_DATA_START_ROW = 2;

var AUDIENCE_COLUMNS = {
  NAME_PATTERN: 1,
  SEGMENT_TYPE: 2,
  EXPRESSION: 3,
  AUTO_ADD: 4,
  INCLUDE_TAGS: 5,
  EXCLUDE_TAGS: 6,
  CUSTOMER_TYPES: 7,
  NOTE: 8
};

var VARIANT_DURATION_CONFIG = {
  A1: 'B1',
  ROW: 1,
  COLUMN: 2
};

var VARIANT_HEADER_ROW = 3;
var VARIANT_DATA_START_ROW = 4;

var VARIANT_COLUMNS = {
  ROW_NUMBER: 1,
  AUDIENCE_NAME: 2,
  BASED_ON: 3,
  MEMBERSHIP_DURATION: 4,
  CREATE: 5,
  STATUS: 6
};

var VARIANT_HEADERS = [
  '#',
  'audience_name',
  'based_on',
  'membership_duration',
  'create',
  'status'
];

var HOW_TO_USE_CONTENT = [
  ['Google Ads Audience Builder'],
  ['Built by Ewald van Kampen (Results Driven Marketing) - https://resultsdriven.nl/'],
  [''],
  ['Workflow'],
  ['1. Enter audience definitions in AUDIENCES. Use copy-paste for fast bulk creation.'],
  ['Source row numbers are auto-derived from the AUDIENCES row. No manual row-id column is needed.'],
  ['2. For segment_type=website_visitors, use expression. (Use this for URL rules and GA4 events like event = view_item).'],
  ['3. For segment_type=event, use include_tags/exclude_tags. (Use this ONLY for exact Conversion Action names like "Quote Form Submitted"). Separate multiple conversions with commas.'],
  ['4. Enter membership durations in VARIANTS!B1, for example 7,30,90.'],
  ['5. Run the script with RUN_MODE = generate_variants.'],
  ['6. Review VARIANTS and set create to yes or no to cherry-pick which variants to push live (e.g. only 30-day).'],
  ['7. Run with RUN_MODE = execute and DRY_RUN = true for a safe test.'],
  ['8. Set DRY_RUN to false only when you are ready to create Google Ads audiences.'],
  [''],
  ['auto_add'],
  ['yes = ask Google Ads to add historical users who already matched the rule.'],
  ['no = start empty and collect only users who match after audience creation.'],
  [''],
  ['Expression examples: basic'],
  ['URL contains /thank-you'],
  ['event = view_item AND URL contains /pricing'],
  ['segment_type=event with include_tags=Quote Form Submitted (exact conversion name)'],
  ['segment_type=event with include_tags=Quote Form Submitted,Phone Call Lead'],
  ['event = add_to_cart IN 7 DAYS'],
  ['URL starts_with /checkout'],
  ['URL equals /pricing'],
  [''],
  ['Expression examples: conditionals'],
  ['URL contains /thank-you AND event = purchase'],
  ['URL contains /pricing OR URL contains /demo'],
  ['URL contains /product NOT URL contains /return'],
  ['event = add_to_cart NOT event = purchase'],
  [''],
  ['Expression examples: advanced'],
  ['(URL contains /thank-you AND event = purchase) OR event = qualified_lead'],
  ['URL contains /product REFINE offer_detail contains sale IN 30 DAYS'],
  ['(event = add_to_cart IN 7 DAYS OR event = begin_checkout IN 7 DAYS) NOT event = purchase'],
  ['Note: simple URL/event rules parse locally; advanced expressions may use Gemini.'],
  [''],
  ['Status colors'],
  ['created = green, already exists = yellow, error: = red.'],
  [''],
  ['Attribution and liability'],
  ['Do not redistribute modified versions without visible attribution to the original author and source.'],
  ['By using this script, you accept all risk. The author is not liable for outcomes, damage, or policy violations.'],
  [''],
  ['Gemini API key (Optional)'],
  ['The Gemini API key is completely optional. If left empty, the script uses basic local logic (works fine for all standard URL and event rules).'],
  ['Why use it? If you add a key, the script uses AI to parse complex, conversational, or nested rules (like the advanced examples above).'],
  ['1. Go to Google AI Studio (https://aistudio.google.com/) with your Workspace account.'],
  ['2. Create a free Gemini API key.'],
  ['3. Paste the key only in your private Google Ads Script configuration.'],
  [''],
  ['Security'],
  ['Never paste your Gemini API key into a shared or public Google Sheet.']
];

var DEFAULT_AUDIENCE_ROWS = [
  ['Buyers_{duration}d', 'website_visitors', 'URL contains /thank-you', YES_VALUE, '', '', '', 'Thank-you page visitors'],
  ['QuoteLeads_{duration}d', 'event', '', YES_VALUE, 'Quote Form Submitted', '', '', 'Literal conversion action name'],
  ['AllPhoneAndMailLeads_{duration}d', 'event', '', YES_VALUE, 'Phone Call Lead,Mail Click (Website)', '', 'converted_leads', 'Comma-separated conversion names'],
  ['PricingVisitors_{duration}d', 'website_visitors', 'URL contains /pricing', YES_VALUE, '', '', '', 'Pricing page visitors'],
  ['CheckoutVisitors_{duration}d', 'website_visitors', 'URL starts_with /checkout', YES_VALUE, '', '', '', 'Checkout flow visitors'],
  ['SpecificItemViewers_{duration}d', 'website_visitors', 'event = view_item AND URL contains /pricing', YES_VALUE, '', '', '', 'Viewed item and pricing (GA4 event)'],
  ['CartNoPurchase_{duration}d', 'website_visitors', 'event = add_to_cart NOT event = purchase', YES_VALUE, '', '', '', 'Added to cart but no purchase (GA4 events)'],
  ['CheckoutNoPurchase_{duration}d', 'website_visitors', 'event = begin_checkout NOT event = purchase', YES_VALUE, '', '', '', 'Checkout started but no purchase'],
  ['RecentConvertersNoSubscribers_{duration}d', 'event', '', YES_VALUE, 'Quote Form Submitted,Phone Call Lead', 'Paid Subscription', 'converted_leads', 'Conversions excluding other conversions'],
  ['WinbackCandidates_{duration}d', 'website_visitors', '(event = add_to_cart OR event = begin_checkout) NOT event = purchase', YES_VALUE, '', '', 'disengaged_customers', 'Re-engagement audience using GA4 events']
];

var SUPPORTED_SEGMENT_TYPES = [
  'website_visitors',
  'event'
];

var SUPPORTED_CUSTOMER_TYPES = [
  '',
  'all_customers',
  'purchasers',
  'high_value_customers',
  'disengaged_customers',
  'qualified_leads',
  'converted_leads',
  'paying_subscribers',
  'cart_abandoners'
];

var CUSTOMER_TYPE_ENUM_MAP = {
  all_customers: 'ALL_CUSTOMERS',
  purchasers: 'PURCHASERS',
  high_value_customers: 'HIGH_VALUE_CUSTOMERS',
  disengaged_customers: 'DISENGAGED_CUSTOMERS',
  qualified_leads: 'QUALIFIED_LEADS',
  converted_leads: 'CONVERTED_LEADS',
  paying_subscribers: 'PAYING_SUBSCRIBERS',
  cart_abandoners: 'CART_ABANDONERS'
};

var SUPPORTED_OPERATORS = [
  'contains',
  'equals',
  'starts_with',
  'ends_with'
];

var GOOGLE_ADS_FIELD_NAMES = {
  page_url: 'url__',
  event_name: 'event',
  offer_detail: 'offerdetail'
};

var GOOGLE_ADS_STRING_OPERATORS = {
  contains: 'CONTAINS',
  equals: 'EQUALS',
  starts_with: 'STARTS_WITH',
  ends_with: 'ENDS_WITH'
};

var PARSED_EXPRESSION_VERSION = 'v0.expression.v1';

var SUPPORTED_EXPRESSION_FIELDS = [
  'page_url',
  'event_name',
  'offer_detail'
];

var SUPPORTED_RULE_NODE_TYPES = [
  'condition',
  'and',
  'or',
  'refine'
];

var PARSED_EXPRESSION_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['version', 'segmentType', 'include', 'exclude', 'warnings'],
  properties: {
    version: {
      type: 'string',
      enum: [PARSED_EXPRESSION_VERSION]
    },
    segmentType: {
      type: 'string',
      enum: SUPPORTED_SEGMENT_TYPES
    },
    include: {
      type: 'object',
      description: 'A rule node. Node types: condition, and, or, refine.'
    },
    exclude: {
      type: 'array',
      items: {
        type: 'object',
        description: 'A rule node for users to exclude from the audience.'
      }
    },
    warnings: {
      type: 'array',
      items: {
        type: 'string'
      }
    }
  }
};

var STATUS_VALUES = {
  CREATED: 'created',
  EXISTS: 'already exists',
  ERROR_PREFIX: 'error: '
};

var RUN_MODES = {
  GENERATE_VARIANTS: 'generate_variants',
  EXECUTE: 'execute'
};

function main() {
  logSetupWarnings_();

  if (RUN_MODE === RUN_MODES.GENERATE_VARIANTS) {
    generateVariants();
    return;
  }

  if (RUN_MODE === RUN_MODES.EXECUTE) {
    executeSelectedVariants_();
    return;
  }

  throw new Error('Unsupported RUN_MODE "' + RUN_MODE + '". Use generate_variants or execute.');
}

function logSetupWarnings_() {
  Logger.log('Google Ads Audience Builder - V1 public release.');
  Logger.log('Built by Ewald van Kampen (Results Driven Marketing) - https://resultsdriven.nl');
  Logger.log('SCRIPT_VERSION is currently set to: ' + SCRIPT_VERSION);
  Logger.log('RUN_MODE is currently set to: ' + RUN_MODE);
  Logger.log('generate_variants writes rows to VARIANTS. execute processes reviewed rows.');
  Logger.log('If SPREADSHEET_URL is empty or the Sheet is missing template tabs, the script creates them.');
  Logger.log('Optional: configure GEMINI_API_KEY to enable AI parsing for advanced expressions.');
  Logger.log('Never paste your Gemini API key into a shared or public Google Sheet.');
  Logger.log('DRY_RUN is currently set to: ' + DRY_RUN);
}

function generateVariants() {
  var spreadsheet = getConfiguredSpreadsheet_();
  ensureDefaultTemplate_(spreadsheet);
  var audiencesSheet = getRequiredSheet_(spreadsheet, SHEET_NAMES.AUDIENCES);
  var variantsSheet = getRequiredSheet_(spreadsheet, SHEET_NAMES.VARIANTS);
  var durationConfig = variantsSheet
    .getRange(VARIANT_DURATION_CONFIG.ROW, VARIANT_DURATION_CONFIG.COLUMN)
    .getValue();
  var durations = parseMembershipDurations_(durationConfig, DEFAULT_MEMBERSHIP_DURATIONS);
  var audienceDefinitions = readAudienceDefinitions_(audiencesSheet);
  var variants = generateVariantRows_(audienceDefinitions, durations);

  writeVariantRows_(variantsSheet, variants);

  Logger.log(
    'Generated ' + variants.length + ' variants from ' +
    audienceDefinitions.length + ' audience definitions and ' +
    durations.length + ' durations.'
  );
  Logger.log('Variant generation complete. Review VARIANTS, then switch RUN_MODE to execute when ready.');

  return variants;
}

function executeSelectedVariants_() {
  var spreadsheet = getConfiguredSpreadsheet_();
  ensureDefaultTemplate_(spreadsheet);
  var audiencesSheet = getRequiredSheet_(spreadsheet, SHEET_NAMES.AUDIENCES);
  var variantsSheet = getRequiredSheet_(spreadsheet, SHEET_NAMES.VARIANTS);
  var audienceDefinitions = readAudienceDefinitions_(audiencesSheet);
  var audienceBySourceRow = mapAudienceDefinitionsBySourceRow_(audienceDefinitions);
  var variants = readSelectedVariantRows_(variantsSheet);
  var existingNames = findExistingUserLists_();
  var summary = {
    created: 0,
    exists: 0,
    dryRun: 0,
    errors: 0
  };

  Logger.log('Selected variants for execution: ' + variants.length);

  for (var i = 0; i < variants.length; i++) {
    var result = executeVariantRow_(
      variants[i],
      audienceBySourceRow,
      existingNames,
      variantsSheet,
      {}
    );
    summary[result] = summary[result] + 1;
  }

  Logger.log(
    'Execution complete. created=' + summary.created +
    ', exists=' + summary.exists +
    ', dryRun=' + summary.dryRun +
    ', errors=' + summary.errors
  );

  return summary;
}

function getConfiguredSpreadsheet_() {
  if (isPlaceholderValue_(SPREADSHEET_URL)) {
    return createDefaultSpreadsheet_();
  }

  return SpreadsheetApp.openByUrl(SPREADSHEET_URL);
}

function createDefaultSpreadsheet_() {
  var spreadsheet = SpreadsheetApp.create(DEFAULT_SPREADSHEET_NAME);

  setupDefaultTemplate_(spreadsheet);

  Logger.log('Created default Google Ads Audience Builder template: ' + spreadsheet.getUrl());
  Logger.log('Paste this URL into SPREADSHEET_URL before the next run: ' + spreadsheet.getUrl());

  return spreadsheet;
}

function setupDefaultTemplate_(spreadsheet) {
  var howToSheet = spreadsheet.getSheets()[0];
  howToSheet.setName(SHEET_NAMES.HOW_TO_USE);
  clearAndWriteSheet_(howToSheet, HOW_TO_USE_CONTENT);

  var audiencesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.AUDIENCES);
  clearAndWriteSheet_(audiencesSheet, [getAudienceHeaders_()].concat(DEFAULT_AUDIENCE_ROWS));

  var variantsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.VARIANTS);
  variantsSheet.clear();
  variantsSheet
    .getRange(VARIANT_DURATION_CONFIG.ROW, VARIANT_DURATION_CONFIG.COLUMN)
    .setValue(DEFAULT_MEMBERSHIP_DURATIONS.join(','));
  variantsSheet
    .getRange(VARIANT_HEADER_ROW, 1, 1, VARIANT_HEADERS.length)
    .setValues([VARIANT_HEADERS]);

  applyDefaultTemplateValidation_(audiencesSheet, variantsSheet);
  applyDefaultTemplateStyles_(howToSheet, audiencesSheet, variantsSheet);
  removeDefaultBlankSheet_(spreadsheet);

  return spreadsheet;
}

function ensureDefaultTemplate_(spreadsheet) {
  if (hasDefaultTemplateSheets_(spreadsheet)) {
    ensureDefaultTemplateContent_(spreadsheet);
    return spreadsheet;
  }

  if (isBlankSpreadsheet_(spreadsheet)) {
    Logger.log('Configured spreadsheet is blank. Creating default template tabs.');
    return setupDefaultTemplate_(spreadsheet);
  }

  Logger.log('Configured spreadsheet is missing one or more template tabs. Creating missing tabs only.');
  ensureDefaultTemplateContent_(spreadsheet);

  return spreadsheet;
}

function hasDefaultTemplateSheets_(spreadsheet) {
  if (!spreadsheet.getSheetByName(SHEET_NAMES.HOW_TO_USE)) {
    return false;
  }
  if (!spreadsheet.getSheetByName(SHEET_NAMES.AUDIENCES)) {
    return false;
  }
  if (!spreadsheet.getSheetByName(SHEET_NAMES.VARIANTS)) {
    return false;
  }
  return true;
}

function ensureDefaultTemplateContent_(spreadsheet) {
  var howToSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.HOW_TO_USE);
  clearAndWriteSheet_(howToSheet, HOW_TO_USE_CONTENT);

  var audiencesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.AUDIENCES);

  if (isSheetBlank_(audiencesSheet)) {
    clearAndWriteSheet_(audiencesSheet, [getAudienceHeaders_()].concat(DEFAULT_AUDIENCE_ROWS));
  }

  var variantsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.VARIANTS);

  if (isSheetBlank_(variantsSheet)) {
    variantsSheet
      .getRange(VARIANT_DURATION_CONFIG.ROW, VARIANT_DURATION_CONFIG.COLUMN)
      .setValue(DEFAULT_MEMBERSHIP_DURATIONS.join(','));
    variantsSheet
      .getRange(VARIANT_HEADER_ROW, 1, 1, VARIANT_HEADERS.length)
      .setValues([VARIANT_HEADERS]);
  }

  applyDefaultTemplateValidation_(audiencesSheet, variantsSheet);
  applyDefaultTemplateStyles_(howToSheet, audiencesSheet, variantsSheet);
  removeDefaultBlankSheet_(spreadsheet);
}


function isBlankSpreadsheet_(spreadsheet) {
  var sheets = spreadsheet.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    if (!isSheetBlank_(sheets[i])) {
      return false;
    }
  }

  return true;
}

function isSheetBlank_(sheet) {
  return sheet.getLastRow() === 0 || (
    sheet.getLastRow() === 1 &&
    sheet.getLastColumn() === 1 &&
    !trimCell_(sheet.getRange(1, 1).getValue())
  );
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function clearAndWriteSheet_(sheet, rows) {
  sheet.clear();

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function removeDefaultBlankSheet_(spreadsheet) {
  var sheets = spreadsheet.getSheets();

  if (sheets.length <= 1) {
    return;
  }

  for (var i = sheets.length - 1; i >= 0; i--) {
    var sheet = sheets[i];
    var name = sheet.getName();

    if ((name === 'Sheet1' || name === 'Blad1') && isSheetBlank_(sheet)) {
      spreadsheet.deleteSheet(sheet);
    }
  }
}

function getAudienceHeaders_() {
  return [
    'name_pattern',
    'segment_type',
    'expression',
    'auto_add',
    'include_tags',
    'exclude_tags',
    'customer_types',
    'note'
  ];
}



function applyDefaultTemplateValidation_(audiencesSheet, variantsSheet) {
  var segmentTypeRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(SUPPORTED_SEGMENT_TYPES, true)
    .setAllowInvalid(false)
    .build();
  var yesNoRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList([YES_VALUE, NO_VALUE], true)
    .setAllowInvalid(false)
    .build();
  var customerTypesRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(SUPPORTED_CUSTOMER_TYPES, true)
    .setAllowInvalid(false)
    .build();

  audiencesSheet
    .getRange(AUDIENCE_DATA_START_ROW, AUDIENCE_COLUMNS.SEGMENT_TYPE, DEFAULT_TEMPLATE_VALIDATION_ROWS, 1)
    .setDataValidation(segmentTypeRule);
  audiencesSheet
    .getRange(AUDIENCE_DATA_START_ROW, AUDIENCE_COLUMNS.AUTO_ADD, DEFAULT_TEMPLATE_VALIDATION_ROWS, 1)
    .setDataValidation(yesNoRule);
  audiencesSheet
    .getRange(AUDIENCE_DATA_START_ROW, AUDIENCE_COLUMNS.CUSTOMER_TYPES, DEFAULT_TEMPLATE_VALIDATION_ROWS, 1)
    .setDataValidation(customerTypesRule);
  variantsSheet
    .getRange(VARIANT_DATA_START_ROW, VARIANT_COLUMNS.CREATE, DEFAULT_TEMPLATE_VALIDATION_ROWS, 1)
    .setDataValidation(yesNoRule);
}

function applyDefaultTemplateStyles_(howToSheet, audiencesSheet, variantsSheet) {
  styleHeaderRange_(howToSheet.getRange(1, 1, 1, 1));
  styleHowToUseBody_(howToSheet, HOW_TO_USE_CONTENT.length);
  styleHowToUseSectionHeaders_(howToSheet, HOW_TO_USE_CONTENT);
  styleHeaderRange_(audiencesSheet.getRange(AUDIENCE_HEADER_ROW, 1, 1, getAudienceHeaders_().length));
  applyAudienceHeaderNotes_(audiencesSheet);
  styleHeaderRange_(variantsSheet.getRange(VARIANT_HEADER_ROW, 1, 1, VARIANT_HEADERS.length));
  
  applyDefaultColumnWidths_(howToSheet, audiencesSheet, variantsSheet);
  applyVariantStatusFormatting_(variantsSheet);
}

function styleHowToUseBody_(howToSheet, rowCount) {
  howToSheet
    .getRange(1, 1, rowCount, 1)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(10)
    .setBackground('#fbfbfb');
}

function styleHowToUseSectionHeaders_(howToSheet, rows) {
  for (var i = 0; i < rows.length; i++) {
    if (isHowToUseSectionHeader_(rows[i][0])) {
      styleHeaderRange_(howToSheet.getRange(i + 1, 1, 1, 1));
    }
  }
}

function isHowToUseSectionHeader_(value) {
  var sectionHeaders = {
    Workflow: true,
    auto_add: true,
    'Expression examples: basic': true,
    'Expression examples: conditionals': true,
    'Expression examples: advanced': true,
    'Status colors': true,
    Security: true
  };

  return Boolean(sectionHeaders[trimCell_(value)]);
}

function styleHeaderRange_(range) {
  range
    .setBackground(getHeaderBackgroundColor_())
    .setFontWeight('bold')
    .setFontColor('#202124');
}

function getHeaderBackgroundColor_() {
  return HEADER_BACKGROUND_COLOR;
}

function applyDefaultColumnWidths_(howToSheet, audiencesSheet, variantsSheet) {
  howToSheet.setColumnWidth(1, 980);

  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.NAME_PATTERN, 190);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.SEGMENT_TYPE, 150);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.EXPRESSION, 380);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.AUTO_ADD, 100);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.INCLUDE_TAGS, 260);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.EXCLUDE_TAGS, 260);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.CUSTOMER_TYPES, 190);
  audiencesSheet.setColumnWidth(AUDIENCE_COLUMNS.NOTE, 220);

  variantsSheet.setColumnWidth(VARIANT_COLUMNS.ROW_NUMBER, 48);
  variantsSheet.setColumnWidth(VARIANT_COLUMNS.AUDIENCE_NAME, 190);
  variantsSheet.setColumnWidth(VARIANT_COLUMNS.BASED_ON, 95);
  variantsSheet.setColumnWidth(VARIANT_COLUMNS.MEMBERSHIP_DURATION, 180);
  variantsSheet.setColumnWidth(VARIANT_COLUMNS.CREATE, 90);
  variantsSheet.setColumnWidth(VARIANT_COLUMNS.STATUS, 220);
}

function applyAudienceHeaderNotes_(audiencesSheet) {
  var notes = [
    'Audience name pattern with {duration}. Example: Buyers_{duration}d',
    'Source type. Examples: website_visitors, event',
    'For website_visitors only. Examples: URL contains /pricing, event = view_item',
    'Prepopulate historical users. Values: yes, no',
    'For event only. Exact Conversion Action names. Examples: Quote Form Submitted, Phone Call Lead',
    'For event only. Optional comma-separated exclusions. Examples: Paid Subscription',
    'Single customer type. Examples: all_customers, converted_leads',
    'Internal note only. Example: Re-engagement focus'
  ];

  for (var i = 0; i < notes.length; i++) {
    audiencesSheet.getRange(AUDIENCE_HEADER_ROW, i + 1).setNote(notes[i]);
  }
}

function applyVariantStatusFormatting_(variantsSheet) {
  var statusRange = variantsSheet.getRange(
    VARIANT_DATA_START_ROW,
    VARIANT_COLUMNS.STATUS,
    DEFAULT_TEMPLATE_VALIDATION_ROWS,
    1
  );
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS_VALUES.CREATED)
      .setBackground(STATUS_CREATED_BACKGROUND_COLOR)
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(STATUS_VALUES.EXISTS)
      .setBackground(STATUS_EXISTS_BACKGROUND_COLOR)
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(STATUS_VALUES.ERROR_PREFIX)
      .setBackground(STATUS_ERROR_BACKGROUND_COLOR)
      .setRanges([statusRange])
      .build()
  ];

  variantsSheet.setConditionalFormatRules(rules);
}



function getRequiredSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Missing required sheet: ' + sheetName + '.');
  }

  return sheet;
}

function parseMembershipDurations_(rawValue, fallbackDurations) {
  var sourceValues = [];
  var rawText = trimCell_(rawValue);

  if (rawText) {
    sourceValues = rawText.split(',');
  } else {
    sourceValues = fallbackDurations || DEFAULT_MEMBERSHIP_DURATIONS;
  }

  var durations = [];
  var seen = {};

  for (var i = 0; i < sourceValues.length; i++) {
    var duration = validateMembershipDuration_(sourceValues[i]);

    if (!seen[duration]) {
      durations.push(duration);
      seen[duration] = true;
    }
  }

  if (durations.length === 0) {
    throw new Error('Add at least one membership duration in VARIANTS!' + VARIANT_DURATION_CONFIG.A1 + '.');
  }

  return durations;
}

function validateMembershipDuration_(value) {
  var text = trimCell_(value);

  if (!/^\d+$/.test(text)) {
    throw new Error('Invalid membership duration "' + value + '". Use whole days from 1 through 540.');
  }

  var duration = Number(text);

  if (duration < MIN_MEMBERSHIP_DURATION || duration > MAX_MEMBERSHIP_DURATION) {
    throw new Error('Invalid membership duration "' + value + '". Use whole days from 1 through 540.');
  }

  return duration;
}

function readAudienceDefinitions_(audiencesSheet) {
  var lastRow = audiencesSheet.getLastRow();

  if (lastRow < AUDIENCE_DATA_START_ROW) {
    return [];
  }

  var rowCount = lastRow - AUDIENCE_DATA_START_ROW + 1;
  var readWidth = getAudienceHeaders_().length;
  var rows = audiencesSheet
    .getRange(AUDIENCE_DATA_START_ROW, 1, rowCount, readWidth)
    .getValues();
  var definitions = [];

  for (var i = 0; i < rows.length; i++) {
    var sheetRowNumber = AUDIENCE_DATA_START_ROW + i;

    try {
      var definition = normalizeAudienceDefinition_(rows[i], sheetRowNumber);

      if (definition) {
        definitions.push(definition);
      }
    } catch (error) {
      Logger.log('Skipping AUDIENCES row ' + sheetRowNumber + ': ' + error.message);
    }
  }

  return definitions;
}

function normalizeAudienceDefinition_(rowValues, sheetRowNumber) {
  if (isBlankRow_(rowValues)) {
    return null;
  }

  var values = rowValues;
  var sourceRowNumber = '';
  var namePattern = normalizeNamePattern_(values[AUDIENCE_COLUMNS.NAME_PATTERN - 1]);
  var segmentType = String(values[AUDIENCE_COLUMNS.SEGMENT_TYPE - 1]).toLowerCase().trim();
  var expression = translateExpressionToEnglish_(values[AUDIENCE_COLUMNS.EXPRESSION - 1]);
  var autoAdd = String(values[AUDIENCE_COLUMNS.AUTO_ADD - 1]).toLowerCase().trim() || YES_VALUE;
  var includeTags = parseTagList_(values[AUDIENCE_COLUMNS.INCLUDE_TAGS - 1]);
  var excludeTags = parseTagList_(values[AUDIENCE_COLUMNS.EXCLUDE_TAGS - 1]);
  var customerType = normalizeCustomerType_(values[AUDIENCE_COLUMNS.CUSTOMER_TYPES - 1]);
  var note = trimCell_(values[AUDIENCE_COLUMNS.NOTE - 1]);
  sourceRowNumber = sourceRowNumber || String(sheetRowNumber - AUDIENCE_HEADER_ROW);

  if (!namePattern) {
    throw new Error('name_pattern is required.');
  }

  if (!segmentType) {
    throw new Error('segment_type is required.');
  }

  if (SUPPORTED_SEGMENT_TYPES.indexOf(segmentType) === -1) {
    throw new Error('Unsupported segment_type "' + segmentType + '".');
  }

  if (segmentType === 'event') {
    if (includeTags.length === 0) {
      throw new Error('include_tags is required when segment_type is "event".');
    }
    if (expression) {
      throw new Error('expression must be empty when segment_type is "event"; use include_tags/exclude_tags.');
    }
  }

  if (segmentType === 'website_visitors') {
    if (!expression) {
      throw new Error('expression is required when segment_type is "website_visitors".');
    }
    if (includeTags.length > 0 || excludeTags.length > 0) {
      throw new Error('include_tags/exclude_tags are only allowed when segment_type is "event".');
    }
  }

  if ([YES_VALUE, NO_VALUE].indexOf(autoAdd) === -1) {
    throw new Error('auto_add must be "' + YES_VALUE + '" or "' + NO_VALUE + '".');
  }

  return {
    sourceRowNumber: sourceRowNumber,
    sheetRowNumber: sheetRowNumber,
    namePattern: namePattern,
    segmentType: segmentType,
    expression: expression,
    autoAdd: autoAdd,
    includeTags: includeTags,
    excludeTags: excludeTags,
    customerType: customerType,
    note: note
  };
}


function parseTagList_(rawValue) {
  var text = trimCell_(rawValue);

  if (!text) {
    return [];
  }

  var parts = text.split(',');
  var values = [];
  var seen = {};

  for (var i = 0; i < parts.length; i++) {
    var value = trimCell_(parts[i]);

    if (!value || seen[value]) {
      continue;
    }

    values.push(value);
    seen[value] = true;
  }

  return values;
}

function normalizeCustomerType_(value) {
  var normalized = value;

  if (SUPPORTED_CUSTOMER_TYPES.indexOf(normalized) === -1) {
    throw new Error('Unsupported customer_types "' + value + '".');
  }

  return normalized;
}

function generateVariantRows_(audienceDefinitions, durations) {
  var variants = [];
  var rowNumber = 1;

  for (var i = 0; i < audienceDefinitions.length; i++) {
    for (var j = 0; j < durations.length; j++) {
      var definition = audienceDefinitions[i];
      var duration = durations[j];

      variants.push({
        rowNumber: rowNumber,
        audienceName: buildVariantName_(definition.namePattern, duration),
        basedOn: definition.sourceRowNumber,
        membershipDuration: duration,
        create: YES_VALUE,
        status: '',
        segmentType: definition.segmentType,
        expression: definition.expression,
        autoAdd: definition.autoAdd,
        note: definition.note
      });

      rowNumber++;
    }
  }

  return variants;
}

function buildVariantName_(namePattern, duration) {
  return normalizeNamePattern_(namePattern).split('{duration}').join(String(duration));
}

function toVariantSheetRows_(variants) {
  var rows = [];

  for (var i = 0; i < variants.length; i++) {
    var variant = variants[i];

    rows.push([
      variant.rowNumber,
      variant.audienceName,
      variant.basedOn,
      variant.membershipDuration,
      variant.create,
      variant.status
    ]);
  }

  return rows;
}

function writeVariantRows_(variantsSheet, variants) {
  variantsSheet
    .getRange(VARIANT_HEADER_ROW, 1, 1, VARIANT_HEADERS.length)
    .setValues([VARIANT_HEADERS]);
  styleHeaderRange_(variantsSheet.getRange(VARIANT_HEADER_ROW, 1, 1, VARIANT_HEADERS.length));

  clearExistingVariantRows_(variantsSheet);

  var rows = toVariantSheetRows_(variants);

  if (rows.length > 0) {
    variantsSheet
      .getRange(VARIANT_DATA_START_ROW, 1, rows.length, VARIANT_HEADERS.length)
      .setValues(rows);
  }
}

function clearExistingVariantRows_(variantsSheet) {
  var lastRow = variantsSheet.getLastRow();

  if (lastRow < VARIANT_DATA_START_ROW) {
    return;
  }

  variantsSheet
    .getRange(VARIANT_DATA_START_ROW, 1, lastRow - VARIANT_DATA_START_ROW + 1, VARIANT_HEADERS.length)
    .clearContent();
}

function readSelectedVariantRows_(variantsSheet) {
  var lastRow = variantsSheet.getLastRow();

  if (lastRow < VARIANT_DATA_START_ROW) {
    return [];
  }

  var rowCount = lastRow - VARIANT_DATA_START_ROW + 1;
  var rows = variantsSheet
    .getRange(VARIANT_DATA_START_ROW, 1, rowCount, VARIANT_HEADERS.length)
    .getValues();
  var variants = [];

  for (var i = 0; i < rows.length; i++) {
    var sheetRowNumber = VARIANT_DATA_START_ROW + i;
    var variant = normalizeVariantRow_(rows[i], sheetRowNumber);

    if (variant && variant.create === YES_VALUE) {
      variants.push(variant);
    }
  }

  return variants;
}

function normalizeVariantRow_(rowValues, sheetRowNumber) {
  if (isBlankRow_(rowValues)) {
    return null;
  }

  var createValue = String(rowValues[VARIANT_COLUMNS.CREATE - 1]).toLowerCase().trim() || NO_VALUE;

  if ([YES_VALUE, NO_VALUE].indexOf(createValue) === -1) {
    throw new Error('VARIANTS row ' + sheetRowNumber + ': create must be "yes" or "no".');
  }

  var membershipDuration = validateMembershipDuration_(rowValues[VARIANT_COLUMNS.MEMBERSHIP_DURATION - 1]);
  var audienceName = trimCell_(rowValues[VARIANT_COLUMNS.AUDIENCE_NAME - 1]);
  var basedOn = trimCell_(rowValues[VARIANT_COLUMNS.BASED_ON - 1]);

  if (!audienceName) {
    throw new Error('VARIANTS row ' + sheetRowNumber + ': audience_name is required.');
  }

  if (!basedOn) {
    throw new Error('VARIANTS row ' + sheetRowNumber + ': based_on is required.');
  }

  return {
    rowNumber: trimCell_(rowValues[VARIANT_COLUMNS.ROW_NUMBER - 1]),
    sheetRowNumber: sheetRowNumber,
    audienceName: audienceName,
    basedOn: basedOn,
    membershipDuration: membershipDuration,
    create: createValue,
    status: normalizeStatusValue_(rowValues[VARIANT_COLUMNS.STATUS - 1])
  };
}

function mapAudienceDefinitionsBySourceRow_(audienceDefinitions) {
  var bySourceRow = {};

  for (var i = 0; i < audienceDefinitions.length; i++) {
    bySourceRow[audienceDefinitions[i].sourceRowNumber] = audienceDefinitions[i];
  }

  return bySourceRow;
}

function isBlankRow_(rowValues) {
  for (var i = 0; i < rowValues.length; i++) {
    if (trimCell_(rowValues[i])) {
      return false;
    }
  }

  return true;
}

function trimCell_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function normalizeNamePattern_(value) {
  return trimCell_(value).split('{duur}').join('{duration}');
}



function normalizeStatusValue_(value) {
  var text = trimCell_(value);

  if (text.indexOf('fout:') === 0) {
    return STATUS_VALUES.ERROR_PREFIX + text.substring('fout:'.length).trim();
  }

  return text;
}

function translateExpressionToEnglish_(value) {
  var expression = trimCell_(value);

  return expression
    .replace(/\bURL bevat\b/g, 'URL contains')
    .replace(/\bURL begint_met\b/g, 'URL starts_with')
    .replace(/\bURL eindigt_met\b/g, 'URL ends_with')
    .replace(/\bevenement\b/g, 'event')
    .replace(/\bEN\b/g, 'AND')
    .replace(/\bOF\b/g, 'OR')
    .replace(/\bNIET\b/g, 'NOT')
    .replace(/\bIN ([0-9]+) DAGEN\b/g, 'IN $1 DAYS')
    .replace(/\bVERFIJN\b/g, 'REFINE')
    .replace(/\bbevat\b/g, 'contains')
    .replace(/\bis_gelijk_aan\b/g, 'equals')
    .replace(/\bbegint_met\b/g, 'starts_with')
    .replace(/\beindigt_met\b/g, 'ends_with')
    .replace(/\bofferdetail\b/g, 'offer_detail')
    .replace(/\/bedankt/g, '/thank-you')
    .replace(/\/retour/g, '/return');
}

function isPlaceholderValue_(value) {
  var text = trimCell_(value);
  return !text || text.indexOf('TODO:') === 0;
}

function parseExpression_(expression, segmentType) {
  var normalizedExpression = translateExpressionToEnglish_(expression);
  var simpleRule = parseSimpleExpression_(normalizedExpression, segmentType);

  if (simpleRule) {
    return validateParsedRule_(simpleRule, segmentType);
  }

  return parseExpressionWithGemini_(normalizedExpression, segmentType);
}

function parseSimpleExpression_(expression, segmentType) {
  var normalizedSegmentType = segmentType;
  var parts = splitExpressionOnce_(expression, ' NOT ');
  var includeNode = parseSimpleLogicalExpression_(parts[0]);

  if (!includeNode) {
    return null;
  }

  var excludeNodes = [];

  if (parts.length > 1) {
    var excludeNode = parseSimpleLogicalExpression_(parts[1]);

    if (!excludeNode) {
      return null;
    }

    excludeNodes.push(excludeNode);
  }

  return {
    version: PARSED_EXPRESSION_VERSION,
    segmentType: normalizedSegmentType,
    include: includeNode,
    exclude: excludeNodes,
    warnings: []
  };
}

function parseSimpleLogicalExpression_(expression) {
  var normalizedExpression = stripWrappingParentheses_(trimCell_(expression));
  var orParts = splitExpressionList_(normalizedExpression, ' OR ');

  if (orParts.length > 1) {
    var orConditions = parseSimpleConditionList_(orParts);

    if (!orConditions) {
      return null;
    }

    return {
      type: 'or',
      conditions: orConditions
    };
  }

  var andParts = splitExpressionList_(normalizedExpression, ' AND ');

  if (andParts.length > 1) {
    var andConditions = parseSimpleConditionList_(andParts);

    if (!andConditions) {
      return null;
    }

    return {
      type: 'and',
      conditions: andConditions
    };
  }

  return parseSimpleCondition_(normalizedExpression);
}

function parseSimpleConditionList_(parts) {
  var conditions = [];

  for (var i = 0; i < parts.length; i++) {
    var condition = parseSimpleLogicalExpression_(parts[i]);

    if (!condition) {
      return null;
    }

    conditions.push(condition);
  }

  return conditions;
}

function parseSimpleCondition_(expression) {
  return parseSimpleUrlCondition_(expression) ||
    parseSimpleEventCondition_(expression) ||
    parseSimpleFieldCondition_(expression);
}

function parseSimpleUrlCondition_(expression) {
  var match = /^URL\s+(contains|equals|starts_with|ends_with)\s+(.+?)(?:\s+IN\s+([0-9]+)\s+DAYS)?$/i.exec(expression);

  if (!match) {
    return null;
  }

  return buildSimpleCondition_('page_url', match[1], match[2], match[3]);
}

function parseSimpleEventCondition_(expression) {
  var match = /^event\s*(=|contains|equals|starts_with|ends_with)\s+(.+?)(?:\s+IN\s+([0-9]+)\s+DAYS)?$/i.exec(expression);

  if (!match) {
    return null;
  }

  return buildSimpleCondition_('event_name', match[1] === '=' ? 'equals' : match[1], match[2], match[3]);
}

function parseSimpleFieldCondition_(expression) {
  var match = /^(page_url|event_name|offer_detail)\s+(contains|equals|starts_with|ends_with)\s+(.+?)(?:\s+IN\s+([0-9]+)\s+DAYS)?$/i.exec(expression);

  if (!match) {
    return null;
  }

  return buildSimpleCondition_(match[1], match[2], match[3], match[4]);
}

function buildSimpleCondition_(field, operator, value, lookbackDays) {
  var condition = {
    type: 'condition',
    field: String(field).toLowerCase(),
    operator: String(operator).toLowerCase(),
    value: stripQuotes_(value)
  };

  if (lookbackDays) {
    condition.lookbackDays = lookbackDays;
  }

  return condition;
}

function splitExpressionOnce_(expression, separator) {
  var index = expression.indexOf(separator);

  if (index === -1) {
    return [expression];
  }

  return [
    expression.substring(0, index),
    expression.substring(index + separator.length)
  ];
}

function splitExpressionList_(expression, separator) {
  var parts = expression.split(separator);

  if (parts.length < 2) {
    return [expression];
  }

  return parts;
}

function stripWrappingParentheses_(value) {
  var text = trimCell_(value);

  if (text.charAt(0) === '(' && text.charAt(text.length - 1) === ')') {
    return trimCell_(text.substring(1, text.length - 1));
  }

  return text;
}

function stripQuotes_(value) {
  var text = trimCell_(value);
  var first = text.charAt(0);
  var last = text.charAt(text.length - 1);

  if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
    return text.substring(1, text.length - 1);
  }

  return text;
}

function parseExpressionWithGemini_(expression, segmentType) {
  if (isPlaceholderValue_(GEMINI_API_KEY)) {
    throw new Error('Configure GEMINI_API_KEY before parsing expressions.');
  }

  var request = buildGeminiExpressionRequest_(expression, segmentType);
  var response = UrlFetchApp.fetch(request.url, request.params);
  var statusCode = response.getResponseCode();
  var responseBody = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Gemini API request failed with HTTP ' + statusCode + ': ' + responseBody);
  }

  var jsonText = extractGeminiText_(responseBody);
  var parsedRule = parseGeminiJsonText_(jsonText);

  return validateParsedRule_(parsedRule, segmentType);
}

function buildGeminiExpressionPrompt_(expression, segmentType) {
  if (SUPPORTED_SEGMENT_TYPES.indexOf(segmentType) === -1) {
    throw new Error('Unsupported segment_type for expression parsing: ' + segmentType + '.');
  }

  return [
    'You parse Google Ads audience expressions into constrained JSON.',
    'Return JSON only. Do not return JavaScript, Markdown, comments, or explanations.',
    'The JSON must match version "' + PARSED_EXPRESSION_VERSION + '".',
    'segmentType must be "' + segmentType + '".',
    'Supported fields: ' + SUPPORTED_EXPRESSION_FIELDS.join(', ') + '.',
    'Supported operators: ' + SUPPORTED_OPERATORS.join(', ') + '.',
    'Supported node types: condition, and, or, refine.',
    'Every rule node must include a type property.',
    'Map "URL" to field "page_url".',
    'Map "event" to field "event_name".',
    'Map "=" to operator "equals".',
    'Map "contains", "starts_with", and "ends_with" directly.',
    'Use lookbackDays only when the expression says "IN n DAYS".',
    'Use include for positive logic.',
    'Use exclude for logic after "NOT".',
    'Use type "and" for AND and type "or" for OR.',
    'Use type "refine" for REFINE with base and refinement objects.',
    '',
    'Rule node shapes:',
    '- condition: {type, field, operator, value, lookbackDays}',
    '- and/or: {type, conditions}',
    '- refine: {type, base, refinement}',
    '',
    'Required top-level shape:',
    JSON.stringify(PARSED_EXPRESSION_SCHEMA),
    '',
    'Expression:',
    expression
  ].join('\n');
}

function buildGeminiExpressionRequest_(expression, segmentType) {
  var prompt = buildGeminiExpressionPrompt_(expression, segmentType);
  var payload = {
    contents: [
      {
        role: 'user',
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseJsonSchema: PARSED_EXPRESSION_SCHEMA
    }
  };

  return {
    url: GEMINI_API_BASE_URL + GEMINI_MODEL + ':generateContent',
    params: {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-goog-api-key': GEMINI_API_KEY
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    }
  };
}

function extractGeminiText_(responseBody) {
  var parsedBody = typeof responseBody === 'string' ? JSON.parse(responseBody) : responseBody;
  var candidates = parsedBody.candidates || [];

  if (
    candidates.length === 0 ||
    !candidates[0].content ||
    !candidates[0].content.parts ||
    candidates[0].content.parts.length === 0 ||
    !candidates[0].content.parts[0].text
  ) {
    throw new Error('Gemini response did not include text output.');
  }

  return candidates[0].content.parts[0].text;
}

function parseGeminiJsonText_(jsonText) {
  if (typeof jsonText === 'object') {
    return jsonText;
  }

  try {
    return JSON.parse(trimCell_(jsonText));
  } catch (error) {
    throw new Error('Gemini returned invalid JSON: ' + error.message);
  }
}

function validateParsedRule_(parsedRule, expectedSegmentType) {
  var rule = typeof parsedRule === 'string' ? parseGeminiJsonText_(parsedRule) : parsedRule;
  var normalizedExpectedSegmentType = expectedSegmentType;

  assertObject_(rule, 'parsedRule');
  normalizeParsedRuleShape_(rule);
  rule.segmentType = rule.segmentType;
  assertEqual_(rule.version, PARSED_EXPRESSION_VERSION, 'parsedRule.version');
  assertSupportedValue_(rule.segmentType, SUPPORTED_SEGMENT_TYPES, 'parsedRule.segmentType');

  if (normalizedExpectedSegmentType && rule.segmentType !== normalizedExpectedSegmentType) {
    throw new Error(
      'parsedRule.segmentType must match source segment_type "' + normalizedExpectedSegmentType + '".'
    );
  }

  validateRuleNode_(rule.include, 'include');
  assertArray_(rule.exclude, 'exclude');

  for (var i = 0; i < rule.exclude.length; i++) {
    validateRuleNode_(rule.exclude[i], 'exclude[' + i + ']');
  }

  assertArray_(rule.warnings, 'warnings');

  for (var j = 0; j < rule.warnings.length; j++) {
    if (typeof rule.warnings[j] !== 'string') {
      throw new Error('warnings[' + j + '] must be a string.');
    }
  }

  return rule;
}

function normalizeParsedRuleShape_(rule) {
  rule.include = normalizeRuleNodeShape_(rule.include);

  if (!rule.exclude) {
    rule.exclude = [];
  } else if (Object.prototype.toString.call(rule.exclude) !== '[object Array]') {
    rule.exclude = [rule.exclude];
  }

  for (var i = 0; i < rule.exclude.length; i++) {
    rule.exclude[i] = normalizeRuleNodeShape_(rule.exclude[i]);
  }

  if (!rule.warnings) {
    rule.warnings = [];
  }
}

function normalizeRuleNodeShape_(node) {
  if (Object.prototype.toString.call(node) === '[object Array]') {
    if (node.length === 1) {
      return normalizeRuleNodeShape_(node[0]);
    }

    return {
      type: 'and',
      conditions: normalizeRuleNodeListShape_(node)
    };
  }

  if (!node || typeof node !== 'object') {
    return node;
  }

  if (!node.type) {
    if (node.conditions) {
      node.type = 'and';
    } else if (node.base && node.refinement) {
      node.type = 'refine';
    } else {
      node.type = 'condition';
    }
  }

  if (node.conditions) {
    node.conditions = normalizeRuleNodeListShape_(node.conditions);
  }

  if (node.base) {
    node.base = normalizeRuleNodeShape_(node.base);
  }

  if (node.refinement) {
    node.refinement = normalizeRuleNodeShape_(node.refinement);
  }

  return node;
}

function normalizeRuleNodeListShape_(nodes) {
  var normalizedNodes = [];

  for (var i = 0; i < nodes.length; i++) {
    normalizedNodes.push(normalizeRuleNodeShape_(nodes[i]));
  }

  return normalizedNodes;
}

function validateRuleNode_(node, path) {
  assertObject_(node, path);
  assertSupportedValue_(node.type, SUPPORTED_RULE_NODE_TYPES, path + '.type');

  if (node.type === 'condition') {
    validateConditionNode_(node, path);
    return;
  }

  if (node.type === 'and' || node.type === 'or') {
    assertArray_(node.conditions, path + '.conditions');

    if (node.conditions.length < 2) {
      throw new Error(path + '.conditions must contain at least two rule nodes.');
    }

    for (var i = 0; i < node.conditions.length; i++) {
      validateRuleNode_(node.conditions[i], path + '.conditions[' + i + ']');
    }

    return;
  }

  if (node.type === 'refine') {
    validateRuleNode_(node.base, path + '.base');
    assertObject_(node.refinement, path + '.refinement');

    if (!node.refinement.type) {
      node.refinement.type = 'condition';
    }

    assertEqual_(node.refinement.type, 'condition', path + '.refinement.type');
    validateConditionNode_(node.refinement, path + '.refinement');
  }
}

function validateConditionNode_(node, path) {
  node.field = node.field;
  node.operator = node.operator;

  assertSupportedValue_(node.field, SUPPORTED_EXPRESSION_FIELDS, path + '.field');
  assertSupportedValue_(node.operator, SUPPORTED_OPERATORS, path + '.operator');

  if (!trimCell_(node.value)) {
    throw new Error(path + '.value is required.');
  }

  if (hasOwnProperty_(node, 'lookbackDays') && trimCell_(node.lookbackDays)) {
    node.lookbackDays = validateMembershipDuration_(node.lookbackDays);
  }
}

function assertObject_(value, path) {
  if (!value || typeof value !== 'object' || Object.prototype.toString.call(value) === '[object Array]') {
    throw new Error(path + ' must be an object.');
  }
}

function assertArray_(value, path) {
  if (Object.prototype.toString.call(value) !== '[object Array]') {
    throw new Error(path + ' must be an array.');
  }
}

function assertEqual_(actual, expected, path) {
  if (actual !== expected) {
    throw new Error(path + ' must be "' + expected + '".');
  }
}

function assertSupportedValue_(value, supportedValues, path) {
  if (supportedValues.indexOf(value) === -1) {
    throw new Error(path + ' has unsupported value "' + value + '".');
  }
}

function hasOwnProperty_(objectValue, propertyName) {
  return Object.prototype.hasOwnProperty.call(objectValue, propertyName);
}

function findExistingUserLists_() {
  var existingNames = {};
  var iterator = AdsApp.userlists().get();

  while (iterator.hasNext()) {
    var userList = iterator.next();
    existingNames[userList.getName()] = true;
  }

  return existingNames;
}



function executeVariantRow_(variant, audienceBySourceRow, existingNames, variantsSheet, services) {
  var activeServices = services || {};
  var writeStatus = activeServices.writeStatus || writeVariantStatus_;
  var dryRun = hasOwnProperty_(activeServices, 'dryRun') ? activeServices.dryRun : DRY_RUN;
  var parseExpression = activeServices.parseExpression || parseExpression_;
  var existingAudienceRows = activeServices.existingAudienceRows || [];
  var mutate = activeServices.mutate || function (operation) {
    return AdsApp.mutate(operation, { partialFailure: true });
  };

  try {
    if (existingNames[variant.audienceName]) {
      writeStatus(variantsSheet, variant.sheetRowNumber, STATUS_VALUES.EXISTS);
      log_('Row ' + variant.sheetRowNumber + ': already exists - ' + variant.audienceName);
      return 'exists';
    }

    var executionVariant = resolveExecutionVariant_(variant, audienceBySourceRow);
    var parsedRule = buildExecutionParsedRule_(executionVariant, parseExpression, existingAudienceRows);
    var operation = buildUserListMutateOperation_(executionVariant, parsedRule);

    if (dryRun) {
      writeStatus(
        variantsSheet,
        variant.sheetRowNumber,
        formatErrorStatus_('DRY_RUN active - would create audience')
      );
      log_('Row ' + variant.sheetRowNumber + ': DRY_RUN would create - ' + variant.audienceName);
      return 'dryRun';
    }

    var mutateResult = mutate(operation);
    var status = handleMutateResult_(mutateResult);

    writeStatus(variantsSheet, variant.sheetRowNumber, status);

    if (status === STATUS_VALUES.CREATED) {
      existingNames[variant.audienceName] = true;
      log_('Row ' + variant.sheetRowNumber + ': created - ' + variant.audienceName);
      return 'created';
    }

    log_('Row ' + variant.sheetRowNumber + ': error - ' + status);
    return 'errors';
  } catch (error) {
    var errorStatus = formatErrorStatus_(error.message);
    writeStatus(variantsSheet, variant.sheetRowNumber, errorStatus);
    log_('Row ' + variant.sheetRowNumber + ': error - ' + errorStatus);
    return 'errors';
  }
}

function log_(message) {
  if (typeof Logger !== 'undefined' && Logger.log) {
    Logger.log(message);
  }
}

function resolveExecutionVariant_(variant, audienceBySourceRow) {
  var source = audienceBySourceRow[variant.basedOn];

  if (!source) {
    throw new Error('No AUDIENCES source row found for based_on=' + variant.basedOn + '.');
  }

  return {
    rowNumber: variant.rowNumber,
    sheetRowNumber: variant.sheetRowNumber,
    audienceName: variant.audienceName,
    basedOn: variant.basedOn,
    membershipDuration: variant.membershipDuration,
    create: variant.create,
    status: variant.status,
    segmentType: source.segmentType,
    expression: source.expression,
    autoAdd: source.autoAdd,
    includeTags: source.includeTags || [],
    excludeTags: source.excludeTags || [],
    customerType: source.customerType || '',
    note: source.note
  };
}

function buildExecutionParsedRule_(variant, parseExpression, existingAudienceRows) {
  if (variant.segmentType === 'event') {
    return buildEventTagsRule_(variant);
  }

  if (variant.includeTags && variant.includeTags.length > 0) {
    return buildCombinedAudienceRule_(variant, existingAudienceRows);
  }

  return parseExpression(variant.expression, variant.segmentType);
}

function buildEventTagsRule_(variant) {
  var includeConditions = [];
  var excludeConditions = [];
  var i;

  for (i = 0; i < (variant.includeTags || []).length; i++) {
    includeConditions.push({
      type: 'condition',
      field: 'event_name',
      operator: 'equals',
      value: variant.includeTags[i]
    });
  }

  for (i = 0; i < (variant.excludeTags || []).length; i++) {
    excludeConditions.push({
      type: 'condition',
      field: 'event_name',
      operator: 'equals',
      value: variant.excludeTags[i]
    });
  }

  return {
    version: PARSED_EXPRESSION_VERSION,
    segmentType: 'event',
    include: includeConditions.length === 1 ? includeConditions[0] : {
      type: 'or',
      conditions: includeConditions
    },
    exclude: excludeConditions,
    warnings: ['event selector mode used (include_tags/exclude_tags)']
  };
}

function buildCombinedAudienceRule_(variant, existingAudienceRows) {
  var includeRefs = resolveCombinedAudienceReferences_(variant.includeTags || [], existingAudienceRows);
  var excludeRefs = resolveCombinedAudienceReferences_(variant.excludeTags || [], existingAudienceRows);
  var includeConditions = [];
  var excludeConditions = [];
  var i;

  for (i = 0; i < includeRefs.length; i++) {
    includeConditions.push({
      type: 'condition',
      field: 'event_name',
      operator: 'equals',
      value: includeRefs[i].tagValue
    });
  }

  for (i = 0; i < excludeRefs.length; i++) {
    excludeConditions.push({
      type: 'condition',
      field: 'event_name',
      operator: 'equals',
      value: excludeRefs[i].tagValue
    });
  }

  return {
    version: PARSED_EXPRESSION_VERSION,
    segmentType: variant.segmentType || 'event',
    include: includeConditions.length === 1 ? includeConditions[0] : {
      type: 'or',
      conditions: includeConditions
    },
    exclude: excludeConditions,
    warnings: ['combined tags mode used (structured include_tags/exclude_tags)']
  };
}

function resolveCombinedAudienceReferences_(tags, existingAudienceRows) {
  var normalizedRows = [];
  var resolved = [];
  var i;
  var j;

  for (i = 0; i < existingAudienceRows.length; i++) {
    normalizedRows.push(normalizeExistingAudienceSheetRow_(existingAudienceRows[i]));
  }

  for (i = 0; i < tags.length; i++) {
    var tag = trimCell_(tags[i]);
    var aliasMatches = [];
    var resourceMatches = [];
    var nameMatches = [];

    for (j = 0; j < normalizedRows.length; j++) {
      var row = normalizedRows[j];

      if (trimCell_(row[EXISTING_AUDIENCE_COLUMNS.ALIAS - 1]) === tag) {
        aliasMatches.push(row);
      }
      if (trimCell_(row[EXISTING_AUDIENCE_COLUMNS.RESOURCE_NAME - 1]) === tag) {
        resourceMatches.push(row);
      }
      if (trimCell_(row[EXISTING_AUDIENCE_COLUMNS.AUDIENCE_NAME - 1]) === tag) {
        nameMatches.push(row);
      }
    }

    var matches = aliasMatches.length ? aliasMatches : (resourceMatches.length ? resourceMatches : nameMatches);

    if (matches.length === 0) {
      throw new Error('Missing combined reference: "' + tag + '". Import existing audiences first or fix include_tags/exclude_tags.');
    }

    if (matches.length > 1) {
      throw new Error('Reference "' + tag + '" is ambiguous. Use alias or resource_name.');
    }

    resolved.push({
      tagValue: tag,
      resourceName: trimCell_(matches[0][EXISTING_AUDIENCE_COLUMNS.RESOURCE_NAME - 1]),
      audienceName: trimCell_(matches[0][EXISTING_AUDIENCE_COLUMNS.AUDIENCE_NAME - 1])
    });
  }

  return resolved;
}

function handleMutateResult_(mutateResult) {
  if (mutateResult.isSuccessful()) {
    return STATUS_VALUES.CREATED;
  }

  return formatErrorStatus_(mutateResult.getErrorMessages().join('; '));
}

function formatErrorStatus_(message) {
  var status = STATUS_VALUES.ERROR_PREFIX + trimCell_(message);

  if (status.length > 250) {
    return status.substring(0, 247) + '...';
  }

  return status;
}

function buildUserListMutateOperation_(variant, parsedRule) {
  var validatedRule = validateParsedRule_(parsedRule, parsedRule.segmentType);
  var customerTypeEnum = mapCustomerTypeToEnum_(variant.customerType);

  return {
    userListOperation: {
      create: {
        name: variant.audienceName,
        description: buildUserListDescription_(variant, validatedRule, customerTypeEnum),
        membershipStatus: 'OPEN',
        ruleBasedUserList: {
          prepopulationStatus: variant.autoAdd === NO_VALUE ? 'NONE' : 'REQUESTED',
          flexibleRuleUserList: buildFlexibleRuleUserList_(validatedRule, variant.membershipDuration)
        }
      }
    }
  };
}

function buildUserListDescription_(variant, parsedRule) {
  var parts = [
    'Created by Google Ads Audience Builder.',
    'Source row: ' + variant.basedOn + '.',
    'Segment type: ' + parsedRule.segmentType + '.'
  ];
  var customerTypeEnum = mapCustomerTypeToEnum_(variant.customerType);

  if (customerTypeEnum) {
    parts.push('Customer type: ' + customerTypeEnum + '.');
  }

  return parts.join(' ');
}

function mapCustomerTypeToEnum_(customerType) {
  var normalized = customerType;

  if (!normalized) {
    return '';
  }

  if (!CUSTOMER_TYPE_ENUM_MAP[normalized]) {
    throw new Error('Unsupported customer_types "' + customerType + '".');
  }

  return CUSTOMER_TYPE_ENUM_MAP[normalized];
}

function buildFlexibleRuleUserList_(parsedRule, defaultLookbackDays) {
  var include = parsedRule.include;
  var inclusiveRuleOperator = 'OR';
  var inclusiveOperands = [];

  if (include.type === 'and' || include.type === 'or') {
    inclusiveRuleOperator = include.type === 'and' ? 'AND' : 'OR';

    for (var i = 0; i < include.conditions.length; i++) {
      inclusiveOperands.push(ruleNodeToOperand_(include.conditions[i], defaultLookbackDays));
    }
  } else {
    inclusiveOperands.push(ruleNodeToOperand_(include, defaultLookbackDays));
  }

  var exclusiveOperands = [];

  for (var j = 0; j < parsedRule.exclude.length; j++) {
    exclusiveOperands.push(ruleNodeToOperand_(parsedRule.exclude[j], defaultLookbackDays));
  }

  return {
    inclusiveRuleOperator: inclusiveRuleOperator,
    inclusiveOperands: inclusiveOperands,
    exclusiveOperands: exclusiveOperands
  };
}

function ruleNodeToOperand_(node, defaultLookbackDays) {
  var groups = nodeToRuleItemGroups_(node);

  return {
    rule: {
      ruleItemGroups: groups
    },
    lookbackWindowDays: resolveLookbackWindowDays_(node, defaultLookbackDays)
  };
}

function nodeToRuleItemGroups_(node) {
  if (node.type === 'condition') {
    return [
      {
        ruleItems: [conditionToRuleItem_(node)]
      }
    ];
  }

  if (node.type === 'refine') {
    var baseGroups = nodeToRuleItemGroups_(node.base);
    var refinementItem = conditionToRuleItem_(node.refinement);

    for (var i = 0; i < baseGroups.length; i++) {
      baseGroups[i].ruleItems.push(refinementItem);
    }

    return baseGroups;
  }

  if (node.type === 'or') {
    var orGroups = [];

    for (var j = 0; j < node.conditions.length; j++) {
      orGroups = orGroups.concat(nodeToRuleItemGroups_(node.conditions[j]));
    }

    return orGroups;
  }

  if (node.type === 'and') {
    var combinedGroups = [{ ruleItems: [] }];

    for (var k = 0; k < node.conditions.length; k++) {
      combinedGroups = combineRuleItemGroups_(combinedGroups, nodeToRuleItemGroups_(node.conditions[k]));
    }

    return combinedGroups;
  }

  throw new Error('Unsupported rule node for Google Ads payload: ' + node.type + '.');
}

function combineRuleItemGroups_(leftGroups, rightGroups) {
  var combined = [];

  for (var i = 0; i < leftGroups.length; i++) {
    for (var j = 0; j < rightGroups.length; j++) {
      combined.push({
        ruleItems: leftGroups[i].ruleItems.concat(rightGroups[j].ruleItems)
      });
    }
  }

  return combined;
}

function conditionToRuleItem_(condition) {
  return {
    name: mapExpressionFieldToGoogleAdsName_(condition.field),
    stringRuleItem: {
      operator: mapExpressionOperatorToGoogleAdsOperator_(condition.operator),
      value: trimCell_(condition.value)
    }
  };
}

function mapExpressionFieldToGoogleAdsName_(field) {
  var normalizedField = field;

  if (!GOOGLE_ADS_FIELD_NAMES[normalizedField]) {
    throw new Error('Unsupported Google Ads field mapping: ' + field + '.');
  }

  return GOOGLE_ADS_FIELD_NAMES[normalizedField];
}

function mapExpressionOperatorToGoogleAdsOperator_(operator) {
  var normalizedOperator = operator;

  if (!GOOGLE_ADS_STRING_OPERATORS[normalizedOperator]) {
    throw new Error('Unsupported Google Ads operator mapping: ' + operator + '.');
  }

  return GOOGLE_ADS_STRING_OPERATORS[normalizedOperator];
}

function resolveLookbackWindowDays_(node, defaultLookbackDays) {
  var lookbackDays = collectLookbackDays_(node, {});
  var keys = Object.keys(lookbackDays);

  if (keys.length > 1) {
    throw new Error('A single Google Ads operand cannot mix multiple lookback windows.');
  }

  if (keys.length === 1) {
    return validateMembershipDuration_(keys[0]);
  }

  return validateMembershipDuration_(defaultLookbackDays);
}

function collectLookbackDays_(node, seen) {
  if (node.type === 'condition') {
    if (hasOwnProperty_(node, 'lookbackDays') && trimCell_(node.lookbackDays)) {
      seen[validateMembershipDuration_(node.lookbackDays)] = true;
    }

    return seen;
  }

  if (node.type === 'refine') {
    collectLookbackDays_(node.base, seen);
    collectLookbackDays_(node.refinement, seen);
    return seen;
  }

  for (var i = 0; i < node.conditions.length; i++) {
    collectLookbackDays_(node.conditions[i], seen);
  }

  return seen;
}

function writeVariantStatus_(variantsSheet, sheetRowNumber, status) {
  variantsSheet
    .getRange(sheetRowNumber, VARIANT_COLUMNS.STATUS)
    .setValue(status);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseMembershipDurations_: parseMembershipDurations_,
    validateMembershipDuration_: validateMembershipDuration_,
    normalizeAudienceDefinition_: normalizeAudienceDefinition_,
    getAudienceHeaders_: getAudienceHeaders_,
    getExistingAudienceHeaders_: getExistingAudienceHeaders_,
    getHeaderBackgroundColor_: getHeaderBackgroundColor_,
    hasDefaultTemplateSheets_: hasDefaultTemplateSheets_,
    isPlaceholderValue_: isPlaceholderValue_,
    generateVariantRows_: generateVariantRows_,
    buildVariantName_: buildVariantName_,
    toVariantSheetRows_: toVariantSheetRows_,
    normalizeVariantRow_: normalizeVariantRow_,
    mapAudienceDefinitionsBySourceRow_: mapAudienceDefinitionsBySourceRow_,
    buildExistingAudienceRow_: buildExistingAudienceRow_,
    mergeExistingAudienceRows_: mergeExistingAudienceRows_,
    normalizeExistingAudienceSheetRow_: normalizeExistingAudienceSheetRow_,
    parseTagList_: parseTagList_,
    normalizeCustomerType_: normalizeCustomerType_,
    buildExecutionParsedRule_: buildExecutionParsedRule_,
    buildEventTagsRule_: buildEventTagsRule_,
    resolveCombinedAudienceReferences_: resolveCombinedAudienceReferences_,
    buildCombinedAudienceRule_: buildCombinedAudienceRule_,
    mapCustomerTypeToEnum_: mapCustomerTypeToEnum_,
    normalizeCustomerId_: normalizeCustomerId_,
    executeVariantRow_: executeVariantRow_,
    resolveExecutionVariant_: resolveExecutionVariant_,
    handleMutateResult_: handleMutateResult_,
    formatErrorStatus_: formatErrorStatus_,
    buildUserListMutateOperation_: buildUserListMutateOperation_,
    buildFlexibleRuleUserList_: buildFlexibleRuleUserList_,
    ruleNodeToOperand_: ruleNodeToOperand_,
    conditionToRuleItem_: conditionToRuleItem_,
    mapExpressionFieldToGoogleAdsName_: mapExpressionFieldToGoogleAdsName_,
    mapExpressionOperatorToGoogleAdsOperator_: mapExpressionOperatorToGoogleAdsOperator_,
    parseExpression_: parseExpression_,
    parseSimpleExpression_: parseSimpleExpression_,
    buildGeminiExpressionPrompt_: buildGeminiExpressionPrompt_,
    buildGeminiExpressionRequest_: buildGeminiExpressionRequest_,
    extractGeminiText_: extractGeminiText_,
    parseGeminiJsonText_: parseGeminiJsonText_,
    validateParsedRule_: validateParsedRule_,
    validateRuleNode_: validateRuleNode_,
    PARSED_EXPRESSION_SCHEMA: PARSED_EXPRESSION_SCHEMA,
    PARSED_EXPRESSION_VERSION: PARSED_EXPRESSION_VERSION,
    GEMINI_MODEL: GEMINI_MODEL,
    isBlankRow_: isBlankRow_,
    trimCell_: trimCell_
  };
}
