# Google Ads Audience Builder

**Status:** V1 public release

Google Ads Audience Builder is a free Google Ads script for performance marketers, freelancers, and agencies that need to create, duplicate, and bulk manage Google Ads audiences without repetitive clicking in the Google Ads interface.

Built by [Ewald van Kampen (Results Driven Marketing)](https://resultsdriven.nl/) as a free tool for the performance marketing community.

V1 focuses on the lowest-friction workflow: a Google Sheets template plus a Google Ads Script. Users define audiences once, generate membership-duration variants, and later run the script to create selected audiences in Google Ads with row-level status feedback.

## V1 Scope

V1 will provide:

- A copyable Google Sheets template for audience definitions and generated variants.
- A Google Ads Script in [`v1-script/audience-builder.js`](v1-script/audience-builder.js).
- URL-based and event-based audience setup.
- Bulk creation capabilities: bypass the slow Google Ads UI by copy-pasting rows directly in the spreadsheet.
- A variant generator for durations like `7,30,90`.
- Selective publishing: use the `create` column to cherry-pick exactly which variants to push live (e.g., only 30-day variants).
- Segment-specific setup with `expression` for `website_visitors` and `include_tags`/`exclude_tags` for `event`.
- Gemini-assisted expression parsing for marketer-readable rules.
- Per-row status feedback such as `created`, `already exists`, or `error: [message]`.
- Sheet formatting for readable columns and status colors: green for `created`, yellow for `already exists`, red for `error:`.

## Current Workflow

1. Open Google Ads -> Tools -> Scripts and create a new script.
2. Paste the script from [`v1-script/audience-builder.js`](v1-script/audience-builder.js).
3. In the script configuration, leave `SPREADSHEET_URL = '';` empty. This allows the script to automatically generate a new template sheet on its first run.
4. (Optional) Add your `GEMINI_API_KEY` if you want the script to use AI to contextually parse complex expressions. Without the API key, the script falls back to local logic for basic URL and event rules.
5. Run the script once (in Preview mode is fine). The script will create your Google Sheet and log its URL.
6. Open the logged Google Sheet URL.
7. Add audience definitions in the `AUDIENCES` tab.
8. Use segment contract:
   - `website_visitors`: fill `expression`, keep `include_tags`/`exclude_tags` empty. (Use this for URL rules and GA4-style remarketing events like `event = view_item`).
   - `event`: leave `expression` empty, fill `include_tags` (required) and optional `exclude_tags`. (Use this ONLY for exact Google Ads Conversion Action names like `Quote Form Submitted` or `Phone Call Lead`). Separate multiple conversion names with a comma. No quotes needed.
9. Optional: set a single `customer_types` value from the dropdown list.
10. Run with `RUN_MODE = generate_variants`.
11. Review `VARIANTS` and set `create` to `yes` or `no` to cherry-pick exactly which audiences to push live (e.g., only select the 30-day variants).
12. Run with `RUN_MODE = execute`.
13. Keep `DRY_RUN = true` for a safe test. Set `DRY_RUN = false` only when you are ready to create Google Ads audiences.
14. Review status per row: `created`, `already exists`, or `error: [message]`.

Safety rules:

- The script fails closed when no `VARIANTS` rows have `create=yes`.
- Always use `DRY_RUN = true` to validate the expected audiences before setting `DRY_RUN = false` to push live creations.

## Audience Expressions

Use `expression` to describe the audience rule. Simple URL and event rules parse locally; more complex expressions can fall back to Gemini.

Segment contract:

- `website_visitors`: use `expression` (URL/action/refine style rules).
- `event`: use `include_tags` and `exclude_tags` as event/conversion names.
- `customer_types`: single dropdown value (MVP single-select).

If local parser cannot parse an `expression`, Gemini fallback is used and validated before mutation.

Basic examples:

- `URL contains /thank-you`
- `URL equals /pricing`
- `URL starts_with /checkout`
- `event = view_item AND URL contains /pricing`
- `segment_type=event` with `include_tags=Quote Form Submitted` (exact Conversion Action name)
- `segment_type=event` with `include_tags=Quote Form Submitted,Phone Call Lead` (comma-separated, spaces allowed)

Conditional examples:

- `URL contains /thank-you AND event = purchase`
- `URL contains /pricing OR URL contains /demo`
- `URL contains /product NOT URL contains /return`
- `event = add_to_cart NOT event = purchase`

Advanced examples:

- `(URL contains /thank-you AND event = purchase) OR event = qualified_lead`
- `URL contains /product REFINE offer_detail contains sale IN 30 DAYS`
- `(event = add_to_cart IN 7 DAYS OR event = begin_checkout IN 7 DAYS) NOT event = purchase`

## Optional: Using Gemini for Complex Expressions

Google Ads Audience Builder includes an optional integration with Google's Gemini AI. 

**Why Gemini?** 
Many online marketers already use Google Workspace, which means you likely already have access to [Google AI Studio](https://aistudio.google.com/) using your existing Google account. This makes it incredibly easy and free to generate a Gemini API key. By adding this key, you unlock the ability to write highly complex, nested, or conversational audience rules. Gemini acts as a smart parser that translates these human-readable conditions into the strict format required by the Google Ads API.

**How to use it:**
1. Go to [Google AI Studio](https://aistudio.google.com/) and create a free API key.
2. Paste this key into the `GEMINI_API_KEY` variable at the top of your Google Ads Script.
3. The script will automatically detect the key and use Gemini whenever it encounters an expression that is too complex for the basic local parser.

**Is it mandatory?**
No, it is completely optional. If you leave `GEMINI_API_KEY = '';` empty, the script simply falls back to its internal local logic. This local fallback works perfectly fine for all standard URL and event rules (like `URL contains /thank-you` or `include_tags=Quote Form Submitted`).

## auto_add

Set `auto_add` to `yes` when Google Ads should prepopulate the audience with historical users who already matched the rule within the membership window.

Set `auto_add` to `no` when the audience should start empty and only collect users from the creation moment onward.

## Repository Structure

```text
/
├── README.md
├── LICENSE
├── v1-script/
│   └── audience-builder.js
└── .planning/
```

## Security Checklist

Never commit:

- Google Ads credentials
- OAuth refresh tokens
- Gemini API keys
- Google Ads account IDs for private customers
- Customer data
- Private spreadsheet URLs unless they are intended as public template links
- Local sheet exports containing customer data

Never paste your Gemini API key into a shared or public Google Sheet. Keep the key in your private Google Ads Script configuration.

## Contributing

Issues and pull requests are welcome. Keep the project focused on audience management only: no reporting, campaign management, bidding, ad copy, or keyword tooling.

## Built By

Built by [Ewald van Kampen (Results Driven Marketing)](https://resultsdriven.nl/) as a free tool for the performance marketing community.

Need help with Google Ads audiences? Contact [resultsdriven.nl](https://resultsdriven.nl).
