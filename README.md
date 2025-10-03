# googlesheets-translate-align
Google Sheets script panel (sidebar) that automates text translation into 11 languages from RU/EN (EN, DE, FR, ES, PT, IT, GR, PL, HU, ID, ZH) using DeepL and ensures a strict 1:1 alignment of rows with the source text. Russian-language interface. 

**Keys and settings are defined in Script Properties inside your Apps Script project.**

How it works:
1) Selects the reference column based on priority.
2) Automatically translates it into the chosen languages via DeepL.
3) Aligns the translations so that the number of rows always matches the original 1:1.
4)Saves the result to a new Google Sheets tab (without altering the source data).

##Features
1) Selection of the reference column (the defaults are set for my project, but you can define your own).
2) Automatic translation via DeepL (batch) or alignment without translation.
3) Autofill of missing rows with translation.
4) Glossary application (simple find & replace rules by target language).
5) Batch processing with a progress bar and output to a new sheet.

##Structure

1) appsscript.json — project manifest (scopes, whitelist, add-on settings)
2) Code.gs — main logic (GAS/JavaScript)
3) Sidebar.html — UI panel (HTML + JS)

##Installation (how to run it yourself)

1) Open script.google.com → New project.

2) Create the files:

Code.gs — paste the contents.
Sidebar.html — paste the contents.
appsscript.json — paste the contents.

3) In appsscript.json, make sure the whitelist for DeepL is present (already included, otherwise it won’t work).

4) In Apps Script: Project Settings → Script properties add:
DEEPL_KEYS — your DeepL API key (or multiple keys separated by commas).

##Script Panel Features:

At the top of the sidebar there is a status line:
🟢 **"Готов к работе😎"** — green text.
This means the DeepL key (-s) were found in Script Properties (DEEPL_KEYS), and the panel is ready to translate.

⏳ **Checking keys…** — the panel is still loading the status (usually takes 1–2 seconds after opening).
⚠️ **Problems detected** — if translation doesn’t start, click Check DeepL.
In diagnostics, you’ll see error codes (e.g., 456 — character limit, 403/401 — invalid key).

*Sheet name* — the name of the sheet with the source data. It must match the sheet name in the document you are translating from/into.

*Source header (reference column)* — optional: you can leave it empty, specify it manually, or click Detect. If left blank and simply run, the script will auto-detect the source column by condition.

*Target columns (manual input)* — a list of language columns to generate as output (an example is shown directly in the input field). You can specify any number of columns in any order.

*Language map* — mapping of languages in the format column code → DeepL code. Defaults are set automatically (Greek, Chinese → Traditional Chinese, Portuguese, English → American English), but you can modify them directly in the field to suit your preferences.

**Checkbox modes (all enabled by default; if not needed, leave them as is):**

*Auto-translate from source via DeepL*

**ON** — translates all rows into the target columns.
**OFF** — aligns existing texts in the target columns with the source (useful if translation was done outside the script).

*Fill missing rows via DeepL* — works only when Auto-translate is **OFF**.
Meaning: if the target column has gaps, the script will translate only the empty rows from the source via DeepL.
Empty cells → filled.
Already filled cells → untouched.

**Use it** when you already have a partial translation (e.g., DE is 70% complete) and want to fill only the missing 30%.
**Disable it** if you only want to align the existing translation with the source (no extra translation).

*Strict 1:1 (don’t add columns)* — checks that all target columns already exist in the first row of the source sheet. If even one is missing → error, and nothing runs.

*Rows per cycle (batch)* — number of rows processed per cycle (default: 150).
If handling very large sheets and errors occur, lower the value (e.g., 100 or 80).
Processing will take slightly longer.

*Write to new sheet* — name of the new sheet to store results (if it already exists, it will be recreated).

*Run* — start the script.

*Glossary* — internal glossary (imported from DeepL by default; you can extend and edit it).

*Check DeepL* — checks DeepL status for errors (for admin use).

**Quick setup examples**

Translate everything from scratch: Auto-translate = ON, Fill missing = irrelevant, Strict = optional.
Translations shifted, need to realign and fill gaps: Auto-translate = OFF, Fill missing = ON, Strict = ON.
Align only (no translation): Auto-translate = OFF, Fill missing = OFF, Strict = ON.
Add a new language when the column doesn’t exist in the header: Strict = OFF (column will be created in the result sheet).

**The DeepL API Free plan provides only 500,000 characters per month, which is often not enough for large-scale translations. For heavy usage, either switch to a DeepL API Pro key or change the translation backend to OpenAI GPT (usually much cheaper). The user interface, job handler, alignment, and glossary logic remain unchanged.**
