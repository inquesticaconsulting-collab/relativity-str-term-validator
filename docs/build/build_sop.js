// Build docs/SOP_STR_Term_Validation.docx   (run: node docs/build/build_sop.js)
const path = require("path");
const fs = require("fs");
const {
  BorderStyle, Document, HeadingLevel, LevelFormat, Packer,
  Paragraph, Table, TableCell, TableRow, TextRun, WidthType,
} = require("docx");

const TABLE_W = 9360;
const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 259 }, ...opts,
  children: [new TextRun({ text, size: 22, ...opts.run })],
});
const italic = (text) => p(text, { run: { italics: true } });
const bold = (text, size) => p(text, { run: { bold: true, size: size || 22 } });
const head = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 },
  children: [new TextRun({ text, size: 24, bold: true, color: "000000" })],
});
const stepHead = (text) => new Paragraph({
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, size: 22, bold: true })],
});
const cmd = (text) => new Paragraph({
  spacing: { before: 40, after: 120 }, indent: { left: 480 },
  children: [new TextRun({ text, font: "Consolas", size: 20 })],
});
const bullet = (text, lead) => new Paragraph({
  numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
  children: lead
    ? [new TextRun({ text: lead, bold: true, size: 22 }), new TextRun({ text, size: 22 })]
    : [new TextRun({ text, size: 22 })],
});
const numbered = (text) => new Paragraph({
  numbering: { reference: "nums", level: 0 }, spacing: { after: 80 },
  children: [new TextRun({ text, size: 22 })],
});

function tbl(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const scaled = widths.map((w) => Math.round((w / total) * TABLE_W));
  const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (t, isHead, width) => new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { after: 0 },
      children: [new TextRun({ text: String(t), size: 20, bold: !!isHead })] })],
  });
  return new Table({
    width: { size: TABLE_W, type: WidthType.DXA }, columnWidths: scaled,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, true, scaled[i])) }),
      ...rows.map((r) => new TableRow({ children: r.map((v, i) => cell(v, false, scaled[i])) })),
    ],
  });
}

const children = [
  bold("Standard Operating Procedure", 28),
  p("Search Term Validation for Relativity Search Terms Reports — Validation Procedure",
    { spacing: { after: 240 }, run: { size: 24 } }),
  tbl(["Field", "Value"], [
    ["Document ID", "SOP-EDISC-002"],
    ["Version", "1.0"],
    ["Effective Date", "August 30, 2026"],
    ["Prepared By", "Neeraj Khandelwal, Inquestica Consulting"],
    ["Applies To", "str_term_validator.py (STR search-term validation script)"],
    ["Review Cycle", "Annually, or immediately upon any change to the script"],
    ["Classification", "Public — released under the MIT License"],
  ], [28, 72]),
  p("", { spacing: { after: 120 } }),

  head("1. Purpose"),
  p("This SOP defines the standardized, repeatable procedure for validating, auto-repairing, and approving proposed search terms before they are loaded into a Relativity Search Terms Report (STR), using the str_term_validator.py script. Every line entered into an STR runs as an individual dtSearch query, and dtSearch does not reject most defective queries — it runs them and silently counts the wrong documents. This procedure ensures that anyone following it produces consistent, auditable, syntactically valid terms suitable for case-team and review reporting."),

  head("2. Scope"),
  p("This procedure applies to any team member responsible for validating search terms received from case teams, clients, or opposing parties before an STR is created or populated. It covers script execution on a standard workstation with Python installed."),
  italic("Out of scope: installing Python itself from scratch, creating the dtSearch index or the searchable-set saved search inside Relativity, negotiating search terms with opposing counsel, and interpreting the legal significance of terms or hit counts."),

  head("3. Definitions"),
  tbl(["Term", "Definition"], [
    ["STR", "Search Terms Report — a Relativity object reporting document hit counts per term; each line of its term list runs as one dtSearch query."],
    ["dtSearch", "The search engine behind STRs. Case-insensitive; searches a pre-built index of document content."],
    ["Noise word", "A common word (the, of, and) a default index does not make searchable; standalone it returns zero hits."],
    ["Metadata condition", "A field:value restriction (custodian:, date:, ext:, folder:, project:, subject:). Not dtSearch syntax — belongs in the searchable-set saved search, not in the term."],
    ["Validation report", "The Excel output: Summary, Validation Detail (one row per term with findings, reasoning and the revised term), and Extracted Conditions sheets."],
    ["Clean term file", "The .txt output of revised terms, one per line, formatted for the STR Add Terms box. Excludes terms whose check is NEEDS REVIEW."],
    ["Revised Term Check", "The script's self-verification of each revised term: CLEAN (no findings), ADVISORY (only warnings/notes remain), or NEEDS REVIEW (an error requires human input)."],
    ["ASSUMPTION APPLIED", "Label marking any repair based on a default rather than the source (e.g. w/5 for a missing proximity distance). Every assumption must be confirmed with the case team."],
  ], [22, 78]),

  head("4. Roles & Responsibilities"),
  bullet(" — executes the script, works the validation report, drafts revisions, obtains approvals, and distributes the output.", "Term Validator (analyst)"),
  bullet(" — owns the terms; approves every revision that changes search meaning, every assumption, and every struck term.", "Case Team / Requestor"),
  bullet(" — confirms the dtSearch index configuration (noise-word list, alphabet file) and applies the extracted metadata conditions to the searchable-set saved search.", "Relativity Administrator"),
  bullet(" (where required by matter protocol) — reviews the final term list before it is loaded.", "Reviewing Manager/Attorney"),

  head("5. Prerequisites"),
  stepHead("5.1 Software"),
  bullet("Windows workstation (Command Prompt) — macOS/Linux Terminal also supported."),
  bullet("Python 3.9 or later, installed and available on PATH."),
  bullet("pip package manager (installed automatically with Python)."),
  bullet("Python packages: openpyxl (required); pyspellchecker (optional — enables spelling suggestions)."),
  stepHead("5.2 Files"),
  bullet("str_term_validator.py — the current approved version of the script."),
  bullet("Source term list as .xlsx, .csv, or .txt — strictly one term per line/row. In a spreadsheet, a column headed \u201CTerm\u201D (or \u201CRevised Term\u201D) is detected automatically; otherwise the first column is read."),

  head("6. Procedure"),
  stepHead("Step 1 — Confirm Python is installed"),
  p("Open Command Prompt: press the Windows key, type cmd, and press Enter. The prompt should read similar to C:\\Users\\YourName> — not >>>. Run:"),
  cmd("py --version"),
  italic("Confirm a version beginning with \u201C3.\u201D is printed. Note: commands in this SOP are typed at the Command Prompt, never at Python's own >>> interactive prompt — typing pip commands at >>> raises SyntaxError."),

  stepHead("Step 2 — Install required packages (one time per machine)"),
  cmd("py -m pip install openpyxl pyspellchecker"),
  p("Confirm the command completes without red \u201CERROR\u201D text. Re-run any time the script reports ModuleNotFoundError. pyspellchecker is optional; without it the script runs normally and skips spelling suggestions."),

  stepHead("Step 3 — Stage the input file and script"),
  p("Place str_term_validator.py and the source term list in a known working folder, for example C:\\data\\. Do NOT pre-clean the terms by hand — paste them exactly as received. The script detects and documents every defect with its reasoning, which preserves the audit trail of what was changed and why; a silent manual edit breaks that trail."),

  stepHead("Step 4 — Run the script"),
  cmd("py C:\\data\\str_term_validator.py C:\\data\\TERMS.xlsx -o C:\\data\\TERMS_validation.xlsx --clean C:\\data\\TERMS_clean.txt"),
  p("Replace TERMS with the actual file name. The flags are optional: without -o the report is written next to the input as <input>_validation.xlsx; --clean additionally writes the STR-ready term list."),

  stepHead("Step 5 — Review console output"),
  p("The script prints a summary to the Command Prompt window. Confirm: the number of terms validated matches the number submitted, and note the self-check line (\u201CRevised-term self-check: X clean, Y advisory, Z need review\u201D). A shortfall in the term count means a formatting problem in the input file — fix the file and re-run rather than reconciling by hand."),

  stepHead("Step 6 — Work the validation report"),
  p("Open the validation .xlsx. On the Validation Detail sheet, work top to bottom and record a decision for every flagged row. The Revised Term Check column governs what may proceed:"),
  bullet(" — the revised term re-validates with no findings. May be loaded after approval.", "CLEAN"),
  bullet(" — only warnings or notes remain (index-dependent or judgment items). May be loaded after the warnings are reviewed and approved.", "ADVISORY"),
  bullet(" — an error could not be auto-fixed (e.g. a standalone noise word, a malformed numeric range, a term over 450 characters). Must be resolved by a person; these terms are excluded from the clean file automatically.", "NEEDS REVIEW"),
  p("Confirm every finding labeled ASSUMPTION APPLIED with the case team — these are defaults the script chose (w/5 distance, w/1 for w/0, removed mid-word #), not values taken from the source. Hand the Extracted Conditions sheet to the Relativity Administrator: those custodian/date/extension/folder restrictions must be configured on the searchable-set saved search, not entered as terms. To re-check edited terms, the report itself can be fed back into the script — it reads the Revised Term column."),

  stepHead("Step 7 — Approve, load, and file"),
  p("Obtain the case team's written approval for meaning-changing revisions, assumptions, and struck terms. Paste the contents of the clean .txt into the STR Add Terms box (one term per line) and confirm the STR's term count matches the approved list — Relativity silently ignores exact duplicates. Rename and save the report using the matter naming convention:"),
  cmd("[MatterName]_STRTermValidation_[YYYYMMDD].xlsx"),
  p("Store the report, the clean file, and the approval record in the designated review-reporting folder and distribute per matter protocol."),

  head("7. Command-Line Reference"),
  tbl(["Argument", "Required", "Description", "Default"], [
    ["input", "Yes", "Path to the source term list (.xlsx, .csv or .txt), or a previous validation report.", "—"],
    ["-o / --output", "No", "Path for the Excel validation report.", "<input stem>_validation.xlsx"],
    ["--clean", "No", "Also writes an STR-ready term list (.txt): revised terms, one per line, NEEDS REVIEW terms excluded, duplicates removed.", "off"],
    ["--no-spellcheck", "No", "Skips the optional spelling suggestions.", "off"],
    ["--version", "No", "Prints the script version and exits.", "—"],
  ], [20, 12, 46, 22]),

  head("8. Quality Control Checklist"),
  bullet("Term count in the console summary matches the number of terms submitted."),
  bullet("No row remains at NEEDS REVIEW, or each remaining one has a documented resolution."),
  bullet("Every ASSUMPTION APPLIED finding confirmed with the case team."),
  bullet("Every meaning-changing revision and struck term has written case-team approval on file."),
  bullet("Extracted Conditions sheet handed to the Relativity Administrator and applied to the searchable-set saved search."),
  bullet("STR term count after pasting reconciles to the approved list."),
  bullet("Output named and saved per the matter naming convention."),
  bullet("Output reviewed by a second team member, where required by matter protocol."),

  head("9. Troubleshooting"),
  tbl(["Symptom", "Likely Cause", "Resolution"], [
    ["ModuleNotFoundError: No module named 'openpyxl'", "Required package not installed for the Python interpreter in use.", "Open Command Prompt (not the Python >>> prompt) and run: py -m pip install openpyxl"],
    ["SyntaxError: invalid syntax when typing \u201Cpip install...\u201D", "Command was typed at Python's own >>> interactive prompt instead of the Windows Command Prompt.", "Type exit() to leave Python; open a new Command Prompt window; confirm the prompt reads C:\\...> before typing any pip command."],
    ["Unsupported input type", "Input file is not .xlsx, .csv or .txt.", "Save the term list in one of the supported formats, one term per line/row."],
    ["Script reads the wrong column", "Spreadsheet has no recognizable \u201CTerm\u201D header.", "Rename the terms column to \u201CTerm\u201D, or move the terms to the first column."],
    ["Many terms flagged FIELD_SYNTAX_UNSUPPORTED", "Terms contain metadata restrictions (custodian:, date:, ext:).", "Expected behavior — this is not dtSearch syntax. Apply the Extracted Conditions sheet to the searchable-set saved search; the content portion of each term is repaired automatically."],
    ["No spelling suggestions in the report", "Optional pyspellchecker package not installed.", "Run: py -m pip install pyspellchecker and re-run the script (the report is complete without it)."],
    ["Revised term still shows ADVISORY after re-running", "Advisory findings are index-dependent or judgment items (noise words, special characters, nested proximity).", "Expected behavior — advisory items are cautions, not defects. Review and approve them; they do not block loading."],
    ["Script cannot be found / \u201Cpython is not recognized\u201D", "Python is not installed, or not added to PATH.", "Reinstall Python from python.org, checking \u201CAdd python.exe to PATH\u201D during setup, or use the py launcher as shown in this SOP."],
  ], [26, 30, 44]),

  head("10. Appendix A — Script Logic Summary"),
  p("For reference, str_term_validator.py performs the following steps internally:"),
  numbered("Reads the term list (.xlsx/.csv/.txt), treating each line as one dtSearch query, and normalizes pasted text (smart quotes, dashes, hidden characters, whitespace)."),
  numbered("Extracts metadata field:value conditions (custodian:, subject:, date:, ext:, folder:, project:) out of each term, validates their values (impossible dates, misspelled extensions and field names), and lists them on the Extracted Conditions sheet for the saved search."),
  numbered("Runs 40+ checks against dtSearch syntax rules and STR-specific limits: operators, proximity, wildcards, stemming/phonic/fuzzy modifiers, noise words, reserved words, special characters, the 450-character limit, and more."),
  numbered("Auto-repairs every fixable defect and re-runs the full check set on the revised term repeatedly until it stops changing, so a fix never leaves behind a defect a second run would flag."),
  numbered("Self-verifies the final revised term and records the result in the Revised Term Check column: CLEAN, ADVISORY, or NEEDS REVIEW. Assumptions are labeled ASSUMPTION APPLIED."),
  numbered("Optionally flags likely misspellings with suggestions (no automatic change — deliberate misspellings are sometimes intended in eDiscovery)."),
  numbered("Writes the color-coded Excel report (Summary, Validation Detail, Extracted Conditions) and, with --clean, the STR-ready term list."),

  head("11. Revision History"),
  tbl(["Version", "Date", "Author", "Description of Change"], [
    ["1.0", "2026-08-30", "Neeraj Khandelwal", "Initial SOP release."],
  ], [14, 18, 22, 46]),
  p("", { spacing: { after: 200 } }),
  italic("\u00A9 2026 Inquestica Consulting \u2014 Inhaber: Neeraj Khandelwal. Released under the MIT License. Not affiliated with or endorsed by Relativity ODA LLC; not legal advice."),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22, color: "000000" } } } },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
      style: { paragraph: { indent: { left: 420, hanging: 300 } } } }] },
    { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      style: { paragraph: { indent: { left: 420, hanging: 300 } } } }] },
  ] },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children }],
});

const out = path.join(__dirname, "..", "SOP_STR_Term_Validation.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
