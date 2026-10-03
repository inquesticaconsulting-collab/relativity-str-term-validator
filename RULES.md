# Validation rule catalog

Every finding `str_term_validator.py` can raise. **ERROR** must be resolved before a term is loaded into an STR, **WARNING** needs a recorded decision, **INFO** is a note. Most findings are auto-repaired; the report's *Revised Term Check* column shows whether the repaired term re-validates CLEAN, ADVISORY, or NEEDS REVIEW.

| Code | Severity | What it detects | Handling | Example (original → revised) |
|---|---|---|---|---|
| `OVER_450_CHARS` | ERROR | Term exceeds the STR 450-character limit. | Relativity rejects it. Not auto-fixable (NEEDS REVIEW): split into several lines — each line is its own query and reports separately. | — |
| `UNBALANCED_QUOTES` | ERROR | Odd number of double quotes; phrase never closed. | Query is invalid/unpredictable. Auto-balanced — verify the revision. | "unclosed phrase  →  "unclosed phrase" |
| `UNBALANCED_PARENS` | ERROR | Parentheses do not balance. | Grouping is ambiguous or invalid. Auto-balanced — verify the revision. | (missing close  →  (missing close) |
| `EMPTY_PARENS` | ERROR | Empty group (). | Nothing to search; removed automatically. | term ()  →  term |
| `COMMA_SEPARATED_LIST` | ERROR | Several terms pasted on one line, separated by commas/semicolons (outside quotes). | dtSearch treats commas as spaces — the line runs as ONE exact phrase and hits almost nothing. Rewritten with OR; splitting to separate lines is better for per-term counts. | apple, pear  →  apple OR pear |
| `STRAY_SEPARATOR` | INFO | A leading/trailing comma or semicolon. | Carries no meaning in dtSearch; removed. | apple,  →  apple |
| `LEADING/TRAILING/DOUBLE_OPERATOR` | ERROR | Operator with a missing operand at start/end, or two binary operators in a row. | Invalid query. Auto-repaired by dropping the dangling operator; in 'AND w/10' the proximity operator is kept. | contract AND AND breach  →  contract AND breach |
| `OPERATOR_AFTER/BEFORE_PAREN` | ERROR | Binary operator immediately inside a group boundary. | A group cannot begin or end with an operator; auto-repaired. | ( AND x)  →  (x) |
| `NOT_BEFORE_OPERATOR` | ERROR | NOT directly followed by AND/OR ('NOT OR'). | Invalid. The NOT is dropped and the operator kept — but this pattern often means AND NOT (exclusion) was intended: the two readings are opposites. | x NOT OR y  →  x OR y (confirm AND NOT wasn't meant) |
| `PROX_SPACING` | ERROR | Space inside a proximity operator. | dtSearch only recognizes w/N or pre/N with no spaces; otherwise 'w' is searched literally. Auto-corrected. | invoice w/ 5 payment  →  invoice W/5 payment |
| `PROX_TRAILING_SLASH` | ERROR | Proximity with a trailing slash, e.g. w/7/. | Unrecognizable to dtSearch (the terms would run as a phrase). Auto-corrected to w/7. | a w/7/ b  →  a W/7 b |
| `PROX_DISTANCE_ASSUMED (was PROX_MALFORMED)` | WARNING | w/ or pre/ with no number. | A default of w/5 is inserted so the term runs, labeled ASSUMPTION APPLIED for confirmation. | invoice w/ payment  →  invoice W/5 payment (assumed) |
| `PROX_ZERO` | WARNING | Distance of zero (w/0). | Auto-replaced with w/1 (within one word, either order), labeled ASSUMPTION APPLIED. A quoted phrase is the ordered alternative. | apple w/0 pear  →  apple W/1 pear (assumed) |
| `LEADING_WILDCARD` | ERROR | Wildcard at the start of a word. | Prohibited in STRs and a documented performance hazard. Removed; use trailing wildcard, stemming (~), or explicit OR variants. | *fraud*  →  fraud* |
| `FUZZY_NOT_SUPPORTED` | ERROR | Fuzzy operator %. | Not supported in STRs. Removed; pre-expand via the dtSearch Dictionary and join variants with OR. | app%le  →  apple (plus Dictionary variants) |
| `STEM_LEADING` | ERROR | ~ at the start of a word. | Stemming goes at the END of the root (English only). Moved there automatically. | ~apply  →  apply~ |
| `STRAY_TILDE` | ERROR | ~~ or ~~~ attached to a word or quoted phrase (not a numeric range). | Invalid decoration — removed. ~~ is only valid between integers (12~~24); a single ~ only as stemming on a bare word. | "confidential"~~  →  "confidential" |
| `PHONIC_MISPLACED` | WARNING | # inside or after a word. | Removed as a likely typo (labeled ASSUMPTION APPLIED); # is only valid immediately before a word. | proj#ect  →  project (assumed typo) |
| `NUMERIC_RANGE_MALFORMED` | ERROR | Invalid N~~M numeric range. | Both bounds must be integers. The intended range can't be guessed, so this stays NEEDS REVIEW. | 12~~abc  →  (human fixes, e.g. 12~~24) |
| `WILDCARD_ONLY` | ERROR | Term is only wildcards. | Matches every word in the index (including xfirstword/xlastword) — returns all documents. | *  →  (replace with a real term) |
| `NOISE_WORD_TERM` | ERROR | Standalone term is a dtSearch noise word. | Returns zero hits on a default index. NEEDS REVIEW: replace, or have the admin remove it from the noise list and rebuild. | the  →  (strike or replace) |
| `FIELD_SYNTAX_UNSUPPORTED` | ERROR | Metadata written as field:value (custodian:, subject:, date:, ext:, folder:, project:). | Not dtSearch syntax — an STR term searches content only. Extracted to the 'Extracted Conditions' sheet; they belong in the searchable-set saved search. Values are validated (impossible dates, misspelled extensions and field names). | custodian:Smith AND breach  →  breach (condition listed for the saved search) |
| `UNSUPPORTED_FUNCTION` | ERROR | Function-style syntax dtSearch doesn't have, e.g. stemming(). | Only date(), mail(), creditcard() exist. Removed; for stemming, use word~. | stemming()  →  (removed) |
| `MIXED_AND_OR` | WARNING | AND and OR mixed with no parentheses. | dtSearch applies OR before AND. The revision adds parentheses showing how the query actually runs — confirm intent. | a AND b OR c  →  a AND (b OR c) |
| `AND_INSIDE_PROX_GROUP` | WARNING | Group joined by w/N or pre/N contains AND. | Documented ambiguous pattern with unpredictable results. Use OR inside proximity-connected groups, or restructure. | (a AND b) w/10 (c AND d)  →  (a w/10 c) AND (b w/10 d) |
| `NESTED_PROX / MANY_PROX_OPS` | WARNING | Chained/nested or numerous proximity operators. | Discouraged for performance; chained proximity can evaluate ambiguously. Split across lines. | a w/3 b w/3 c  →  (a w/3 b) AND (b w/3 c) |
| `PROX_TOO_WIDE` | WARNING | Very wide proximity window (above ~25 words). | Behaves like AND across paragraphs; inflates hits. Confirm width. | a w/50 b  →  confirm or narrow |
| `NOISE_WORD_IN_QUERY` | WARNING | Noise word inside an unquoted expression. | The word is skipped by the index; in phrases it matches any single word in that position. Verify counts make sense. | statement of work — 'of' acts as a placeholder |
| `RESERVED_WORD_UNQUOTED` | WARNING | 'to' or 'contains' used as an ordinary word outside quotes. | These are dtSearch connector words; quote the phrase so they are searched literally. | go to market  →  "go to market" |
| `SPECIAL_CHAR` | WARNING | Character outside the default searchable alphabet (&, _, :, @, etc.). | Not matched literally by a default index (AT&T indexes as 'at' + 't'). Verify alphabet file; consider phrase forms, mail(), or regex. | AT&T  →  "at t" (default index) |
| `HYPHENATED_TERM` | WARNING | Hyphenated word. | Hyphen handling is index-configurable (commonly behaves as a space). Search all variants. | e-mail  →  ("e-mail" OR email OR "e mail") |
| `UNICODE_CHARS` | WARNING | Smart quotes/dashes/non-breaking spaces from Word. | Not treated as phrase delimiters or indexed as typed. Auto-replaced with ASCII equivalents. | “phrase”  →  "phrase" |
| `EARLY_WILDCARD / MULTI_WILDCARD` | WARNING | * near word start, or several * in one word. | Performance and over-inclusion risk. Narrow or enumerate variants. | b*n*a  →  banana OR ... |
| `MULTI_WILDCARD_RUN` | WARNING | Consecutive asterisks (**). | Behaves like a single * but signals a typo. Auto-collapsed to one *. | .doc**  →  .doc* |
| `X_WILDCARD_TRAP` | WARNING | Pattern matching built-in xfirstword/xlastword tokens. | Those tokens exist in every document — the pattern returns everything. | x*  →  narrow the pattern |
| `OVERBROAD_TERM` | WARNING | Very short standalone term. | Matches enormous volume. Add context via proximity or qualifiers. | ab  →  ab w/10 <context> |
| `REDUNDANT_OPERAND` | WARNING | Same word on both sides of AND/OR. | Equivalent to the single word; usually a paste slip. Collapsed — check whether a different term was intended. | deal OR deal  →  deal |
| `DUPLICATE_TERM` | WARNING | Line duplicates an earlier term (after normalization). | Relativity ignores duplicates on load; removed from the clean list. | — |
| `IMPLICIT_PHRASE` | INFO | Adjacent words with no operator. | dtSearch runs them as an exact phrase, not an AND. Quoted for clarity; use AND if both-anywhere was intended. | market analysis  →  "market analysis" |
| `OPERATOR_CASE` | INFO | Operators in lowercase. | Recognized either way; capitalized by convention so operators are visually distinct from literal words. | a and b  →  a AND b |
| `NOT_WN_ASYMMETRY` | INFO | Query uses NOT w/N. | NOT w/N is not symmetrical — order matters. Confirm direction. | a NOT w/5 b  ≠  b NOT w/5 a |
| `POSSIBLE_MISSPELLING` | INFO | Words not found in the dictionary (optional; needs pyspellchecker). | No auto-change — misspellings are sometimes deliberate in eDiscovery (they catch the same typo in documents). Suggestions listed for the requestor. | contrct  →  suggestion: contract? (no change made) |
| `NOISE_WORD_IN_PHRASE / WHITESPACE / CONTROL_CHARS / EMPTY_LINE` | INFO | Hygiene notes. | Phrase noise words act as one-word placeholders; whitespace/control characters trimmed; blank lines skipped. | — |

_Generated by `docs/build/build_review_tracker.py`._
