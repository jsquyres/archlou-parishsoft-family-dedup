# Notes for Claude (and other maintainers)

This file records what was learned while building `dedup.py` (October
2026), so a future session can resume without rediscovering it. Read
README.md first for usage; this file covers the why.

## Hard rules

- **No personal data in this repository, ever.** No names, DUIDs, emails,
  phones, addresses or birthdates from the exports. That covers code
  comments, tests, docs, commit messages and issue text. Use invented
  examples ("John + Mary vs John + Susan"). Parish names and Organization
  IDs are not personal data.
- Exports, caches, reports and the regression expectations CSV live in the
  data directory outside the checkout (default `~/archlou-dedup`). Never
  write output into the repo; `--out-dir` is required for that reason, and
  `run.sh` refuses a data directory inside the repo.
- Never send export data to any external service.
- Commits: `git commit -s`, no AI attribution trailers, short subject line
  plus a body explaining why.

## Who wants what

The user (Jeff Squyres) is producing this for the Archdiocese of
Louisville, which will use ParishSoft's Archdiocese-level merge tool. Their
stated requirements:

1. List all duplicate Families by name and DUID, with a confidence level,
   the reasons for each match (grouped so that shared reasons are listed
   once), and the DUID that should survive the merge.
2. Order by confidence (highest first), then last name.
3. A closing summary of each type of match and how many Families it
   covers.
4. A Family that moved merges into the parish where it is Active.
5. Survivor order: Registered, Active, then the record with the most
   recent activity (then has a living member, most fields, lowest DUID were
   added and accepted).
6. Divorce and remarriage, grown children, and widowed-and-remarried must
   be handled properly (the first two are not duplicates; the third is).
7. The tool is one-off but will be rerun many times during the cleanup,
   and will change over time; keep it in git with full docs.
8. Run an independent code review after significant changes, when it is
   worthwhile. The user asked for this repeatedly; reviews found real bugs
   in every round of new logic.

## The data (as of the October 2026 exports)

Family export: 159,235 rows, one per DUID (no duplicate DUIDs). 105
Organization names but 112 Organization IDs (some parishes share a name),
so group by ID. Only 35% Registered; 54% in the Inactive group. Group Name
values include Active, Inactive, Contributor, No Surviving Members,
Unknown, Visitor, School Family-Registered, Sacrament Only, Friend, Staff,
Moved and a few stray surnames from shifted rows. About 4,000 Families sit
in the Archdiocese's own organization rather than a parish.

Quirks the code handles:

- **Placeholders:** `No Name` first names, "Loose Collection"/"Loose
  Cash", "Unknown Family NNNNN", businesses and organizations stored as
  Families, "Certificate People" catch-all records (6+ member surnames).
- **Sentinel values:** "Unknown", "N/A", "no email", "none", "visitor",
  "UNLISTED", "No phone on record", "Blank on SIF 2006", "*".
- **Shared contacts:** nursing homes and religious houses (one address,
  50+ Families), parish office phones. Any contact used by more than three
  surnames is ignored as evidence.
- **ZIPs:** mixed ints and strings; leading zeros lost; Home Zip Plus holds
  only the +4.
- **Shifted rows:** 3 Family rows have a phone number in the ZIP column
  (excluded). The enhanced export has ~350 rows broken by multi-line Notes
  (values shifted into Publish/Country/etc. columns); only rows whose DUID
  parses and whose "is Registered" is Yes/No are used, and only for Date
  Last Updated and Emergency Phone.
- **Enhanced "Duplicate Family DUID":** 975 values, all pointing at DUIDs
  that no longer exist (past merges). Not useful; ignored.
- **Bulk dates:** Date Unregistered has bulk stamps (14,869 at
  2024-12-11 10:45:06; others in 2014, 2021–2024; one job unregistered
  ~600 Families at ~1/second on 2024-10-25). Member Date Modified is mostly
  2023–24 conversion jobs (up to 14,000 in one minute). Family Date Last
  Updated has many bulk days, some confined to one parish. Hence the
  day/parish-day/minute bulk thresholds.
- **Member export:** 381,561 rows, ~159k Families. Birthdates on 78%; only
  ~2,400 are Jan 1 placeholders (treated as year-only). Member Status
  includes "Deceased" (also check is Deceased, Date of Death, "Dec."
  types). 119 Member Type values: Head, Child, Wife, Husband, Unknown (19k),
  Son, Daughter, Adult (8k), Young Adult, Other, Grandchild, Mother,
  Father, Ex-Spouse, Divorced, truncated forms ("Grandchi", "Adult Daug").
- **Placeholder members:** ~2,300 Members have the surname typed as the
  first name (contribution records). Dropped unless they have a birthdate,
  a different nickname, or a specific type.
- **Corrupted birth years:** ~35 single Inactive Heads in Organization
  9329 have birth years 1999–2001 that match
  people born 1937–43 elsewhere on month and day. Left as a data issue:
  they read as different people. The Archdiocese may want to fix them.
- **Personal Email column** in the Member export is empty.

## Design decisions and why

- **Explicit rules, not ML** (splink/dedupe were rejected): every match
  needs human-readable reasons.
- **Blocking** on email, phone, address, surname+ZIP and person
  (given-name alias group + exact birthdate + surname/maiden). The person
  key catches moves where every contact changed.
- **Names:** `NAME_GROUPS` are alias groups, not a single canonical form,
  because "Pat" must match both Patrick and Patricia while Patricia must
  not match Patrick. `word_match` adds a ≤3-letter prefix rule and
  Jaro-Winkler ≥0.93, and blocks masculine/feminine forms (`CROSS_GENDER`,
  `-a` endings, o/a endings) unless an alias group says otherwise.
  Two-word given names must agree in both words when both have two.
- **Roles:** partner (head/spouse), child, ex, other. Unknown/Adult types
  are resolved by age: a minor, or 18+ years younger than another member,
  is a child; 18+ years older than an explicit partner is "other" (a parent
  living with the household); with no birthdate and an explicit couple
  present, "other"; otherwise a "guessed" partner.
- **Partner presence is lenient** (same name incl. middle-name aliases,
  same exact birthdate, one-letter spelling variant, name in the other
  Family's first-names field), but only against partners/exes, and never
  against someone whose birthdate conflicts. This tolerates maiden names,
  typos and swapped spouse birthdates without letting a child named after
  a parent stand in for an absent spouse.
- **Kept-apart pairs are also grouping constraints** ("cannot-links").
  Without that, two divorced households each matching an old single record
  were chained into one group. Pairs that were never compared directly are
  checked on demand (`apart()` in `main`) before two groups join.
- **Two single-adult records that share no one** are not cannot-links:
  each spouse's pre-marriage record must be able to join the couple's
  record. Exception: a generation gap.
- **Widowed and remarried** needs exactly one new spouse, a deceased absent
  spouse in the other Family, a living partner in both, and both spouses
  within 18 years of the shared partner. Looser versions merged a widowed
  parent into a grown child's household.
- **Grouping order:** tier, then evidence strength (`edge_strength`), then
  the better survivor, then DUIDs, so ties are deterministic and strong
  evidence claims records first.
- **Survivor activity** ignores bulk jobs (see above).

## Review history

Five independent reviews (October 12, 2026), each run as a separate agent
with data access, found real issues each time:

1. Grown children typed "Adult"; lone exact match at Very high;
   asymmetric member evidence; Patricia/Patrick; "estate"/"funeral"
   surnames; bulk unregistration by second; grouping confidence; address
   street-type removal.
2. Spouse recorded differently counted as missing (false divorce);
   role mismatch on duplicate records; name gaps (Terry/Teresa…);
   overstated indirect confidence; loose birth-year typos.
3. Child's name standing in for a missing spouse; widowed rule with no one
   in common; swapped birthdates vetoed; placeholder members; one-letter
   variants; two-word names; per-parish bulk days. Also found by me while
   checking: chaining through third records → cannot-links.
4. Greedy grouping order orphaning true duplicates (→ strength order and
   the Not Grouped sheet); a couple's single records kept apart; Jr./Sr.
   and unscored pairs not constrained; parents as partners; name gaps.
5. Pairs with no match evidence skipped the related checks (→ verdict
   before the early return; `apart()` trusts only matched pairs); widowed
   rule still too loose; singles a generation apart; tie-breaking.

A sixth review was offered but not run; the round-5 fixes were small and
verified against every reviewer example.

## Latest results (October 12, 2026 exports)

32,056 Families to merge in 23,157 groups (Very high 18,799; High 6,299;
Medium 5,097; Low 1,861). 23,746 pairs kept apart; 282 records on the Not
Grouped sheet. Most common case: moved between parishes (18,360). See the
data directory's `output/summary.md` after a run for current numbers.

## Working on this

- Environment used so far: `~/git/parishkit/venv` (Python 3.14) has
  openpyxl, rapidfuzz, pytest and ruff. Run scripts with `python -I`. In
  this Claude setup `python -c` is blocked; write a small script file
  instead.
- After any rule change:
  1. `python -m pytest`, `python -m ruff check .`,
     `python -m ruff format --check .`
  2. `./run.sh` (about 2 minutes; caches rebuild after edits).
  3. `tools/check_pairs.py ~/archlou-dedup/output/archlou-duplicate-families.xlsx
     ~/archlou-dedup/regression-pairs.csv` — 38 known outcomes from the
     reviews; all pass as of the latest run. Add new cases there (DUIDs and a
     name-free note), never here.
  4. Spot-check changed counts in the Summary, and use
     `tools/explain_pair.py` on surprising pairs.
  5. For significant logic changes, run an independent review agent with
     data access, and give it the requirements above.
- Keep `docs/flowchart.html` in sync with rule changes and headline
  numbers. It is also published as a private Claude artifact
  (https://claude.ai/artifact/9ucJAkyiWCAc4frcGDvF4G); republish to that
  URL after editing.
- Pre-git history: `~/archlou-dedup/` holds earlier snapshots
  (`dedup-family-only.py`, `dedup-v2…v6-*.py`), throwaway analysis
  scripts, and the reviewers' scratch directories. They contain personal
  data in places; keep them out of this repo.
- The input spreadsheets were originally in `~/git/parishkit/` (untracked)
  and later moved; place current exports in the data directory under the
  default names (see README).
