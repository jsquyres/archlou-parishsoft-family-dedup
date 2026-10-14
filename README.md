# Archdiocese ParishSoft Family dedup

`dedup.py` finds ParishSoft Family records that are really the same
household, across all parishes in an Archdiocese-wide export. For each set
of duplicates it picks the record that should survive and produces
(survivor, merged) DUID pairs for ParishSoft's Family merge tool, where the
second Family is merged into the first.

It also keeps apart Families that legitimately share a person but are
different households: divorce and remarriage, an ex-spouse's own household,
a grown child with their own Family, a parent who moved in with a child,
separated parents who share children, and fathers and sons with the same
name.

It is a single-purpose tool for one data cleanup. It is expected to be
rerun many times as the Archdiocese merges records and re-exports.

A visual overview of the pipeline and rules is in
[docs/flowchart.html](docs/flowchart.html).

## Personal data

**This repository contains no personal data, and must never contain any.**
The ParishSoft exports, the script's caches and reports, and any list of
real names or DUIDs live in a data directory outside the git checkout (by
default `~/archlou-dedup`). `.gitignore` blocks spreadsheets, CSVs, caches
and outputs as a safety net. Tests and docs use invented names only.

## Setup

Python 3.12 or newer.

```sh
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt        # runtime: openpyxl, rapidfuzz
pip install -r requirements-dev.txt    # also: pytest, ruff
```

## Inputs

Three ParishSoft Archdiocese-level exports (`.xlsx`, first sheet, one header
row):

| Export | Default file name | Used for |
|---|---|---|
| Families | `archlou-families-list-to-deduplicate.xlsx` | one row per Family: names, Primary address, Home/Primary phone, email, Organization (parish), Is Registered, Group Name, Membership Date, Date Unregistered |
| Members | `archlou-members-list.xlsx` | one row per Member, keyed by Family DUID: formal/nickname/middle/last/maiden names, suffix, birthdate, Member Type, deceased flag, date of death, Member Status, cell phone, email, Date Modified |
| Enhanced Families (optional) | `archlou-families-list-to-deduplicate-enhanced.xlsx` | Family "Date Last Updated" and Emergency Phone (+ description) |

The exact column headers the script reads are the `COLS`, `MEMBER_COLS` and
`ENHANCED_COLS` tables at the top of `dedup.py`. If ParishSoft renames a
column, update those tables.

## Running

```sh
./run.sh
```

`run.sh` reads the exports from `$DATA_DIR` (default `~/archlou-dedup`) and
writes reports and caches to `$DATA_DIR/output`. It refuses a `DATA_DIR`
inside the repository. Override any path with `DATA_DIR`, `FAMILIES`,
`MEMBERS`, `ENHANCED`, `OUT_DIR` or `PYTHON`, for example:

```sh
DATA_DIR=/secure/archlou PYTHON=.venv/bin/python ./run.sh
```

Or call the script directly:

```sh
python -I dedup.py FAMILIES.xlsx MEMBERS.xlsx [--enhanced ENHANCED.xlsx] --out-dir DIR
```

A full run takes about 2 minutes, most of it reading the spreadsheets. The
parsed exports are cached in the output directory (`*.dedup.cache.pickle`)
and reused while both the spreadsheet and `dedup.py` are unchanged, so a
rerun with no changes is faster. Any edit to `dedup.py` rebuilds the caches.

## Output

`DIR/archlou-duplicate-families.xlsx`, plus `DIR/summary.md` (the Summary
sheet as a Markdown table, also printed to the terminal).

| Sheet | Contents |
|---|---|
| Duplicates | One row per record, grouped. Each group starts with its SURVIVE row, then its merge rows. Sorted by group confidence, then the survivor's last name. Each merge row has its own confidence, a reason code and the reasons in words. Includes every Family field the decision used and a Members column (name, type, birthdate, deceased). |
| Merge Pairs | (Survivor DUID, Merged DUID) pairs in ParishSoft's merge order, sorted by each pair's confidence then survivor last name, with a coarse match type. |
| Reasons | Each distinct combination of confidence and reasons, described once, with how many Families it covers. |
| Not Grouped | Records that have a Very high/High/Medium match but are in no group, because every link would have put two Families that were kept apart into one group. Review these by hand. |
| Summary | Totals, how many pairs each rule rejected or set aside, families by confidence and match type, group sizes, and the bulk dates ignored. |

Nothing is merged automatically. The workbook is for the Archdiocese to
review and apply.

## How it works

1. **Load and clean.** Normalize names, nicknames, addresses, phones, ZIPs
   and emails; drop placeholder Families (Loose Collection, Unknown Family,
   Anonymous, certificate catch-alls) and placeholder Members (surname
   typed as the first name, with no birthdate, nickname or specific type).
2. **Pick candidate pairs.** Two Families are compared only if they share an
   email, a phone (home, primary, Member cell, or own emergency phone), a
   street address, surname + ZIP, or a person (given name + exact birthdate
   + surname or maiden name). The last key finds Families that moved.
3. **Score the Family fields.** A base confidence from surname (same or
   similar), first names (same, overlapping or different) and the number of
   shared contacts. Contact values shared by more than three surnames
   (nursing homes, parish offices) don't count.
4. **Check the Members.** In order, a pair is kept apart (never merged and
   never grouped together) if:
   - the shared person is an ex-spouse in one Family;
   - a head/spouse in one Family is a child or other relative in the
     other, and no head/spouse is in both (grown child, parent moved in);
   - the same person has a different living partner in each Family
     (divorce and remarriage);
   - same-named adults have clearly different birthdates, or Jr. vs Sr.;
   - both Families have adults but none in common (different households).
     Exception: two single-adult records that share no one are neither
     merged nor kept apart, so both can join a couple's record, unless the
     two adults are a generation apart.

   Otherwise exact Member matches set the confidence. If one Family has an
   extra living adult, confidence drops one level, unless the other
   Family's spouse is deceased and both spouses are of the shared person's
   generation (widowed and remarried).
5. **Group.** Matches join into groups strongest evidence first. A link is
   skipped if it would put a kept-apart pair, or a pair that fails the
   Jr./Sr. and Member checks when compared directly, into one group.
6. **Choose the survivor.** First difference wins: Registered = Yes, Group
   = Active (a Family that moved survives in the parish where it is
   Active), has a living Member, most recent real activity, most fields
   filled in, lowest DUID.
7. **Write the workbook.**

### Confidence levels

| Level | Typical evidence |
|---|---|
| Very high | Two or more people match by name and exact birthdate, or same surname and first names with shared contact info |
| High | Overlapping first names and one shared contact, or one adult matches exactly with no shared contact |
| Medium | Same names in the same ZIP with no shared contact; a similar surname; only children match exactly |
| Low (review) | Different first names at one address; one Family has an extra adult; linked only through other records in a large group |

A record linked to the survivor only through another record gets that
path's weakest link, one level lower.

### Recent activity

"Most recent activity" is the latest of Membership Date, Date Unregistered,
Family Date Last Updated and any Member Date Modified, ignoring bulk jobs:
any day with 500+ such dates Archdiocese-wide, any day with 100+ in one
parish, and any minute with 200+ Member edits.

## Known limitations

- Records with no birthdates and no shared contact info can't be confirmed.
- A child typed "Unknown" with no birthdates anywhere in the Family is
  treated as an adult when the Family has no explicit couple.
- About 35 records in one parish have birth years that look corrupted
  (1999–2001 vs 1937–43 elsewhere for the same people). They are treated as
  different people, which can push a true duplicate onto the Not Grouped
  sheet.
- Ignoring a trailing street type means "123 Oak St" and "123 Oak Ct" in
  the same ZIP compare as the same address.
- A Low link from an ungrouped record to a record already in a stronger
  group is not used, and that record is not listed anywhere.

## Development

```sh
python -m pytest
python -m ruff check .
python -m ruff format --check .
```

Helper tools, both of which read personal data from the data directory and
print it to the terminal only:

- `tools/explain_pair.py FAMILIES MEMBERS DUID_A DUID_B --cache-dir DIR
  [--enhanced E] [--report R]` shows how two Families are scored and why,
  with their Members and group placement.
- `tools/check_pairs.py REPORT.xlsx EXPECTATIONS.csv` checks a report
  against known outcomes (`duid_a,duid_b,expect,note`, where `expect` is
  `together` or `apart`). The expectations file holds real DUIDs, so it
  lives in the data directory (`~/archlou-dedup/regression-pairs.csv`),
  never in this repository.

See [CLAUDE.md](CLAUDE.md) for the design history and maintainer notes.

## License

BSD 3-Clause; see [LICENSE](LICENSE).
