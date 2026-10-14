"""Unit tests for dedup.py's normalization and matching rules.

All names, dates and addresses here are invented; this repository must never
contain personal data from the ParishSoft exports.
"""

import datetime

import pytest

import dedup


def member(first, last, birth=None, *, role="partner", suffix="", deceased=False, middle="", maiden=""):
    """Build a member dict shaped like load_members' output (invented data)."""
    firsts = dedup.first_names(first)
    d = datetime.date.fromisoformat(birth) if birth else None
    return {
        "duid": f"{first}-{last}-{birth}",
        "type": "",
        "role": role,
        "guessed": False,
        "firsts": firsts,
        "aliases": firsts | dedup.first_names(middle),
        "last": dedup.norm_last(last),
        "lasts": frozenset(x for x in (dedup.norm_last(last), dedup.norm_last(maiden)) if x),
        "birth": d,
        "precise": bool(d) and not (d.month == 1 and d.day == 1),
        "suffix": suffix,
        "deceased": deceased,
        "email": "",
        "cell": "",
        "modified": None,
        "label": first,
    }


@pytest.mark.parametrize(
    "a, b, expected",
    [
        ("bill", "william", True),
        ("pat", "patricia", True),
        ("pat", "patrick", True),
        ("patricia", "patrick", False),
        ("juan", "juana", False),
        ("alejandro", "alejandra", False),
        ("ann", "anna", True),
        ("ana", "anna", True),
        ("jamie", "james", False),
        ("terry", "teresa", True),
        ("joseph", "joan", False),
        ("joan", "joanie", True),
        ("katherine", "kathryn", True),
        ("mary helen", "mary margie", False),
        ("mary helen", "mary", True),
        ("la shauna", "lashauna", True),
        ("mary lou", "mary louise", True),
        ("ma guadalupe", "maria guadalupe", True),
    ],
)
def test_name_match(a, b, expected):
    """Given-name equivalence: nicknames, variants, and gendered look-alikes."""
    assert dedup.name_match(a, b) is expected
    assert dedup.name_match(b, a) is expected


def test_first_names_splits_couples_and_drops_initials():
    """A Family first-names field yields one name per person."""
    assert dedup.first_names("Chick & Kathy") == {"chick", "kathy"}
    assert dedup.first_names("J. Robert and Mary Ann") == {"robert", "mary ann"}
    assert dedup.first_names("No Name") == frozenset()


@pytest.mark.parametrize(
    "a, b, same",
    [
        (("126 N. Birchwood Avenue", None), ("126 N Birchwood", None), True),
        (("123 North Main Street", "Apt. 4"), ("123 N Main #4", None), True),
        (("P. O. Box 12", None), ("PO Box 12", None), True),
        (("100 Saint Matthews Ave", None), ("100 St Matthews", None), True),
        (("603 Lane Ave", None), ("603 Court St", None), False),
        (("12 Oak Dr", "Unit 1"), ("12 Oak Dr", "Unit 2"), False),
    ],
)
def test_norm_addr(a, b, same):
    """Street normalization: types, directionals, units, Saint/P.O. Box."""
    assert (dedup.norm_addr(*a) == dedup.norm_addr(*b)) is same


def test_norm_addr_requires_a_number():
    """Placeholders like 'Unknown' are not addresses."""
    assert dedup.norm_addr("Unknown", None) == ""


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("(502) 459-9681", "5024599681"),
        ("1-502-459-9681", "5024599681"),
        ("502-459-9681 x12", "5024599681"),
        ("502-555-1234", ""),
        ("UNLISTED", ""),
        ("459-9681", ""),
    ],
)
def test_norm_phone(raw, expected):
    """Phones: 10 digits, country code and extension ignored, fakes dropped."""
    assert dedup.norm_phone(raw) == expected


@pytest.mark.parametrize("raw, expected", [(40213, "40213"), ("40213-1463", "40213"), (2134, "02134")])
def test_norm_zip(raw, expected):
    """ZIPs: 5 digits, leading zeros restored."""
    assert dedup.norm_zip(raw) == expected


def test_norm_last_and_last_relation():
    """Surnames: apostrophes, suffixes, compound names, particles."""
    assert dedup.norm_last("O'Brien Jr.") == "obrien"
    assert dedup.last_relation("smith-jones", "smith") == "similar"
    assert dedup.last_relation("van dyke", "van horn") is None
    assert dedup.last_relation("estate of smith", "estate of brown") is None


@pytest.mark.parametrize(
    "b1, b2, expected",
    [
        ("1954-06-15", "1954-06-15", "strong"),
        ("1954-06-15", "1954-06-16", "year"),  # mistyped day
        ("1954-06-15", "1945-06-15", "year"),  # swapped year digits
        ("1954-06-15", "1955-06-15", "year"),  # one year digit
        ("1954-06-15", "1979-06-15", "conflict"),  # father and son sharing a birthday
        ("1954-06-15", "1960-02-03", "conflict"),
        ("1954-01-01", "1954-07-04", "year"),  # Jan 1 is a year-only placeholder
        ("1954-06-15", None, "name"),
    ],
)
def test_member_status_birthdates(b1, b2, expected):
    """Same name: birthdates decide same person, typo, or different people."""
    assert dedup.member_status(member("John", "Doe", b1), member("John", "Doe", b2)) == expected


def test_member_status_suffix_and_maiden():
    """Jr./Sr. are different people; a maiden name still matches."""
    assert dedup.member_status(member("John", "Doe", suffix="jr"), member("John", "Doe", suffix="sr")) == "conflict"
    wed = member("Ann", "Doe", "1970-03-04", maiden="Roe")
    assert dedup.member_status(wed, member("Ann", "Roe", "1970-03-04")) == "strong"


def family(first, last, members):
    """Minimal Family dict for member_evidence (invented data)."""
    return {"n_first": dedup.first_names(first), "n_last": dedup.norm_last(last), "members": members, "catchall": False}


def test_divorce_and_remarriage_is_kept_apart():
    """Same man, a different living wife in each Family."""
    f = family("John & Mary", "Doe", [member("John", "Doe", "1960-05-01"), member("Mary", "Doe", "1962-07-02")])
    g = family("John & Susan", "Doe", [member("John", "Doe", "1960-05-01"), member("Susan", "Doe", "1965-09-03")])
    assert "divorce" in dedup.member_verdict(dedup.member_evidence(f, g))


def test_widowed_and_remarried_is_a_duplicate():
    """Same man; the first wife is deceased; the new wife is his generation."""
    f = family(
        "John & Mary", "Doe", [member("John", "Doe", "1960-05-01"), member("Mary", "Doe", "1962-07-02", deceased=True)]
    )
    g = family("John & Susan", "Doe", [member("John", "Doe", "1960-05-01"), member("Susan", "Doe", "1965-09-03")])
    ev = dedup.member_evidence(f, g)
    assert dedup.member_verdict(ev) is None
    assert dedup.widowed(ev)


def test_grown_child_is_kept_apart():
    """A child in the parents' Family heads their own Family."""
    f = family(
        "Tom & Rita",
        "Doe",
        [
            member("Tom", "Doe", "1950-01-02"),
            member("Rita", "Doe", "1952-02-03"),
            member("Sam", "Doe", "1980-04-05", role="child"),
        ],
    )
    g = family("Sam", "Doe", [member("Sam", "Doe", "1980-04-05")])
    assert "grown child" in dedup.member_verdict(dedup.member_evidence(f, g))


def test_spouse_recorded_under_maiden_name_is_present():
    """The same wife under two surnames is not an extra adult."""
    f = family("Mark & Teri", "Doe", [member("Mark", "Doe", "1970-01-02"), member("Teri", "Doe", "1971-03-04")])
    g = family("Mark & Terri", "Doe", [member("Mark", "Doe", "1970-01-02"), member("Terri", "Lowe", "1971-03-04")])
    ev = dedup.member_evidence(f, g)
    assert ev["extra_f"] == ev["extra_g"] == 0
    assert dedup.member_verdict(ev) is None


def test_build_groups_never_joins_a_kept_apart_pair():
    """A and B are kept apart; both match C; C joins only the stronger link."""
    fams = [
        {
            "duid": d,
            "registered": "No",
            "group": "Inactive",
            "living": True,
            "activity": datetime.datetime.min,
            "filled": 1,
        }
        for d in (1, 2, 3)
    ]
    edges = {(0, 2): (1, ("strong",)), (1, 2): (3, ("weak",))}
    groups, _, blocked = dedup.build_groups(fams, edges, {(0, 1)}, lambda i, j: False)
    assert len(groups) == 1 and blocked == 1
    members = {groups[0]["survivor"]} | {m for m, _, _ in groups[0]["members"]}
    assert members == {0, 2}
