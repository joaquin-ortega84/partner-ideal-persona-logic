"""
Partner Ideal Persona Tagger
============================

Adds two columns to a contacts dataframe:

  - "Partner Ideal Persona"  -> True / False
  - "Persona Type"           -> "IT Leadership" | "DEX Leadership" | "Alliance & Partnerships" |
                                "Offering & Portfolio" | "Senior Service Delivery" |
                                "Client Sales" | "Architects" | ""

A contact is tagged True for a persona when BOTH are true:
  (a) their Title contains at least one of that persona's keywords, and
  (b) their Job Level is one of that persona's qualifying levels.

If a title matches keywords for more than one persona, the persona is chosen
by PERSONA_PRIORITY (most specific personas first, broadest keywords last).

Job Level model (7 canonical tiers)
------------------------------------
Executive, Vice President, Senior Director, Director, Senior Manager,
Manager, Entry Level (Individual Contributor)

Qualifying levels per persona
------------------------------
  IT Leadership           : Executive, Vice President, Senior Director, Director
  DEX Leadership          : Executive, Vice President, Senior Director, Director
  Alliance & Partnerships : Executive, Vice President, Senior Director, Director,
                            Senior Manager, Manager
  Offering & Portfolio    : Executive, Vice President, Senior Director, Director
  Senior Service Delivery : Executive, Vice President, Senior Director, Director
  Client Sales            : Executive, Vice President, Senior Director, Director,
                            Senior Manager, Manager, Entry Level
  Architects              : Entry Level

(Adjust PERSONA_JOB_LEVELS below if any of these should be different.)

Usage
-----
    python logic.py input.xlsx output.csv \
        --title-col Title --level-col "Job Level"
"""

import argparse
import re
import sys
import pandas as pd

# ---------------------------------------------------------------------------
# 1. Keyword map — one list of title keywords per persona
# ---------------------------------------------------------------------------

PERSONA_KEYWORDS = {
    "IT Leadership": [
        # C-suite
        "ceo", "cto", "chief executive office", "chief information officer", "chief technology officer",
        "ciso", "cio", "chief security officer", "cso", "cco", "chief commercial officer",
        "cfo", "chief financial officer", "coo", "chief operating officer",

        # General & managing leadership
        "general manager services", "general manager it", "general manager digital workplace",
        "managing director", "managing director it", "managing director digital workplace", "managing partner",

        # IT & information systems
        "it", "information technology", "information systems", "technology", "dsi", "informatique", "architecture",

        # Strategy
        "it strategy", "stategy and innovation", "it strategy",

        # Transformation & innovation
        "transformation", "transformacion", "innovation", "intelligent automation", "xmo",

        # Managed services & infrastructure
        "managed services", "managed infrastructure services", "infrastructure",
    ],

    "DEX Leadership": [
        # Digital workplace, end user & employee experience
        "digital workplace", "workplace", "workplace services", "workspace", "dwp", "dws",
        "end user services", "end user computing", "end user devices", "eus",
        "endpoint", "end point", "euc", "compute", "computing", "post de travail",
        "dex", "digital experience", "employee experience", "user experience", "eux", "experience management",
    ],

    "Alliance & Partnerships": [
        # Alliances & partnerships
        "alliance", "alliances", "alianzas", "partners",
        "partnership", "partnerships", "partner development", "partner management",

        # Ecosystem & Nexthink-specific
        "ecosystem partnerships", "partner ecosystem", "nexthink partner", "nexthink partnership",

        # Relationship
        "strategic relationship", "relationship", "strategic relationship", "relationship",
    ],

    "Offering & Portfolio": [
        # Portfolio & offering
        "portfolio", "offering", "enterprise solutions",
    ],

    "Senior Service Delivery": [
        # Delivery & operations
        "delivery", "service delivery", "global services", "global service",
        "operational excellence", "solution service", "services and solution",

        # Service management & support
        "service management", "service desk", "service desks", "it service", "itsm service",
        "it support", "support services",

        # Applications
        "application", "application services",
    ],

    "Client Sales": [
        # Account management
        "account manager", "account management partner", "key account", "account director",
        "strategic account", "major account", "account delivery",
        "account sales", "sales account", "solutions account",

        # Client management
        "client partner", "client executive", "client director", "client partner", "client management",
        "client experience executive", "client relationship", "client services",
        "client solutions", "strategic client",

        # Sales execution
        "account executive", "sales executive", "head of sales", "strategic sales", "services sales",
        "enterprise solutions sales", "large deal",

        # New business
        "business development", "new business", "customer acquisition",

        # Pre-sales
        "presales", "pre sales", "pre-sales",

        # Non-English
        "gerente de cuentas", "ejecutivo de cuentas", "director comercial", "negocio",
        "commericial", "commerciales",

        # Account technical leadership
        "account cto",
    ],

    "Architects": [
        # Pre-sales & solution architects
        "solution architect", "solution offering architect", "deal architect", "pre sales architect",
        "solution architect",

        # Portfolio & chief architects
        "chief architect", "cheif architect", "enterprise architect", "portfolio architect",
        "architecture", "chief architect",

        # Domain & technical architects
        "euc architect", "dex architect", "workplace architect", "modern workplace architect",
        "architect modern workplace", "it architect", "technical architect", "service architect",
        "dex architect",
    ],
}

# Priority used when a title matches keywords for more than one persona:
# most specific personas first, broadest keywords ("it", "technology",
# "workplace", "delivery") last.
PERSONA_PRIORITY = [
    "Architects",
    "Client Sales",
    "Alliance & Partnerships",
    "Offering & Portfolio",
    "DEX Leadership",
    "Senior Service Delivery",
    "IT Leadership",
]

# Titles containing any of these phrases are never tagged as an Ideal
# Persona, regardless of persona keyword / job level matches (e.g. an
# "Assistant to the VP of Sales" is not the VP of Sales).
EXCLUSION_KEYWORDS = [
    "assistant to",
    "assistant of",
    "ea to",
    "executive assistant",
    "sales operations",
    "financial services",
    "sales & markeing", "sales and marketing", "customer success",
    "product manager",
    "inside sales", "junior manager", "junior executive",
    "erp", "inside",
    "merchandising",
    "security presales", "data center", "agent", "admin", "technician", "supervisor", "coordinator", "professional", "specialist",
    "assistant manager", "hr transformation", "cyber security", "engineer", "analyst"

]

# ---------------------------------------------------------------------------
# 2. Job level model — 7 canonical tiers + qualifying sets per persona
# ---------------------------------------------------------------------------

# Raw values seen in data -> canonical tier name.
LEVEL_ALIASES = {
    "executive": "Executive",
    "vice president": "Vice President",
    "vp": "Vice President",
    "senior director": "Senior Director",
    "sr director": "Senior Director",
    "sr. director": "Senior Director",
    "director": "Director",
    "senior manager": "Senior Manager",
    "sr manager": "Senior Manager",
    "sr. manager": "Senior Manager",
    "manager": "Manager",
    "entry level": "Entry Level",
    "individual contributor": "Entry Level",
    "ic": "Entry Level",
}

DIRECTOR_AND_ABOVE = {"Executive", "Vice President", "Senior Director", "Director"}

PERSONA_JOB_LEVELS = {
    "IT Leadership": DIRECTOR_AND_ABOVE,
    "DEX Leadership": DIRECTOR_AND_ABOVE,
    "Alliance & Partnerships": DIRECTOR_AND_ABOVE | {"Senior Manager", "Manager"},
    "Offering & Portfolio": DIRECTOR_AND_ABOVE,
    "Senior Service Delivery": DIRECTOR_AND_ABOVE,
    "Client Sales": DIRECTOR_AND_ABOVE | {"Senior Manager", "Manager", "Entry Level"},
    "Architects": {"Entry Level"},
}

# For these (persona, job level) combos, a keyword match is not enough on
# its own — the title must ALSO contain one of the given words.
#
# The framework qualifies Client Sales at Entry Level (Individual Contributor)
# without a condition, so the previous rule is switched off. To restore it
# (only Entry Level titles containing "executive" / "exec", e.g. Account
# Executive), uncomment the Client Sales line below.

SA_MANAGER_REQUIRED_WORDS = [
    "alliance", "alliances", "alianzas", "partners",
    "partnership", "partnerships", "partner development", "partner management",
    "ecosystem partnerships", "partner ecosystem", "nexthink partner", "nexthink partnership",
    "strategic relationship", "relationship", "account cto"
]

EXTRA_TITLE_REQUIREMENT = {
    # ("Client Sales", "Entry Level"): ["executive", "exec"],
    ("Alliance & Partnerships", "Senior Manager"): SA_MANAGER_REQUIRED_WORDS,
    ("Alliance & Partnerships", "Manager"): SA_MANAGER_REQUIRED_WORDS,
}

def normalize_job_level(raw_level) -> str:
    """Map a raw Job Level string to one of the 7 canonical tiers, or 'Unknown'."""
    if raw_level is None or (isinstance(raw_level, float) and pd.isna(raw_level)):
        return "Unknown"
    s = str(raw_level).strip().lower()
    if s in ("", "nan", "none", "(blank)"):
        return "Unknown"
    return LEVEL_ALIASES.get(s, "Unknown")


# ---------------------------------------------------------------------------
# 3. Keyword matching
# ---------------------------------------------------------------------------

_COMPILED_KEYWORDS = {
    persona: [re.compile(r"\b" + re.escape(kw.lower()) + r"\b") for kw in kws]
    for persona, kws in PERSONA_KEYWORDS.items()
}

_COMPILED_EXCLUSIONS = [
    re.compile(r"\b" + re.escape(kw.lower()) + r"\b") for kw in EXCLUSION_KEYWORDS
]


def title_matches(title, persona: str) -> bool:
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    t = str(title).lower()
    return any(pattern.search(t) for pattern in _COMPILED_KEYWORDS[persona])


def title_excluded(title) -> bool:
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    t = str(title).lower()
    return any(pattern.search(t) for pattern in _COMPILED_EXCLUSIONS)


def title_contains_word(title, word: str) -> bool:
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    t = str(title).lower()
    return re.search(r"\b" + re.escape(word.lower()) + r"\b", t) is not None


# ---------------------------------------------------------------------------
# 4. Core tagging logic
# ---------------------------------------------------------------------------


def tag_contact(title, raw_job_level) -> tuple:
    """Returns (persona_type: str, is_ideal_persona: bool)."""
    if title_excluded(title):
        return "", False

    level = normalize_job_level(raw_job_level)

    for persona in PERSONA_PRIORITY:
        if not title_matches(title, persona) or level not in PERSONA_JOB_LEVELS[persona]:
            continue

        required_words = EXTRA_TITLE_REQUIREMENT.get((persona, level))
        if required_words and not any(title_contains_word(title, w) for w in required_words):
            continue

        return persona, True

    return "", False


def tag_dataframe(df: pd.DataFrame, title_col: str, level_col: str, id_col: str = None) -> pd.DataFrame:
    results = df.apply(
        lambda row: tag_contact(row.get(title_col), row.get(level_col)),
        axis=1,
        result_type="expand",
    )
    df = df.copy()
    df["Persona Type"] = results[0]
    df["Partner Ideal Persona"] = results[1]

    # Move the two new columns to sit right after the ID column, if present.
    if id_col and id_col in df.columns:
        cols = [c for c in df.columns if c not in ("Persona Type", "Partner Ideal Persona")]
        insert_at = cols.index(id_col) + 1
        cols[insert_at:insert_at] = ["Persona Type", "Partner Ideal Persona"]
        df = df[cols]

    return df


# ---------------------------------------------------------------------------
# 5. File I/O + CLI
# ---------------------------------------------------------------------------


def process_file(input_path: str, output_path: str, title_col: str, level_col: str, id_col: str = None) -> None:
    # Input can be .xlsx/.xls or .csv; output is always written as CSV.
    if input_path.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(input_path)
    else:
        df = pd.read_csv(input_path)

    df = tag_dataframe(df, title_col, level_col, id_col)

    if not output_path.lower().endswith(".csv"):
        output_path = output_path.rsplit(".", 1)[0] + ".csv"

    df.to_csv(output_path, index=False)

    total = len(df)
    tagged = int(df["Partner Ideal Persona"].sum())
    print(f"Processed {total} contacts -> {tagged} tagged as Partner Ideal Persona.")
    print(df["Persona Type"].replace("", "None").value_counts())
    print(f"\nOutput written to: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Tag contacts with the Partner Ideal Persona model.")
    parser.add_argument("input_file", nargs="?", help="Path to input .xlsx or .csv")
    parser.add_argument("output_file", nargs="?", help="Path to write tagged output (always CSV)")
    parser.add_argument("--title-col", default="Title")
    parser.add_argument("--level-col", default="Job Level")
    parser.add_argument("--id-col", default="Lead or Contact ID",
                         help="Column after which Persona Type / Partner Ideal Persona are inserted")
    args = parser.parse_args()

    if not args.input_file or not args.output_file:
        print("Usage: python logic.py input.xlsx output.csv", file=sys.stderr)
        sys.exit(1)

    process_file(args.input_file, args.output_file, args.title_col, args.level_col, args.id_col)


if __name__ == "__main__":
    main()
