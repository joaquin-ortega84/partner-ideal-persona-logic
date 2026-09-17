"""
Partner Ideal Persona Tagger (simplified)
==========================================

Adds two columns to a contacts dataframe:

  - "Partner Ideal Persona"  -> True / False
  - "Persona Type"           -> "Partner Relationship" | "Client Sales" | "Delivery" | ""

A contact is tagged True for a persona when BOTH are true:
  (a) their Title contains at least one of that persona's keywords, and
  (b) their Job Level is one of that persona's qualifying levels.

If a title matches keywords for more than one persona, the persona is chosen
by priority: Partner Relationship > Client Sales > Delivery.

Job Level model (7 canonical tiers)
------------------------------------
Executive, Vice President, Senior Director, Director, Senior Manager,
Manager, Entry Level (Individual Contributor)

Qualifying levels per persona
------------------------------
  Partner Relationship : Executive, Vice President, Senior Director,
                          Director, Senior Manager, Manager
  Client Sales         : Executive, Vice President, Senior Director,
                          Director, Senior Manager, Manager, Entry Level
  Delivery            : Executive, Vice President, Senior Director, Director

(Adjust PERSONA_JOB_LEVELS below if any of these should be different.)

Usage
-----
    python tag_partner_ideal_persona.py input.xlsx output.xlsx \
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
    "Partner Relationship": [
        "managing director it", "managing director digital workplace",
        "partnership", "partnerships", "alliance",
        "alliances", "partner development", "partner management",
        "alianzas", "ecosystem partnerships", "strategic deals",
        "business relationship", "strategic relationship",
        "strategic supplier relationship management", "partner ecosystem",
        "global offering", "global portfolio", "group offering", "group proposition", "workplace offering"

    ],
    "Client Sales": [
        "services sales", "strategic accounts", "strategic clients", "business development",
        "large deals", "client executive", "client director", "client management",
        "account delivery", "account executive", "account management partner", "key account"
        "client management executive", "client experience executive",
        "client relationship", "client services", "client solutions",
        "client partner", "sales executive", "strategic client", "solutions account",
        "account manager", "account sales", "enterprise business development",
        "enterprise solutions sales", "major accounts",
        "principal client executive", "sales account",
        "strategic account manager", "customer acquisition",
        "desarrollo de negocios", "director comercial", "negocio",
        "ejecutivo de cuentas", "gerente de cuentas", "gerente de negocio",
        "solution offering architect", "deal architect", "service architect"
    ],
    "Delivery": [
        "delivery", "service delivery", "service management", "service desk", "end user services",
        "workplace services", "digital workpace", "delivery service",
        "client delivery", "operational excellence", "platform services", "solution service",
        "platforms", "application delivery", "architecture",
        "customer success", "delivery director", "delivery operations",
        "enterprise and solution architecture", "it service", "itsm service", "cloud service",
        "services and solution", "transformacion", "transformation", "architecture",
        "transformation programs", "solution architect", "enterprise architect", "architect modern workplace", "workspace architect", "chief architect", "delivery architect",
        "portfolio architect", "pre sales architect", "service delivery executive",
    ],
}

# Priority used when a title matches keywords for more than one persona.
PERSONA_PRIORITY = ["Partner Relationship", "Client Sales", "Delivery"]

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
    "product manager"
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
    "consultant": "Entry Level",
}

PERSONA_JOB_LEVELS = {
    "Partner Relationship": {
        "Executive", "Vice President", "Senior Director", "Director",
        "Senior Manager",
    },
    "Client Sales": {
        "Executive", "Vice President", "Senior Director", "Director",
        "Senior Manager", "Manager", "Entry Level",
    },
    "Delivery": {
        "Executive", "Vice President", "Senior Director", "Director", "Senior Manager", "Manager"
    },
}

# For these (persona, job level) combos, a keyword match is not enough on
# its own — the title must ALSO contain the given word. This exists because
# many Client Sales keywords are broad generic role words ("sales manager",
# "account director", "large deals", ...) that are fine signals at
# Manager/Director+ level, but too loose at Entry Level: an Entry Level
# "Sales Manager" or "Account Director" title shouldn't qualify unless the
# title itself says "executive" (e.g. "Sales Executive", "Account Executive").
EXTRA_TITLE_REQUIREMENT = {
    ("Client Sales", "Entry Level"): ["executive", "manager"],
    ("Delivery", "Senior Manager"): [
        "architecture", "transformation", "experience delivery", "euc",
        "principal architect", "principal solutions owner", "experience",
        "modern workplace", "service desk", "architecture"
    ],
    ("Delivery", "Manager"): [
        "architecture", "experience delivery", "euc",
        "principal architect", "principal solutions owner", "experience",
        "modern workplace", "service desk", "architecture"
    ],
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
    parser.add_argument("output_file", nargs="?", help="Path to write tagged .xlsx or .csv")
    parser.add_argument("--title-col", default="Title")
    parser.add_argument("--level-col", default="Job Level")
    parser.add_argument("--id-col", default="Lead or Contact ID",
                         help="Column after which Persona Type / Partner Ideal Persona are inserted")
    args = parser.parse_args()

    if not args.input_file or not args.output_file:
        print("Usage: python tag_partner_ideal_persona.py input.xlsx output.xlsx", file=sys.stderr)
        sys.exit(1)

    process_file(args.input_file, args.output_file, args.title_col, args.level_col, args.id_col)


if __name__ == "__main__":
    main()
