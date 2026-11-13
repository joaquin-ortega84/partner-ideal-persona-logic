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
    python logic_new.py input.xlsx output.csv \
        --title-col "CM Title" --level-col "LeadorContact.Job_Level" \
        --region-col "Activity Region L2"

Japan fallback
--------------
For rows where Activity Region L2 == "Japan", if no English keyword qualifies,
Japanese title keywords are checked (substring match). Seniority always comes
from the Job Level column. See JAPANESE_PERSONA_KEYWORDS.
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
        "ceo", "cto", "chief executive officer", "chief information officer", "chief technology officer",
        "ciso", "cio", "chief security officer", "cso", "cco", "chief commercial officer",
        "cfo", "chief financial officer", "coo", "chief operating officer", "chief information security officer",
        "chief product officer", "cpo", "chief services officer", "chief strategy officer", "chief digital officer",
        # General & managing leadership
        "general manager services", "general manager it", "general manager digital workplace",
        "managing director", "managing director it", "managing director digital workplace", "managing partner",
        # IT & information systems
        "it", "information technology", "information systems", "technology", "dsi", "informatique", "architecture", "itsm",
        # Strategy
        "it strategy", "strategy and innovation", "it strategy",
        # Transformation & innovation
        "transformation", "transformacion", "innovation", "intelligent automation", "xmo", "modernization",
        # Managed services & infrastructure
        "managed services", "managed infrastructure services", "infrastructure",
        "geschäftsführer", "geschaftsfuhrer", "strategic tooling", "observability",
    ],

    "DEX Leadership": [
        # Digital workplace, end user & employee experience
        "digital workplace", "workplace", "workplace services", "workspace", "dwp", "dws",
        "end user services", "end user computing", "end user devices", "eus",
        "endpoint", "end point", "euc", "compute", "computing", "post de travail",
        "dex", "digital experience", "employee experience", "user experience", "eux", "experience management", "poste de travail",
        "service desk", "desktop", "desktops", "workforce solution",
    ],

    "DEX Managers": [
        # Digital workplace, end user & employee experience
        "digital workplace", "workplace", "workplace services", "workspace", "dwp", "dws",
        "end user services", "end user computing", "end user devices", "eus", "dx", "dxp",
        "endpoint", "end point", "euc", "compute", "computing", "post de travail",
        "dex", "digital experience", "employee experience", "user experience", "eux", "experience management", "poste de travail",
        "service desk", "desktop", "desktops", "workforce solutions",
    ],

    "Alliance & Partnerships": [
        # Alliances & partnerships
        "alliance", "alliances", "alianzas", "partners",
        "partnership", "partnerships", "partner development", "partner management", "partneships",

        # Ecosystem & Nexthink-specific
        "ecosystem partnerships", "partner ecosystem", "nexthink partner", "nexthink partnership",

        # Relationship
        "strategic relationship", "relationship", "strategic relationship", "relationship",
    ],

    "Offering & Portfolio": [
        # Portfolio & offering
        "portfolio", "offering", "enterprise solutions", "portfolios"
    ],

    "Senior Service Delivery": [
        # Delivery & operations
        "delivery", "service delivery", "global services", "global service",
        "operational excellence", "solution service", "services and solution", "service excellence", "operational excellence",
        "observability",

        # Service management & support
        "service management", "service desk", "service desks", "it service", "itsm service", "servicedesk",
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
        "client solutions", "strategic client", "strategic clients",
        # Sales execution
        "account executive", "sales executive", "head of sales", "strategic sales", "services sales",
        "enterprise solutions sales", "large deal",
        # New business & business development
        "business development", "new business", "customer acquisition",
        # Pre-sales
        "presales", "pre sales", "pre-sales",
        # Non-English
        "gerente de cuentas", "ejecutivo de cuentas", "director comercial", "negocio",
        "commericial", "commerciales", "commerciale",
        # Account technical leadership
        "account cto", "technical sales",
    ],

    "Architects": [
        # Pre-sales & solution architects
        "solution architect", "solution offering architect", "deal architect", "pre sales architect",
        "solutions architect",
        # Portfolio & chief architects
        "chief architect", "cheif architect", "enterprise architect", "portfolio architect",
        "architecture", "chief architect", "chief enterprise architect",
        # Domain & technical architects
        "euc architect", "dex architect", "workplace architect", "modern workplace architect",
        "architect modern workplace", "it architect", "technical architect", "service architect",
        "dex architect",
        # Architecture managers
        "architecture"
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
    "DEX Managers",
    "Senior Service Delivery",
    "IT Leadership",
]

# Titles containing any of these phrases are never tagged as an Ideal
# Persona, regardless of persona keyword / job level matches (e.g. an
# "Assistant to the VP of Sales" is not the VP of Sales).
EXCLUSION_KEYWORDS = [
    "intern",
    "supply chain",
    "assistant to",
    "assistant of",
    "ea to",
    "executive assistant",
    "assistant director",
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
# 1b. Japan fallback — Japanese title keywords (from Christine Wu, 9/14/2026)
# ---------------------------------------------------------------------------
# Used ONLY for rows where Activity Region L2 == Japan and no English keyword
# qualified the contact. Seniority still comes from the Job Level column.
# Only Highest and High priority terms are used. Medium terms (DX推進, デジタル戦略,
# 働き方改革, ワークスタイル, デジタル推進, 従業員体験) never qualify on their own,
# per the guidance that they should not qualify automatically.
# "Conditional" leader titles (リーダー, チームリーダー, グループリーダー, 責任者)
# are not persona keywords, so they only count when a term above also matches.
# Matching is substring-based, because Japanese has no word boundaries.

JAPAN_REGION = "Japan"  # compared case-insensitively in is_japan()

JAPANESE_PERSONA_KEYWORDS = {
    "IT Leadership": [
        "情報システム", "情シス", "IT企画", "IT戦略", "IT基盤", "ITインフラ",
        "システム企画", "システム基盤", "社内情報システム", "社内IT", "社内システム",
    ],
    "DEX Leadership": [
        "デジタルワークプレイス", "デジタルワークスペース", "エンドポイント",
        "PC管理", "PC運用", "クライアント管理", "ヘルプデスク", "サービスデスク",
        "ユーザーサポート",
    ],
    "DEX Managers": [
        "デジタルワークプレイス", "デジタルワークスペース", "エンドポイント",
        "PC管理", "PC運用", "クライアント管理", "ヘルプデスク", "サービスデスク",
        "ユーザーサポート",
    ],
    "Senior Service Delivery": [
        "ITサービスマネジメント", "ITサービス", "IT運用", "システム運用",
        "運用管理", "運用保守", "ITサポート",
    ],
}

# Japanese titles that are excluded outright (Christine: "Usually exclude").
#
# Pseudo code:
#   Japanese exclusion list = [
#       "係長" -> Assistant Manager / Section Chief
#       "主任" -> Senior Staff / Supervisor
#   ]
#   For every row where Activity Region L2 is Japan:
#       if the title contains ANY word in the list (plain text search):
#           do not tag as Ideal Persona, stop here
#   The check runs before keyword matching, so these titles are excluded
#   even if an English keyword would otherwise match.
JAPANESE_EXCLUSION_KEYWORDS = ["係長", "主任"]

# Titles containing Japanese characters (hiragana, katakana, kanji).
_JAPANESE_CHARS = re.compile(r"[\u3040-\u30ff\u4e00-\u9fff]")

# Seniority fallback for Japan rows whose Job Level is blank or unrecognized.
# Job Level always wins when it is filled in. This table is used only when it isn't.
#
# Pseudo code:
#   For a Japan row with Job Level = Unknown:
#       Sort the title terms longest first (so 副部長 is checked before 部長).
#       For each term, if the title contains it, the tier is the mapped value.
#       If no term matches, the tier stays Unknown.
#
# Mapping (Japanese title -> canonical tier):
#   執行役員                         -> Executive
#   本部長, 統括部長                  -> Senior Director
#   部長, 担当部長, 室長, ジェネラルマネージャー -> Director
#   副部長, 次長                      -> Senior Manager   (DECISION: confirm)
#   グループマネージャー, グループ長    -> Senior Manager
#   シニアマネージャー                 -> Senior Manager
#   課長, マネージャー                 -> Manager
#   リーダー, チームリーダー, 責任者 etc. -> not mapped (Unknown)
#   係長, 主任                        -> not mapped (excluded above)
JAPANESE_TITLE_TIERS = {
    "執行役員": "Executive",
    "本部長": "Senior Director",
    "統括部長": "Senior Director",
    "部長": "Director",
    "担当部長": "Director",
    "室長": "Director",
    "ジェネラルマネージャー": "Director",
    "副部長": "Senior Manager",
    "次長": "Senior Manager",
    "グループマネージャー": "Senior Manager",
    "グループ長": "Senior Manager",
    "シニアマネージャー": "Senior Manager",
    "課長": "Manager",
    "マネージャー": "Manager",
}

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
    "DEX Managers": {"Senior Manager", "Manager"},
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
# 4. Core tagging new_logic
# ---------------------------------------------------------------------------


def is_japan(region) -> bool:
    if region is None or (isinstance(region, float) and pd.isna(region)):
        return False
    return str(region).strip().lower() == JAPAN_REGION.lower()


def title_has_japanese(title) -> bool:
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    return _JAPANESE_CHARS.search(str(title)) is not None


def japanese_title_matches(title, persona: str) -> bool:
    """Substring match against the Japanese keywords (no word boundaries in Japanese)."""
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    t = str(title)
    return any(kw in t for kw in JAPANESE_PERSONA_KEYWORDS.get(persona, []))


def japanese_title_excluded(title) -> bool:
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return False
    t = str(title)
    return any(kw in t for kw in JAPANESE_EXCLUSION_KEYWORDS)


def japanese_title_level(title) -> str:
    """Seniority tier from a Japanese title, or 'Unknown'. Longest terms are checked first."""
    if title is None or (isinstance(title, float) and pd.isna(title)):
        return "Unknown"
    t = str(title)
    for term in sorted(JAPANESE_TITLE_TIERS, key=len, reverse=True):
        if term in t:
            return JAPANESE_TITLE_TIERS[term]
    return "Unknown"


def tag_contact(title, raw_job_level, region=None) -> tuple:
    """Returns (persona_type: str, is_ideal_persona: bool).

    1. English keywords are checked first, for every region.
    2. If none qualify and the row is Activity Region L2 == Japan, the Japanese
       keywords are checked as a fallback. Seniority always comes from the
       Job Level column.
    """
    if title_excluded(title):
        return "", False
    if is_japan(region) and japanese_title_excluded(title):
        return "", False

    level = normalize_job_level(raw_job_level)

    # Job Level fallback: for Japan rows with no usable Job Level, derive the
    # tier from the Japanese title. A filled-in Job Level is never overridden.
    if level == "Unknown" and is_japan(region):
        level = japanese_title_level(title)

    # 1. English keywords
    for persona in PERSONA_PRIORITY:
        if not title_matches(title, persona) or level not in PERSONA_JOB_LEVELS[persona]:
            continue

        required_words = EXTRA_TITLE_REQUIREMENT.get((persona, level))
        if required_words and not any(title_contains_word(title, w) for w in required_words):
            continue

        return persona, True

    # 2. Japan fallback
    if is_japan(region) and title_has_japanese(title):
        for persona in PERSONA_PRIORITY:
            if persona not in JAPANESE_PERSONA_KEYWORDS:
                continue
            if not japanese_title_matches(title, persona) or level not in PERSONA_JOB_LEVELS[persona]:
                continue
            return persona, True

    return "", False


def tag_dataframe(df: pd.DataFrame, title_col: str, level_col: str, id_col: str = None,
                  region_col: str = None) -> pd.DataFrame:
    has_region = region_col is not None and region_col in df.columns
    results = df.apply(
        lambda row: tag_contact(
            row.get(title_col),
            row.get(level_col),
            row.get(region_col) if has_region else None,
        ),
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


def process_file(input_path: str, output_path: str, title_col: str, level_col: str,
                 id_col: str = None, region_col: str = None) -> None:
    # Input can be .xlsx/.xls or .csv; output is always written as CSV.
    if input_path.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(input_path)
    else:
        df = pd.read_csv(input_path)

    df = tag_dataframe(df, title_col, level_col, id_col, region_col)
    df = df[df[title_col].notna() & (df[title_col].astype(str).str.strip() != "")]

    if not output_path.lower().endswith(".csv"):
        output_path = output_path.rsplit(".", 1)[0] + ".csv"

    df = df[['CM Full Name','CM Title','Ideal Persona', 'BDR IP', 'Current_Account_Status',
            'Activity Region L1', 'Activity Region L2', 'Activity Region L3','Related Record ID',
            'Account Type', 'LeadorContact.Job_Level','BDR IP','Auto IP', 'Persona Type', 'Partner Ideal Persona']]

    df = df.drop_duplicates()

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
    parser.add_argument("--title-col", default="CM Title")
    parser.add_argument("--level-col", default="LeadorContact.Job_Level")
    parser.add_argument("--id-col", default="Lead or Contact ID",
                         help="Column after which Persona Type / Partner Ideal Persona are inserted")
    parser.add_argument("--region-col", default="Activity Region L2",
                         help="Column used for the Japan fallback (Japan rows only)")
    args = parser.parse_args()

    if not args.input_file or not args.output_file:
        print("Usage: python logic_new.py input.xlsx output.csv", file=sys.stderr)
        sys.exit(1)

    process_file(args.input_file, args.output_file, args.title_col, args.level_col,
                 args.id_col, args.region_col)


if __name__ == "__main__":
    main()
