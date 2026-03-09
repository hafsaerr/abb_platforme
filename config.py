# config.py - ISIN Codes and Application Configuration

# ─── ABB ISIN CODES ───────────────────────────────────────────────────────────

OCT_QUOTIDIEN = [
    "MA0000038960", "MA0000040396", "MA0000040768", "MA0000041717",
    "MA0000037616", "MA0000041394", "MA0000042152", "MA0000037962",
    "MA0000038002", "MA0000036261", "MA0000038754", "MA0000040024",
    "MA0000037715", "MA0000037624", "MA0000038655",
]

OMLT_QUOTIDIEN = [
    "MA0000042186", "MA0000041329", "MA0000041261", "MA0000040214",
    "MA0000038978", "MA0000038309", "MA0000038267", "MA0000038200",
    "MA0000030785", "MA0000035917", "MA0000036915", "MA0000030280",
    "MA0000040016", "MA0000042210", "MA0000039695",
]

DIVERSIFIE_QUOTIDIEN = [
    "MA0000030470", "MA0000042202", "MA0000040065", "MA0000038986",
    "MA0000038358", "MA0000038259", "MA0000038077", "MA0000030520",
    "MA0000036501",
]

OMLT_HEBDOMADAIRE = [
    "MA0000042079", "MA0000041170", "MA0000041014", "MA0000039190",
    "MA0000037475", "MA0000037087", "MA0000035099",
]

DIVERSIFIE_HEBDOMADAIRE = [
    "MA0000042087", "MA0000042004", "MA0000041725", "MA0000039554",
    "MA0000038408", "MA0000037665", "MA0000037640", "MA0000036634",
    "MA0000036782", "MA0000039398",
]

# No weekly OCT
OCT_HEBDOMADAIRE = []

# ─── LOOKUP TABLES ─────────────────────────────────────────────────────────────

CATEGORY_CODES = {
    "OCT": {
        "quotidien": OCT_QUOTIDIEN,
        "hebdomadaire": OCT_HEBDOMADAIRE,
    },
    "OMLT": {
        "quotidien": OMLT_QUOTIDIEN,
        "hebdomadaire": OMLT_HEBDOMADAIRE,
    },
    "DIVERSIFIE": {
        "quotidien": DIVERSIFIE_QUOTIDIEN,
        "hebdomadaire": DIVERSIFIE_HEBDOMADAIRE,
    },
}

ALL_ABB_CODES = set(
    OCT_QUOTIDIEN + OMLT_QUOTIDIEN + DIVERSIFIE_QUOTIDIEN
    + OMLT_HEBDOMADAIRE + DIVERSIFIE_HEBDOMADAIRE
)

# ─── COLUMN ALIASES (flexible detection) ───────────────────────────────────────

COLUMN_ALIASES = {
    "isin": [
        "Code ISIN", "code isin", "ISIN", "CodeISIN", "code_isin",
        "Code Isin", "CODE ISIN",
    ],
    "opcvm": [
        "OPCVM", "opcvm", "Libellé", "Libelle", "Nom OPCVM",
        "Fonds", "NOM", "Nom", "LIBELLE", "Libellé OPCVM",
    ],
    "societe_gestion": [
        "Société de Gestion", "Societe de Gestion", "société de gestion",
        "SGS", "Gestionnaire", "Société de gestion", "SOCIETE DE GESTION",
        "Sté de Gestion",
    ],
    "souscripteurs": [
        "Souscripteurs", "Souscripteur", "SOUSCRIPTEURS",
    ],
    "periodicite": [
        "Périodicité VL", "Periodicite VL", "Périodicité",
        "Periodicite", "PERIODICITE VL", "Périodicité de la VL",
    ],
    "actif_net": [
        "Actif Net", "AN", "actif net", "ACTIF NET",
        "Actif Net (MAD)", "AN (MAD)",
    ],
    "classification": [
        "Classification", "Catégorie", "Categorie",
        "CLASSIFICATION", "Type", "TYPE",
    ],
    "perf_1j": [
        "1J", "Perf 1J", "Performance 1J", "1 jour", "1 Jour",
        "Perf. 1J", "Performance quotidienne", "PERF 1J",
        "Variation 1J", "1J (%)", "Variation J-1",
        "Perf quotidienne", "PERF QUOTIDIENNE", "perf quotidienne",
        "Perf. quotidienne", "PERF. QUOTIDIENNE",
    ],
    "perf_1s": [
        "1S", "1 semaine", "1 Semaine", "Perf 1S", "Perf hebdomadaire",
        "Performance hebdomadaire", "PERF HEBDOMADAIRE", "Perf. hebdomadaire",
        "Variation 1S", "1S (%)", "Variation semaine", "PERF 1S",
    ],
    "ytd": [
        "YTD", "Perf YTD", "Performance YTD", "YTD (%)",
        "Depuis début année", "Depuis Début Année", "PERF YTD",
        "Perf. YTD", "YTD%",
    ],
}

# ─── CLASSIFICATION KEYWORDS ────────────────────────────────────────────────────

CLASSIFICATION_KEYWORDS = {
    "OCT": ["oct", "obligataire court terme", "trésorerie", "tresorerie", "monetaire", "monétaire"],
    "OMLT": ["omlt", "obligataire moyen", "obligataire long", "mlt", "oblig mlt"],
    "DIVERSIFIE": ["diversifié", "diversifie", "diversified", "balance", "équilibre", "equilibre"],
}

# ─── COLORS ─────────────────────────────────────────────────────────────────────

COLORS = {
    "primary":    "#C8501E",   # Albarid Bank warm orange
    "primary_lt": "#F5C6A8",   # Light orange
    "header_bg":  "#C8501E",
    "header_fg":  "#FFFFFF",
    "positive":   "#92D050",
    "negative":   "#FF6B6B",
    "neutral":    "#F9F9F9",
    "border":     "#CCCCCC",
    "text_dark":  "#1A1A2E",
    "sidebar_bg": "#1A1A2E",
}
