import re
"""
config_google_framework.py
22 controls: F001–F022
Binary methodology: OK or FLAG only — no PARTIAL.
Manual controls (F006, F008, F013–F016, F019–F022) always return OK with reviewer note.
"""
from __future__ import annotations
from config import STATUS_OK, STATUS_FLAG, STATUS_PARTIAL, ControlResult

PILLAR = "google_framework"

# No controls excluded from scoring — manual ones always return OK so they never penalise
SCORING_EXCLUDED: set = set()

PRIORITY_POINTS = {
    10: -18, 9: -15, 8: -13, 7: -11,
     6:  -9, 5:  -7, 4:  -5, 3:  -3, 2: -2, 1: 0,
}

IMPACT_LABEL = {
    10: "Critical", 9: "Critical", 8: "High", 7: "High",
     6: "Medium",   5: "Medium",  4: "Low",  3: "Low",
     2: "Visibility", 1: "Visibility",
}

# Matches Importance column (M) in Framework_Reference template
IMPORTANCE = {
    "F001": 8,   # Naming — QT Prefix
    "F002": 8,   # Naming — Campaign Type Token
    "F003": 8,   # Naming — Strategy Tag
    "F004": 6,   # Display Expansion Disabled
    "F005": 5,   # Promotion End Dates
    "F006": 8,   # PMAX Automation Settings (manual → always OK)
    "F007": 8,   # Match Type — BROAD Dominance
    "F008": 8,   # TM Terms Uploaded (manual proxy → always OK)
    "F009": 10,  # Branded Search Campaign Active
    "F010": 8,   # Search Term Waste
    "F011": 7,   # Negative Keyword Coverage
    "F012": 8,   # Budget Concentration — PMAX Dominance
    "F013": 6,   # Ad Extensions — Sitelinks (manual → always OK)
    "F014": 4,   # Ad Extensions — Structured Snippets (manual → always OK)
    "F015": 4,   # Ad Extensions — Callouts (manual → always OK)
    "F016": 4,   # Business Name Configured (manual → always OK)
    "F017": 6,   # Logos Approved
    "F018": 4,   # Ad Strength
    "F019": 5,   # Keyword Expander (manual → always OK)
    "F020": 6,   # PMAX Audience Requirements (manual → always OK)
    "F021": 5,   # PMAX Search Themes (manual → always OK)
    "F022": 8,   # Match Type — EXACT Coverage
}

CONTROL_NAMES = {
    "F001": "Naming Convention — QT Prefix",
    "F002": "Naming Convention — Campaign Type Token",
    "F003": "Naming Convention — Strategy Tag",
    "F004": "Display Expansion Disabled on Search",
    "F005": "Promotion End Dates",
    "F006": "PMAX Automation Settings Disabled",
    "F007": "Match Type Governance — BROAD Dominance",
    "F008": "TM Terms Uploaded to QT Portal",
    "F009": "Branded Search Campaign Active",
    "F010": "Search Term Waste",
    "F011": "Negative Keyword Coverage",
    "F012": "Budget Concentration — PMAX Dominance",
    "F013": "Ad Extensions — Sitelinks",
    "F014": "Ad Extensions — Structured Snippets",
    "F015": "Ad Extensions — Callouts",
    "F016": "Business Name Configured",
    "F017": "Logos Approved",
    "F018": "Ad Strength",
    "F019": "Quartile Keyword Expander Enabled",
    "F020": "PMAX Audience Requirements",
    "F021": "PMAX Search Themes",
    "F022": "Match Type Governance — EXACT Coverage",
}

WHY = {
    "F001": "Non-QT_ campaigns are unmanaged or legacy — outside CoE governance and audit traceability.",
    "F002": "Invalid campaign type tokens break automated parsing and make campaign type segmentation unreliable.",
    "F003": "Invalid strategy tags prevent performance-based campaign segmentation and DPL targeting by tier.",
    "F004": "Display expansion routes Search budget to placements with fundamentally different intent signals — inflating impressions while suppressing CVR.",
    "F005": "Expired promotional campaigns waste budget on irrelevant messaging and risk policy violations.",
    "F006": "PMAX automations override QT optimization logic with uncontrolled asset and bid changes.",
    "F007": "A BROAD-dominated keyword mix (>80%) gives the bidding system minimal query-level constraints, exposing the account to irrelevant traffic at scale.",
    "F008": "Without TM terms in the portal, branded query routing cannot be governed — competitor ads can capture high-intent branded searches.",
    "F009": "Without a branded campaign, TM terms are unprotected and brand-driven conversions are unattributed.",
    "F010": "Unmanaged waste terms signal missing negatives and direct budget leakage into non-converting queries.",
    "F011": "Accounts with zero keyword negatives have completely uncontrolled query routing — every irrelevant query competes for budget.",
    "F012": "PMAX dominance (>85%) without Search means a black-box account with no query-level visibility, no negative governance surface, and no branded query protection.",
    "F013": "Fewer than 4 sitelinks reduce ad real estate and limit navigation for high-intent users.",
    "F014": "Missing snippets reduce ad completeness and relevance signals for category-based queries.",
    "F015": "Missing callouts leave available ad messaging space unused, reducing competitive ad strength.",
    "F016": "Missing business name reduces ad credibility and trust signals for new users.",
    "F017": "Missing approved image assets prevent PMAX from serving Display and YouTube ad formats — eliminating upper-funnel inventory.",
    "F018": "An account where all rated assets show LOW has insufficient creative variety — Google deprioritizes these asset groups in auctions.",
    "F019": "Keyword Expander disabled removes QT's automated Search expansion advantage, limiting keyword growth.",
    "F020": "PMAX without audience signals operates in pure query-matching mode with no behavioral targeting layer.",
    "F021": "Fewer than 10 Search Themes gives Google insufficient intent signal to guide PMAX toward relevant queries.",
    "F022": "Without sufficient EXACT coverage (<10%), the account cannot enforce precise query control on high-value, high-CVR terms.",
}

WYSD = {
    "F001": "Rename all non-QT_ campaigns to follow the QT_ naming standard or pause/remove legacy campaigns.",
    "F002": "Rename campaigns to use a valid campaign type token as the second component (e.g., QT_Pmax_, QT_Search_, QT_Shopping_).",
    "F003": "Rename campaigns to use a valid strategy tag as the third component (e.g., General, TopProducts, Zombie, Suppression).",
    "F004": "Disable Display Network expansion on all Search campaigns: Campaign Settings > Networks > uncheck 'Display Network'.",
    "F005": "Pause or end promotional campaigns whose event window has passed. Set campaign end dates to match the actual promotion period.",
    "F006": "Disable auto-created assets, Final URL expansion, and store goals auto-apply for all active PMAX campaigns.",
    "F007": "Add EXACT and PHRASE match keywords to reduce BROAD dominance below 80%. Prioritise top-converting search terms from the Search Terms report.",
    "F008": "Log into QT Portal > Google > Branded Terms and confirm all trademark terms are uploaded and current.",
    "F009": "Ensure a dedicated branded Search campaign (TM, SKW, or Branded in name) is active with positive spend in the window.",
    "F010": "Add zero-conversion waste terms as negative keywords to relevant campaigns or shared exclusion lists. Prioritise by waste spend descending.",
    "F011": "Add shared negative keyword lists covering brand protection, competitor terms, and top waste queries. Assign to all active Search and Shopping campaigns.",
    "F012": "Introduce or scale Search campaigns alongside PMAX. Target at least 10–15% of total spend in Search for query-level visibility.",
    "F013": "Add or activate sitelink extensions to reach minimum 4 at account level. Ensure all are approved and point to meaningful destination pages.",
    "F014": "Add structured snippets relevant to the account's product categories (e.g., Product Types, Services, Styles).",
    "F015": "Add callout extensions highlighting USPs: free shipping, easy returns, warranty, same-day delivery.",
    "F016": "Add an approved business name asset at account level matching the actual brand/company name.",
    "F017": "Upload and approve at least one image/logo asset per active PMAX asset group. Verify PolicyApprovalStatus = APPROVED.",
    "F018": "Add asset variety to PMAX asset groups: additional headlines, descriptions, images, and video assets to improve ad strength.",
    "F019": "Enable Keyword Expander for all active Search campaigns in QT Portal > Google Channel.",
    "F020": "Add at minimum: 1 custom intent segment and 1 remarketing list to each active PMAX asset group.",
    "F021": "Add at least 10 Search Themes per PMAX asset group, aligned with top product categories and key search terms.",
    "F022": "Add EXACT match versions of the account's highest-converting search terms. Priority: branded terms, top product queries, high-CVR terms from Tab 10.",
}

# Approved naming convention values
VALID_CAMPAIGN_TYPES = {
    "pmax", "pmaxfeed", "pmaxassets", "pmaxnca", "pmaxncafeed", "pmaxncaassets",
    "pmaxlocal", "search", "searchdsa", "searchdsapagefeed", "searchdsapagefeedNCA".lower(),
    "searchnca", "shopping", "shoppinglia", "shoppingsqf",
    "demandgen", "display", "youtube",
}

VALID_STRATEGY_TAGS = {
    "general", "topproducts", "lowperformers", "zombie", "suppression", "shopifytop",
}

PROMO_KEYWORDS = re.compile(
    r'\b(sale|promo|promotion|holiday|blackfriday|black_friday|cyber|cybermonday|'
    r'bfcm|seasonal|clearance|discount|flash|laborday|memorialday|'
    r'mothersday|fathersday|valentines|halloween|thanksgiving|christmas|xmas|'
    r'primetime|primeday)\b',
    re.IGNORECASE,
)
