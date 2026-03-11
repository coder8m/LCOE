"""
Landed Cost of Power Calculator — Tkinter GUI
Based on Book1_R2.xlsx | All 24 route cases verified against Excel
Editable assumptions saved to landed_cost_assumptions.json (beside the exe)

Run:  python landed_cost_calculator.py
      OR as compiled exe — assumptions.json sits next to the exe
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json, os, sys, copy

# ═══════════════════════════════════════════════════════════════
#  PATH HELPER  (works both as .py and as PyInstaller .exe)
# ═══════════════════════════════════════════════════════════════
def _app_dir():
    """Directory of the exe / script — where JSON will be saved."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

ASSUMPTIONS_FILE = os.path.join(_app_dir(), "landed_cost_assumptions.json")

# ═══════════════════════════════════════════════════════════════
#  DEFAULT RATE VALUES  (Paise/kWh unless noted)
# ═══════════════════════════════════════════════════════════════
DEFAULT_RATES = {
    # Gujarat tariff
    "gj_app_fees":          4.0,
    "gj_ists_ltoa":         37.4,
    "gj_ists_stoa":         40.0,
    "gj_stx_ltoa":          16.32,
    "gj_stx_stoa":          37.75,
    "gj_stx_loss_pct":      3.37,    # %
    "gj_ists_loss_pct":     3.5,     # %
    "gj_add_surcharge":     82.0,
    "gj_css":               129.0,
    "gj_wheeling_chg":      20.53,
    "gj_wheeling_loss_pct": 7.25,    # %
    "gj_elec_duty":         60.0,
    "gj_rpo":               9.0,
    "gj_cd_stu":            90.0,
    "gj_lmc_default":       30.0,
    # Maharashtra tariff
    "mh_app_fees":          4.0,
    "mh_ists_ltoa":         50.0,
    "mh_ists_stoa":         53.2,
    "mh_stx_ltoa":          43.19,
    "mh_stx_stoa":          49.0,
    "mh_stx_loss_pct":      3.18,    # %
    "mh_ists_loss_pct":     3.5,     # %
    "mh_add_surcharge":     139.0,
    "mh_css":               179.0,
    "mh_wheeling_chg":      60.0,
    "mh_wheeling_loss_pct": 7.50,    # %
    "mh_elec_duty":         69.0,
    "mh_rpo":               9.0,
    "mh_cd_stu":            90.0,
    "mh_lmc_default":       30.0,
}

# Human-friendly labels for the edit dialog
RATE_LABELS = {
    "gj_app_fees":          ("Gujarat",       "Application Fees & Op. Charges",       "Ps./kWh"),
    "gj_ists_ltoa":         ("Gujarat",       "ISTS Charges – LTOA/MTOA",             "Ps./kWh"),
    "gj_ists_stoa":         ("Gujarat",       "ISTS Charges – STOA",                  "Ps./kWh"),
    "gj_stx_ltoa":          ("Gujarat",       "STU Transmission Charges – LTOA/MTOA", "Ps./kWh"),
    "gj_stx_stoa":          ("Gujarat",       "STU Transmission Charges – STOA",      "Ps./kWh"),
    "gj_stx_loss_pct":      ("Gujarat",       "STU Transmission Losses",              "%"),
    "gj_ists_loss_pct":     ("Gujarat",       "ISTS Losses",                          "%"),
    "gj_add_surcharge":     ("Gujarat",       "Additional Surcharge",                 "Ps./kWh"),
    "gj_css":               ("Gujarat",       "Cross Subsidy Surcharge",              "Ps./kWh"),
    "gj_wheeling_chg":      ("Gujarat",       "Wheeling Charges (≤33kV)",             "Ps./kWh"),
    "gj_wheeling_loss_pct": ("Gujarat",       "Wheeling Losses (≤33kV)",              "%"),
    "gj_elec_duty":         ("Gujarat",       "Electricity Duty",                     "Ps./kWh"),
    "gj_rpo":               ("Gujarat",       "RPO (REC Rate from IEX)",              "Ps./kWh"),
    "gj_cd_stu":            ("Gujarat",       "CD Charges for STU",                   "Ps./kWh"),
    "gj_lmc_default":       ("Gujarat",       "Last Mile Connectivity (default)",     "Ps./kWh"),
    "mh_app_fees":          ("Maharashtra",   "Application Fees & Op. Charges",       "Ps./kWh"),
    "mh_ists_ltoa":         ("Maharashtra",   "ISTS Charges – LTOA/MTOA",             "Ps./kWh"),
    "mh_ists_stoa":         ("Maharashtra",   "ISTS Charges – STOA",                  "Ps./kWh"),
    "mh_stx_ltoa":          ("Maharashtra",   "STU Transmission Charges – LTOA/MTOA", "Ps./kWh"),
    "mh_stx_stoa":          ("Maharashtra",   "STU Transmission Charges – STOA",      "Ps./kWh"),
    "mh_stx_loss_pct":      ("Maharashtra",   "STU Transmission Losses",              "%"),
    "mh_ists_loss_pct":     ("Maharashtra",   "ISTS Losses",                          "%"),
    "mh_add_surcharge":     ("Maharashtra",   "Additional Surcharge",                 "Ps./kWh"),
    "mh_css":               ("Maharashtra",   "Cross Subsidy Surcharge",              "Ps./kWh"),
    "mh_wheeling_chg":      ("Maharashtra",   "Wheeling Charges (≤33kV)",             "Ps./kWh"),
    "mh_wheeling_loss_pct": ("Maharashtra",   "Wheeling Losses (≤33kV)",              "%"),
    "mh_elec_duty":         ("Maharashtra",   "Electricity Duty",                     "Ps./kWh"),
    "mh_rpo":               ("Maharashtra",   "RPO (REC Rate from IEX)",              "Ps./kWh"),
    "mh_cd_stu":            ("Maharashtra",   "CD Charges for STU",                   "Ps./kWh"),
    "mh_lmc_default":       ("Maharashtra",   "Last Mile Connectivity (default)",     "Ps./kWh"),
}

# ── Assumptions text table (editable via JSON key "assumptions_text") ──────
DEFAULT_ASSUMPTIONS_TEXT = {
    "Maharashtra Grid Tariff (2024-25)": [
        ["Application Fees & Operating Charges",  "4 Ps./kWh",  ""],
        ["T-GNA Charges (ISTS Charges)",          "50 Ps./kWh (LTOA) | 53.2 Ps./kWh (STOA)", "MH: Rs.133/MW/Block"],
        ["ISTS Losses",                           "3.5% of base cost", ""],
        ["STU Transmission Charges (LTOA/MTOA)",  "43.19 Ps./kWh",  "310.98 Rs./kW/Month"],
        ["STU Transmission Charges (STOA)",       "49 Ps./kWh",  ""],
        ["STU Transmission Losses",               "3.18% of base cost", "Pg. 4 MH tariff order"],
        ["Additional Surcharge",                  "139 Ps./kWh", "Not applicable for RE / Captive"],
        ["Cross Subsidy Surcharge",               "179 Ps./kWh", "Not applicable for RE / Captive"],
        ["Wheeling Charges",                      "0 (>33kV) | 60 Ps./kWh (≤33kV)", ""],
        ["Wheeling Losses",                       "7.5% of base cost (≤33kV)", ""],
        ["Electricity Duty",                      "69 Ps./kWh", "ACS=8.91; FAC=0.3; ED=0.075"],
        ["RPO (REC Rate from IEX)",               "9 Ps./kWh", ""],
        ["Last Mile Connectivity – CTU",          "30 Ps./kWh (assumption)", ""],
        ["CD Charges for STU",                    "90 Ps./kWh", ""],
    ],
    "Gujarat Grid Tariff (2024-25)": [
        ["Application Fees & Operating Charges",  "4 Ps./kWh",  ""],
        ["T-GNA Charges (ISTS Charges)",          "37.4 Ps./kWh (LTOA) | 40.0 Ps./kWh (STOA)", "GJ: Rs.100/MW/Block"],
        ["ISTS Losses",                           "3.5% of base cost", ""],
        ["STU Transmission Charges (LTOA/MTOA)",  "16.32 Ps./kWh", ""],
        ["STU Transmission Charges (STOA)",       "37.75 Ps./kWh", "Pg.no.130"],
        ["STU Transmission Losses",               "3.37% of base cost", "Pg.no.75"],
        ["Additional Surcharge",                  "82 Ps./kWh", "Not applicable for RE / Captive; Pg.no.16"],
        ["Cross Subsidy Surcharge",               "129 Ps./kWh", "Not applicable for RE / Captive; Pg.no.15"],
        ["Wheeling Charges",                      "0 (>33kV) | 20.53 Ps./kWh (≤33kV)", "Pg.no.349"],
        ["Wheeling Losses",                       "7.25% of base cost (≤33kV)", "Pg.no.350"],
        ["Electricity Duty",                      "60 Ps./kWh | 15% for RE", ""],
        ["RPO (REC Rate from IEX)",               "9 Ps./kWh", ""],
        ["Last Mile Connectivity – CTU",          "30 Ps./kWh (assumption)", ""],
        ["CD Charges for STU",                    "90 Ps./kWh", ""],
    ],
    "CTU / ISTS Charges (2024-25)": [
        ["GNA Capacity",                          "500 MW", ""],
        ["T-GNA Open Access Charges",             "0.525 Rs./kWh", "GJ=Rs.100.05; MH=Rs.133/MW/Block"],
        ["AC-UBC (Usage Based)",                  "0.0063 Rs./kWh", "Rs.21.2 Lac"],
        ["AC-BC (Balance)",                       "0.1862 Rs./kWh", "Rs.625.5 Lac"],
        ["NC-RE (National RE Component)",         "0.0395 Rs./kWh", "Rs.132.6 Lac"],
        ["NC-HVDC",                               "0.0327 Rs./kWh", "Rs.109.9 Lac"],
        ["RC (Regional Component)",               "0.0152 Rs./kWh", "Rs.50.9 Lac"],
        ["Total ISTS Transmission Charges",       "0.2798 Rs./kWh", ""],
        ["GNA Calculation – Gujarat",             "37.4 Ps./kWh", "Monthly bill Rs.3400.8 Cr | 12,623 MW"],
        ["GNA Calculation – Maharashtra",         "50.0 Ps./kWh", "Monthly bill Rs.3400.3 Cr | 9,410 MW"],
        ["GNA Calculation – Jamnagar",            "29.8 Ps./kWh", "Rs.107.4 Cr | 500 MW"],
        ["ISTS Losses",                           "3.5% of base cost", "Applied to all CTU routes"],
    ],
    "Open Access Charge Applicability": [
        ["Additional Surcharge",                  "Not applicable for Captive / RE", "Power exchange only"],
        ["Cross Subsidy Surcharge",               "Not applicable for Captive / RE", "Power exchange only"],
        ["Wheeling Charges & Losses",             "Not applicable for voltage > 33kV", "0.6 Rs./kWh for <33kV"],
        ["CD Charges",                            "Not applicable for Captive", "Check for export transactions"],
        ["STU Charges",                           "Not applicable for CTU-to-CTU routes", ""],
        ["ISTS Charges",                          "Not applicable for Intra-state STU-to-STU", ""],
        ["Last Mile Connectivity (LMC)",          "1× single CTU; 2× for CTU-to-CTU", "Default 30 Ps./kWh per leg"],
        ["MH STU Wheeling",                       "Under discussion", "To be confirmed"],
        ["RE Transmission (MTOA/LTOA)",           "2× STOA charges for RE under MTOA/LTOA", "Pg.no.18 MH order"],
        ["CD Charges for RE",                     "Not applicable for RE-based OA", "Pg.no.17 MH order"],
    ],
    "Calculator Business Rules": [
        ["Power Type",            "Thermal / Renewable", "RE zeroes Additional Surcharge & CSS"],
        ["Mode of OA",            "Captive / Third Party", "Captive zeroes Add. Surcharge, CSS, CD"],
        ["Voltage Level",         ">33kV / ≤33kV", ">33kV zeroes Wheeling Charges & Losses"],
        ["OA Duration",           "LTOA / MTOA / STOA", "STOA uses higher ISTS & STU rates"],
        ["Injection/Wth State",   "Gujarat / Maharashtra / IEX", "Determines GJ or MH tariff rates"],
        ["Connectivity",          "STU / CTU", "CTU attracts ISTS; STU attracts STU tariff"],
        ["Last Mile Connectivity","User-defined (default 30 Ps./kWh)", "1× or 2× depending on CTU legs"],
        ["Base Cost",             "User-defined (default 400 Ps./kWh)", "Losses are % of base cost"],
        ["Formula",               "Landed Cost = Σ(mults × tariff rates) + Base Cost", "Rs./kWh"],
        ["Pending Changes",       "Diff. STU charges for Thermal vs RE", "Changes sheet"],
    ],
    "Cost Applicability (MEL vs IEX)": [
        ["ISTS Charges",              "MEL: Yes | IEX: Yes",  ""],
        ["ISTS Losses",               "MEL: Yes | IEX: Yes",  ""],
        ["STU Transmission Charges",  "MEL: 0.49 Rs./kWh | IEX: 0.49 Rs./kWh", ""],
        ["STU Losses",                "MEL: 3.18% | IEX: 3.18%", ""],
        ["Electrical Duty",           "MEL: 0.69 Rs./kWh | IEX: 0.69 Rs./kWh", ""],
        ["Wheeling Charges",          "MEL: No | IEX: No", "Not applicable >33kV"],
        ["Wheeling Losses",           "MEL: No | IEX: No", "Not applicable >33kV"],
        ["Cross Subsidy Surcharge",   "MEL: No | IEX: Yes", "0 for MEL captive"],
        ["Additional Surcharge",      "MEL: No | IEX: Yes", "0 for MEL captive"],
        ["Other Charges",             "MEL: Yes | IEX: Yes", "RPO, LMC, CD as applicable"],
    ],
}


# ═══════════════════════════════════════════════════════════════
#  PERSISTENCE  — load / save JSON
# ═══════════════════════════════════════════════════════════════
def load_settings():
    """Load rates + assumptions text from JSON; fall back to defaults."""
    if os.path.exists(ASSUMPTIONS_FILE):
        try:
            with open(ASSUMPTIONS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            rates = {k: data.get("rates", {}).get(k, DEFAULT_RATES[k])
                     for k in DEFAULT_RATES}
            text  = data.get("assumptions_text", DEFAULT_ASSUMPTIONS_TEXT)
            return rates, text
        except Exception:
            pass
    return copy.deepcopy(DEFAULT_RATES), copy.deepcopy(DEFAULT_ASSUMPTIONS_TEXT)


def save_settings(rates, assumptions_text):
    """Persist rates + assumptions text to JSON."""
    with open(ASSUMPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump({"rates": rates, "assumptions_text": assumptions_text},
                  f, indent=2, ensure_ascii=False)


# ═══════════════════════════════════════════════════════════════
#  MULTIPLIER MATRIX  (fixed — from Main Sheet rows 6-18)
# ═══════════════════════════════════════════════════════════════
CASE_MATRIX = {
    ("GUJARAT",     "STU", "GUJARAT",     "STU"):  ([1,0,0,1,1,0,0,0,0,1,1,0,1],"GJ",None,None),
    ("MAHARASHTRA", "STU", "MAHARASHTRA", "STU"):  ([1,0,0,1,1,0,0,0,0,1,1,0,1],"MH",None,None),
    ("GUJARAT",     "STU", "GUJARAT",     "CTU"):  ([1,1,1,1,1,0,0,0,0,1,1,1,1],"GJ",None,None),
    ("MAHARASHTRA", "STU", "MAHARASHTRA", "CTU"):  ([1,1,1,1,1,0,0,0,0,1,1,1,0],"MH",None,None),
    ("GUJARAT",     "CTU", "GUJARAT",     "CTU"):  ([1,1,1,0,0,0,0,0,0,1,1,2,1],"GJ",None,None),
    ("MAHARASHTRA", "CTU", "MAHARASHTRA", "CTU"):  ([1,1,1,0,0,0,0,0,0,1,1,2,1],"MH",None,None),
    ("GUJARAT",     "CTU", "GUJARAT",     "STU"):  ([1,1,1,1,1,0,0,0,0,1,1,1,1],"MH",None,None),
    ("MAHARASHTRA", "CTU", "MAHARASHTRA", "STU"):  ([1,1,1,1,1,0,0,0,0,1,1,1,1],"MH",None,None),
    ("GUJARAT",     "STU", "MAHARASHTRA", "STU"):  ([1,0,0,1,1,0,0,0,0,0,0,0,0],"GJ",[1,0,0,1,1,0,0,0,0,1,1,0,1],"MH"),
    ("MAHARASHTRA", "CTU", "GUJARAT",     "STU"):  ([1,0,0,1,1,0,0,0,0,1,1,0,1],"GJ",[1,1,1,0,0,0,0,0,0,0,0,1,0],"MH"),
    ("GUJARAT",     "CTU", "MAHARASHTRA", "STU"):  ([1,1,1,0,0,0,0,0,0,0,0,1,0],"GJ",[1,0,0,1,1,0,0,0,0,1,1,0,1],"MH"),
    ("MAHARASHTRA", "STU", "GUJARAT",     "CTU"):  ([1,1,1,0,0,0,0,0,0,1,1,1,0],"GJ",[1,0,0,1,1,0,0,0,0,0,0,0,0],"MH"),
    ("GUJARAT",     "CTU", "MAHARASHTRA", "CTU"):  ([1,1,1,0,0,0,0,0,0,0,0,0,0],"GJ",[1,1,1,0,0,0,0,0,0,1,1,1,0],"MH"),
    ("MAHARASHTRA", "STU", "GUJARAT",     "STU"):  ([1,0,0,1,1,0,0,0,0,1,1,0,1],"GJ",[1,0,0,1,1,0,0,0,0,0,0,0,0],"MH"),
    ("GUJARAT",     "STU", "MAHARASHTRA", "CTU"):  ([1,0,0,1,1,0,0,0,0,0,0,0,0],"GJ",[1,1,1,0,0,0,0,0,0,1,1,1,0],"MH"),
    ("MAHARASHTRA", "CTU", "GUJARAT",     "CTU"):  ([1,1,1,0,0,0,0,0,0,1,1,1,0],"GJ",[1,1,1,0,0,0,0,0,0,0,0,1,0],"MH"),
    ("IEX","CTU","GUJARAT",     "STU"):            ([1,1,1,1,1,0,0,0,0,1,1,1,1],"GJ",None,None),
    ("IEX","CTU","GUJARAT",     "CTU"):            ([1,1,1,0,0,0,0,0,0,1,1,1,1],"GJ",None,None),
    ("IEX","CTU","MAHARASHTRA", "CTU"):            ([1,1,1,0,0,0,0,0,0,1,1,1,1],"MH",None,None),
    ("IEX","CTU","MAHARASHTRA", "STU"):            ([1,1,1,1,1,0,0,0,0,1,1,1,1],"MH",None,None),
    ("IEX","STU","GUJARAT",     "CTU"):            ([1,1,1,0,0,0,0,0,0,1,1,1,1],"GJ",None,None),
    ("IEX","STU","GUJARAT",     "STU"):            ([1,0,0,1,1,0,0,0,0,1,1,1,1],"GJ",None,None),
    ("IEX","STU","MAHARASHTRA", "STU"):            ([1,0,0,1,1,0,0,0,0,1,1,1,1],"MH",None,None),
    ("IEX","STU","MAHARASHTRA", "CTU"):            ([1,1,1,0,0,0,0,0,0,1,1,1,1],"MH",None,None),
}

CHARGE_NAMES = [
    "Application Fees & Operating Charges",
    "ISTS Charges (GNA / T-GNA)",
    "ISTS Losses",
    "Transmission Charges – STU",
    "Transmission Losses – STU",
    "Additional Surcharge",
    "Cross Subsidy Surcharge",
    "Wheeling Charges",
    "Wheeling Losses",
    "Electricity Duty",
    "RPO (REC Rate from IEX)",
    "Last Mile Connectivity (CTU)",
    "CD Charges for STU",
]


def _rate_vector(side, base_cost, lmc, rates, stoa):
    """Build 13-element vector (Ps./kWh) for 'GJ' or 'MH'."""
    p = "gj_" if side == "GJ" else "mh_"
    return [
        rates[p+"app_fees"],
        rates[p+"ists_stoa"]  if stoa else rates[p+"ists_ltoa"],
        rates[p+"ists_loss_pct"] / 100 * base_cost,
        rates[p+"stx_stoa"]   if stoa else rates[p+"stx_ltoa"],
        rates[p+"stx_loss_pct"]  / 100 * base_cost,
        rates[p+"add_surcharge"],
        rates[p+"css"],
        rates[p+"wheeling_chg"],
        rates[p+"wheeling_loss_pct"] / 100 * base_cost,
        rates[p+"elec_duty"],
        rates[p+"rpo"],
        lmc,
        rates[p+"cd_stu"],
    ]


def compute_landed_cost(power_type, base_cost, inj_state, inj_conn,
                        wth_state, wth_conn, voltage, mode_oa, lmc,
                        oa_duration, rates):
    key   = (inj_state, inj_conn, wth_state, wth_conn)
    entry = CASE_MATRIX.get(key)
    if entry is None:
        return None, None, f"Case not found: {inj_state} [{inj_conn}] → {wth_state} [{wth_conn}]"

    ab_raw, ab_type, ac_raw, ac_type = entry
    stoa         = (oa_duration == "STOA")
    is_captive   = (mode_oa    == "CAPTIVE")
    is_renewable = (power_type == "Renewable")
    is_above_33  = (voltage    == ">33kV")

    def rule(m):
        if m is None: return None
        m = list(m)
        if is_captive or is_renewable: m[5] = m[6] = 0
        if is_above_33:                m[7] = m[8] = 0
        return m

    ab = rule(ab_raw);  ac = rule(ac_raw)
    ab_v = _rate_vector(ab_type, base_cost, lmc, rates, stoa) if ab_type else None
    ac_v = _rate_vector(ac_type, base_cost, lmc, rates, stoa) if ac_type else None

    total = base_cost / 100
    breakdown = {}
    for i, name in enumerate(CHARGE_NAMES):
        val = ((ab[i] * ab_v[i] / 100) if (ab and ab_v) else 0) + \
              ((ac[i] * ac_v[i] / 100) if (ac and ac_v) else 0)
        if abs(val) > 1e-9:
            breakdown[name] = val
        total += val

    return round(total, 4), breakdown, None


# ═══════════════════════════════════════════════════════════════
#  COLOUR PALETTE
# ═══════════════════════════════════════════════════════════════
C = {
    "bg":      "#0D1B2A", "panel":   "#142233", "card":    "#1E3148",
    "border":  "#2A4A6A", "accent":  "#00C6AE", "accent2": "#F4A261",
    "success": "#00E5B0", "text":    "#E8F0F7", "subtext": "#7EA8C4",
    "red":     "#E76F51", "entry":   "#243548", "hdr_bg":  "#0A3A5A",
    "warn":    "#FFD166",
}


# ═══════════════════════════════════════════════════════════════
#  EDIT ASSUMPTIONS DIALOG
# ═══════════════════════════════════════════════════════════════
class EditAssumptionsDialog(tk.Toplevel):
    """
    Two-tab dialog:
      Tab A — Edit numeric rates (GJ / MH tariff values)
      Tab B — Edit assumption text table rows (per section)
    Changes are saved to JSON and applied immediately.
    """
    def __init__(self, parent, rates, assumptions_text, on_save):
        super().__init__(parent)
        self.title("Edit Assumptions & Rates")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.geometry("920x680")
        self.minsize(760, 520)
        self.grab_set()

        self._rates_work  = copy.deepcopy(rates)
        self._text_work   = copy.deepcopy(assumptions_text)
        self._on_save     = on_save
        self._entries     = {}   # key → StringVar for numeric fields
        self._text_vars   = {}   # (section, row, col) → StringVar

        self._setup_styles()
        self._build_ui()
        self.transient(parent)

    def _setup_styles(self):
        s = ttk.Style(self)
        s.configure("TNotebook",      background=C["bg"],    borderwidth=0)
        s.configure("TNotebook.Tab",  background=C["panel"], foreground=C["subtext"],
                    font=("Segoe UI", 10, "bold"), padding=(14, 7))
        s.map("TNotebook.Tab",
              background=[("selected", C["card"])],
              foreground=[("selected", C["accent"])])

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=C["panel"], pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✏  EDIT ASSUMPTIONS & RATES",
                 bg=C["panel"], fg=C["accent"],
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=18)
        tk.Label(hdr,
                 text="Changes are saved to  landed_cost_assumptions.json  and applied instantly",
                 bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 8, "italic")).pack(side="left", padx=6)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(8, 0))

        tab_rates = tk.Frame(nb, bg=C["bg"])
        nb.add(tab_rates, text="  📊  Tariff Rates  ")
        self._build_rates_tab(tab_rates)

        tab_text = tk.Frame(nb, bg=C["bg"])
        nb.add(tab_text, text="  📋  Assumption Text  ")
        self._build_text_tab(tab_text)

        # Buttons
        btn_bar = tk.Frame(self, bg=C["panel"], pady=10)
        btn_bar.pack(fill="x", side="bottom")
        tk.Button(btn_bar, text="💾  Save & Apply",
                  bg=C["accent"], fg="#0D1B2A",
                  font=("Segoe UI", 11, "bold"), relief="flat",
                  padx=18, pady=6, cursor="hand2",
                  command=self._save).pack(side="right", padx=14)
        tk.Button(btn_bar, text="↺  Restore Defaults",
                  bg=C["border"], fg=C["text"],
                  font=("Segoe UI", 10), relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self._restore_defaults).pack(side="right", padx=4)
        tk.Button(btn_bar, text="✕  Cancel",
                  bg=C["panel"], fg=C["subtext"],
                  font=("Segoe UI", 10), relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=self.destroy).pack(side="left", padx=14)

    # ── Tab A: Numeric rates ─────────────────────────────────
    def _build_rates_tab(self, parent):
        tk.Label(parent,
                 text="  Edit tariff rates below. Changes affect all calculations immediately after saving.",
                 bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 8, "italic")).pack(anchor="w", padx=10, pady=(6, 2))

        wrap = tk.Frame(parent, bg=C["bg"])
        wrap.pack(fill="both", expand=True, padx=8, pady=4)

        cv = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg=C["bg"])
        inner.bind("<Configure>",
            lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        cv.bind_all("<MouseWheel>",
            lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Column headers
        hdr = tk.Frame(inner, bg=C["border"], pady=5)
        hdr.pack(fill="x", padx=4)
        for txt, w in [("  State", 14), ("  Charge / Parameter", 36),
                       ("  Value", 14), ("  Unit", 10)]:
            tk.Label(hdr, text=txt, bg=C["border"], fg=C["accent"],
                     font=("Segoe UI", 9, "bold"), width=w, anchor="w",
                     padx=4).pack(side="left")

        row_bgs = [C["card"], "#192D42"]
        current_group = None
        idx = 0
        for key, (group, label, unit) in RATE_LABELS.items():
            if group != current_group:
                current_group = group
                g_hdr = tk.Frame(inner, bg=C["hdr_bg"], pady=4)
                g_hdr.pack(fill="x", padx=4, pady=(8, 1))
                tk.Label(g_hdr, text=f"  ◈  {group}",
                         bg=C["hdr_bg"], fg=C["accent2"],
                         font=("Segoe UI", 10, "bold")).pack(anchor="w")

            bg = row_bgs[idx % 2]; idx += 1
            row = tk.Frame(inner, bg=bg, pady=3)
            row.pack(fill="x", padx=4)

            tk.Label(row, text=f"  {group}", bg=bg, fg=C["subtext"],
                     font=("Segoe UI", 9), width=14, anchor="w").pack(side="left")
            tk.Label(row, text=label, bg=bg, fg=C["text"],
                     font=("Segoe UI", 9), width=36, anchor="w",
                     padx=4).pack(side="left")

            var = tk.StringVar(value=str(self._rates_work[key]))
            self._entries[key] = var
            tk.Entry(row, textvariable=var, width=14,
                     bg=C["entry"], fg=C["text"],
                     insertbackground=C["text"], relief="flat",
                     font=("Segoe UI", 10),
                     highlightthickness=1,
                     highlightbackground=C["border"],
                     highlightcolor=C["accent"]).pack(side="left", padx=4)
            tk.Label(row, text=unit, bg=bg, fg=C["subtext"],
                     font=("Segoe UI", 9), width=10, anchor="w").pack(side="left")

    # ── Tab B: Assumption text rows ──────────────────────────
    def _build_text_tab(self, parent):
        top = tk.Frame(parent, bg=C["panel"], padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="Section:", bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 10)).pack(side="left")
        self._v_section = tk.StringVar(value=list(self._text_work.keys())[0])
        cb = ttk.Combobox(top, textvariable=self._v_section,
                          values=list(self._text_work.keys()),
                          state="readonly", width=44,
                          font=("Segoe UI", 10))
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self._render_text_rows())

        tk.Label(top,
                 text="Edit cells directly. Each row: Particulars | Rate/Value | Remarks",
                 bg=C["panel"], fg=C["border"],
                 font=("Segoe UI", 8, "italic")).pack(side="left", padx=10)

        btn_frame = tk.Frame(top, bg=C["panel"])
        btn_frame.pack(side="right")
        tk.Button(btn_frame, text="＋ Add Row",
                  bg=C["border"], fg=C["text"],
                  font=("Segoe UI", 9), relief="flat", padx=8, pady=3,
                  cursor="hand2",
                  command=self._add_row).pack(side="left", padx=2)
        tk.Button(btn_frame, text="－ Delete Last",
                  bg=C["border"], fg=C["text"],
                  font=("Segoe UI", 9), relief="flat", padx=8, pady=3,
                  cursor="hand2",
                  command=self._del_row).pack(side="left", padx=2)

        self._text_content = tk.Frame(parent, bg=C["bg"])
        self._text_content.pack(fill="both", expand=True, padx=8, pady=4)

        self._txt_cv = tk.Canvas(self._text_content, bg=C["bg"], highlightthickness=0)
        tsb = ttk.Scrollbar(self._text_content, orient="vertical",
                            command=self._txt_cv.yview)
        self._txt_inner = tk.Frame(self._txt_cv, bg=C["bg"])
        self._txt_inner.bind("<Configure>",
            lambda e: self._txt_cv.configure(
                scrollregion=self._txt_cv.bbox("all")))
        self._txt_cv.create_window((0, 0), window=self._txt_inner, anchor="nw")
        self._txt_cv.configure(yscrollcommand=tsb.set)
        self._txt_cv.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")

        self._render_text_rows()

    def _render_text_rows(self):
        for w in self._txt_inner.winfo_children():
            w.destroy()
        self._text_vars = {}

        section = self._v_section.get()
        rows    = self._text_work.get(section, [])

        # Header
        hdr = tk.Frame(self._txt_inner, bg=C["border"], pady=5)
        hdr.pack(fill="x")
        for txt, w in [("  Particulars", 38), ("  Rate / Value", 30), ("  Remarks", 30)]:
            tk.Label(hdr, text=txt, bg=C["border"], fg=C["accent"],
                     font=("Segoe UI", 9, "bold"), width=w, anchor="w",
                     padx=4).pack(side="left")

        row_bgs = [C["card"], "#192D42"]
        for r_idx, row_data in enumerate(rows):
            bg = row_bgs[r_idx % 2]
            row_f = tk.Frame(self._txt_inner, bg=bg, pady=2)
            row_f.pack(fill="x")
            for c_idx, (w, ph) in enumerate(
                    [(38, "Particular…"), (30, "Rate / Value…"), (30, "Remark…")]):
                val = row_data[c_idx] if c_idx < len(row_data) else ""
                var = tk.StringVar(value=val)
                self._text_vars[(section, r_idx, c_idx)] = var
                tk.Entry(row_f, textvariable=var, width=w,
                         bg=C["entry"], fg=C["text"],
                         insertbackground=C["text"], relief="flat",
                         font=("Segoe UI", 9),
                         highlightthickness=1,
                         highlightbackground=bg,
                         highlightcolor=C["accent"]).pack(side="left", padx=2, pady=1)

    def _collect_text_section(self):
        """Flush current section entry vars back into _text_work."""
        section = self._v_section.get()
        n_rows  = len(self._text_work.get(section, []))
        for r in range(n_rows):
            for c in range(3):
                key = (section, r, c)
                if key in self._text_vars:
                    self._text_work[section][r][c] = self._text_vars[key].get()

    def _add_row(self):
        self._collect_text_section()
        section = self._v_section.get()
        self._text_work[section].append(["New item", "Value", "Remark"])
        self._render_text_rows()

    def _del_row(self):
        self._collect_text_section()
        section = self._v_section.get()
        if self._text_work[section]:
            self._text_work[section].pop()
            self._render_text_rows()

    def _save(self):
        # Validate & collect numeric entries
        new_rates = {}
        errors = []
        for key, var in self._entries.items():
            try:
                new_rates[key] = float(var.get())
            except ValueError:
                _, label, _ = RATE_LABELS[key]
                errors.append(label)
        if errors:
            messagebox.showerror("Invalid Values",
                "Please enter valid numbers for:\n" + "\n".join(f"  • {e}" for e in errors),
                parent=self)
            return

        # Collect text edits from current section
        self._collect_text_section()

        # Persist & notify parent
        save_settings(new_rates, self._text_work)
        self._on_save(new_rates, self._text_work)
        messagebox.showinfo("Saved",
            f"Assumptions saved to:\n{ASSUMPTIONS_FILE}\n\nCalculator updated instantly.",
            parent=self)
        self.destroy()

    def _restore_defaults(self):
        if not messagebox.askyesno("Restore Defaults",
                "Reset ALL rates and assumption text to original values?",
                parent=self):
            return
        self._rates_work = copy.deepcopy(DEFAULT_RATES)
        self._text_work  = copy.deepcopy(DEFAULT_ASSUMPTIONS_TEXT)
        # Refresh entries
        for key, var in self._entries.items():
            var.set(str(self._rates_work[key]))
        self._render_text_rows()


# ═══════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════
class LandedCostApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Landed Cost of Power Calculator")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.geometry("1240x880")
        self.minsize(980, 680)

        # Load persisted settings
        self._rates, self._assumptions_text = load_settings()

        self._setup_styles()
        self._build_ui()
        self.after(120, self.calculate)

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",        background=C["bg"])
        s.configure("TLabel",        background=C["bg"],    foreground=C["text"],    font=("Segoe UI", 10))
        s.configure("Sub.TLabel",    background=C["bg"],    foreground=C["subtext"], font=("Segoe UI", 9))
        s.configure("Panel.TLabel",  background=C["panel"], foreground=C["text"],    font=("Segoe UI", 10))
        s.configure("TCombobox",
                    fieldbackground=C["entry"], background=C["entry"],
                    foreground=C["text"], selectbackground=C["accent"],
                    selectforeground="#000", font=("Segoe UI", 10))
        s.map("TCombobox",
              fieldbackground=[("readonly", C["entry"])],
              foreground=[("readonly", C["text"])],
              selectbackground=[("readonly", C["accent"])])
        s.configure("Calc.TButton",  background=C["accent"], foreground="#0D1B2A",
                    font=("Segoe UI", 12, "bold"), padding=(12, 9), relief="flat")
        s.map("Calc.TButton",  background=[("active", C["success"])])
        s.configure("Edit.TButton",  background=C["accent2"], foreground="#0D1B2A",
                    font=("Segoe UI", 10, "bold"), padding=(10, 7), relief="flat")
        s.map("Edit.TButton",  background=[("active", C["warn"])])
        s.configure("Reset.TButton", background=C["border"],  foreground=C["text"],
                    font=("Segoe UI", 10), padding=(8, 6), relief="flat")
        s.map("Reset.TButton", background=[("active", C["card"])])
        s.configure("TNotebook",     background=C["bg"],    borderwidth=0)
        s.configure("TNotebook.Tab", background=C["panel"], foreground=C["subtext"],
                    font=("Segoe UI", 10, "bold"), padding=(16, 8))
        s.map("TNotebook.Tab",
              background=[("selected", C["card"])],
              foreground=[("selected", C["accent"])])

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=C["panel"], pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚡  LANDED COST OF POWER CALCULATOR",
                 bg=C["panel"], fg=C["accent"],
                 font=("Segoe UI", 18, "bold")).pack(side="left", padx=22)
        tk.Label(hdr, text="Open Access  ·  Gujarat & Maharashtra  ·  FY 2024-25",
                 bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 9)).pack(side="left", padx=4)

        # ── Edit Assumptions button (top-right) ──
        ttk.Button(hdr, text="✏  Edit Assumptions",
                   style="Edit.TButton",
                   command=self._open_edit_dialog).pack(side="right", padx=18)

        # Notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(8, 10))

        calc_tab  = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(calc_tab, text="  ⚡  Calculator  ")
        self._build_calculator_tab(calc_tab)

        self._assum_tab = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(self._assum_tab, text="  📋  Assumptions & Rates  ")
        self._build_assumptions_tab(self._assum_tab)

    # ── Edit dialog callback ─────────────────────────────────
    def _open_edit_dialog(self):
        EditAssumptionsDialog(self, self._rates, self._assumptions_text,
                              on_save=self._on_assumptions_saved)

    def _on_assumptions_saved(self, new_rates, new_text):
        self._rates            = new_rates
        self._assumptions_text = new_text
        self.calculate()                    # recalculate with new rates
        self._refresh_assumptions_tab()    # refresh text tab

    # ─────────────────────────────────────────────────────────
    #  CALCULATOR TAB
    # ─────────────────────────────────────────────────────────
    def _build_calculator_tab(self, parent):
        body = tk.Frame(parent, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=14, pady=12)
        body.columnconfigure(0, weight=2, minsize=355)
        body.columnconfigure(1, weight=3)
        body.rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=C["bg"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        self._build_inputs(left)

        right = tk.Frame(body, bg=C["bg"])
        right.grid(row=0, column=1, sticky="nsew")
        self._build_results(right)

    def _sec(self, p, title):
        row = tk.Frame(p, bg=C["bg"])
        row.pack(fill="x", pady=(13, 2))
        tk.Label(row, text=title, bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Frame(row, bg=C["border"], height=1).pack(
            side="left", fill="x", expand=True, padx=(8, 0))

    def _irow(self, p, label, make):
        row = tk.Frame(p, bg=C["bg"])
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 9), width=34, anchor="w").pack(side="left")
        make(row).pack(side="left", fill="x", expand=True)

    def _combo(self, p, vals, var):
        cb = ttk.Combobox(p, textvariable=var, values=vals,
                          state="readonly", width=20)
        cb.bind("<<ComboboxSelected>>", lambda e: self.calculate())
        return cb

    def _ent(self, p, var):
        e = tk.Entry(p, textvariable=var, width=22,
                     bg=C["entry"], fg=C["text"],
                     insertbackground=C["text"], relief="flat",
                     font=("Segoe UI", 10),
                     highlightthickness=1,
                     highlightbackground=C["border"],
                     highlightcolor=C["accent"])
        e.bind("<Return>",   lambda _: self.calculate())
        e.bind("<FocusOut>", lambda _: self.calculate())
        return e

    def _build_inputs(self, parent):
        self.v_ptype    = tk.StringVar(value="Thermal")
        self.v_bcost    = tk.StringVar(value="400")
        self.v_inj_st   = tk.StringVar(value="MAHARASHTRA")
        self.v_inj_conn = tk.StringVar(value="CTU")
        self.v_voltage  = tk.StringVar(value=">33kV")
        self.v_wth_st   = tk.StringVar(value="GUJARAT")
        self.v_wth_conn = tk.StringVar(value="CTU")
        self.v_mode     = tk.StringVar(value="CAPTIVE")
        self.v_lmc      = tk.StringVar(value=str(self._rates.get("gj_lmc_default", 30)))
        self.v_oadr     = tk.StringVar(value="LTOA")

        self._sec(parent, "◈  POWER SOURCE")
        self._irow(parent, "Power Type",
            lambda p: self._combo(p, ["Thermal","Renewable"], self.v_ptype))
        self._irow(parent, "Base Cost  (Ps./kWh)",
            lambda p: self._ent(p, self.v_bcost))

        self._sec(parent, "◈  INJECTION POINT  (FROM)")
        self._irow(parent, "Injection State",
            lambda p: self._combo(p, ["GUJARAT","MAHARASHTRA","IEX"], self.v_inj_st))
        self._irow(parent, "Injection Connectivity",
            lambda p: self._combo(p, ["STU","CTU"], self.v_inj_conn))
        self._irow(parent, "Voltage Level",
            lambda p: self._combo(p, [">33kV","≤33kV"], self.v_voltage))

        self._sec(parent, "◈  WITHDRAWAL POINT  (TO)")
        self._irow(parent, "Withdrawal State",
            lambda p: self._combo(p, ["GUJARAT","MAHARASHTRA"], self.v_wth_st))
        self._irow(parent, "Withdrawal Connectivity",
            lambda p: self._combo(p, ["STU","CTU"], self.v_wth_conn))

        self._sec(parent, "◈  OPEN ACCESS PARAMETERS")
        self._irow(parent, "Mode of OA",
            lambda p: self._combo(p, ["CAPTIVE","THIRD PARTY"], self.v_mode))
        self._irow(parent, "Last Mile Connectivity  (Ps./kWh)",
            lambda p: self._ent(p, self.v_lmc))
        self._irow(parent, "OA Duration",
            lambda p: self._combo(p, ["LTOA","MTOA","STOA"], self.v_oadr))

        bf = tk.Frame(parent, bg=C["bg"])
        bf.pack(fill="x", pady=(18, 0))
        ttk.Button(bf, text="⚡   CALCULATE LANDED COST",
                   style="Calc.TButton", command=self.calculate
                   ).pack(fill="x", ipady=4)
        ttk.Button(bf, text="✏  Edit Assumptions & Rates",
                   style="Edit.TButton", command=self._open_edit_dialog
                   ).pack(fill="x", pady=(7, 0), ipady=2)
        ttk.Button(bf, text="↺  Reset to Defaults",
                   style="Reset.TButton", command=self.reset
                   ).pack(fill="x", pady=(7, 0))

        hint = tk.Frame(parent, bg=C["card"], padx=10, pady=8)
        hint.pack(fill="x", pady=(14, 0))
        tk.Label(hint,
                 text=("ℹ  Wheeling charges auto-zero for voltage > 33kV\n"
                       "   Add. Surcharge & CSS zero for Captive / Renewable\n"
                       "   All 24 route cases verified against Excel sheet\n"
                       "   Edited assumptions save to JSON and apply instantly"),
                 bg=C["card"], fg=C["subtext"],
                 font=("Segoe UI", 8), justify="left").pack(anchor="w")

    def _build_results(self, parent):
        card = tk.Frame(parent, bg=C["panel"],
                        highlightthickness=1, highlightbackground=C["accent"],
                        padx=22, pady=16)
        card.pack(fill="x", pady=(0, 12))
        tk.Label(card, text="LANDED COST OF POWER",
                 bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 9, "bold")).pack()
        self.lbl_val = tk.Label(card, text="—",
                                bg=C["panel"], fg=C["success"],
                                font=("Segoe UI", 44, "bold"))
        self.lbl_val.pack()
        tk.Label(card, text="Rs. / kWh",
                 bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 11)).pack()
        self.lbl_case = tk.Label(card, text="Calculating…",
                                 bg=C["panel"], fg=C["accent2"],
                                 font=("Segoe UI", 9, "italic"),
                                 wraplength=620, justify="center")
        self.lbl_case.pack(pady=(7, 0))

        sec_row = tk.Frame(parent, bg=C["bg"])
        sec_row.pack(fill="x", pady=(2, 4))
        tk.Label(sec_row, text="CHARGE BREAKDOWN",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Frame(sec_row, bg=C["border"], height=1).pack(
            side="left", fill="x", expand=True, padx=(8, 0))

        wrap = tk.Frame(parent, bg=C["bg"])
        wrap.pack(fill="both", expand=True)
        self._cv = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self._cv.yview)
        self.tbl = tk.Frame(self._cv, bg=C["bg"])
        self.tbl.bind("<Configure>",
            lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.create_window((0, 0), window=self.tbl, anchor="nw")
        self._cv.configure(yscrollcommand=sb.set)
        self._cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._cv.bind_all("<MouseWheel>",
            lambda e: self._cv.yview_scroll(int(-1*(e.delta/120)), "units"))

    # ─────────────────────────────────────────────────────────
    #  ASSUMPTIONS TAB
    # ─────────────────────────────────────────────────────────
    def _build_assumptions_tab(self, parent):
        top = tk.Frame(parent, bg=C["panel"], padx=14, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="Section:", bg=C["panel"], fg=C["subtext"],
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))
        self.v_asec = tk.StringVar(value=list(self._assumptions_text.keys())[0])
        self._asec_cb = ttk.Combobox(top, textvariable=self.v_asec,
                                     values=list(self._assumptions_text.keys()),
                                     state="readonly", width=45,
                                     font=("Segoe UI", 10))
        self._asec_cb.pack(side="left")
        self._asec_cb.bind("<<ComboboxSelected>>", lambda e: self._render_assumptions())

        ttk.Button(top, text="✏  Edit Assumptions",
                   style="Edit.TButton",
                   command=self._open_edit_dialog).pack(side="right", padx=4)

        content = tk.Frame(parent, bg=C["bg"])
        content.pack(fill="both", expand=True, padx=14, pady=(10, 10))

        self._assum_cv = tk.Canvas(content, bg=C["bg"], highlightthickness=0)
        asb = ttk.Scrollbar(content, orient="vertical", command=self._assum_cv.yview)
        self._assum_frame = tk.Frame(self._assum_cv, bg=C["bg"])
        self._assum_frame.bind("<Configure>",
            lambda e: self._assum_cv.configure(
                scrollregion=self._assum_cv.bbox("all")))
        self._assum_cv.create_window((0, 0), window=self._assum_frame, anchor="nw")
        self._assum_cv.configure(yscrollcommand=asb.set)
        self._assum_cv.pack(side="left", fill="both", expand=True)
        asb.pack(side="right", fill="y")

        self._render_assumptions()

    def _refresh_assumptions_tab(self):
        """Rebuild the assumption tab after edit (update dropdown values too)."""
        self._asec_cb["values"] = list(self._assumptions_text.keys())
        if self.v_asec.get() not in self._assumptions_text:
            self.v_asec.set(list(self._assumptions_text.keys())[0])
        self._render_assumptions()

    def _render_assumptions(self):
        for w in self._assum_frame.winfo_children():
            w.destroy()
        section = self.v_asec.get()
        rows    = self._assumptions_text.get(section, [])

        title_f = tk.Frame(self._assum_frame, bg=C["hdr_bg"], padx=16, pady=12)
        title_f.pack(fill="x", pady=(0, 12))
        tk.Label(title_f, text=f"📋  {section}",
                 bg=C["hdr_bg"], fg=C["accent"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(title_f,
                 text="Source: Book1_R2.xlsx  |  FY 2024-25  |  Editable via ✏ Edit Assumptions",
                 bg=C["hdr_bg"], fg=C["subtext"],
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")

        hdr_f = tk.Frame(self._assum_frame, bg=C["border"], pady=6)
        hdr_f.pack(fill="x")
        for txt, w in [("  Particulars", 38), ("  Rate / Value (2024-25)", 32),
                       ("  Remarks / Reference", 40)]:
            tk.Label(hdr_f, text=txt, bg=C["border"], fg=C["accent"],
                     font=("Segoe UI", 9, "bold"), width=w, anchor="w",
                     padx=4).pack(side="left")

        row_bgs = [C["card"], "#192D42"]
        for idx, row_data in enumerate(rows):
            bg = row_bgs[idx % 2]
            rf = tk.Frame(self._assum_frame, bg=bg, pady=4)
            rf.pack(fill="x")
            cols = [38, 32, 40]
            wrap = [300, 260, 320]
            fgs  = [C["text"], C["success"], C["subtext"]]
            fts  = [("Segoe UI",9), ("Segoe UI",9,"bold"), ("Segoe UI",8,"italic")]
            for c_idx, (w, wr, fg, ft) in enumerate(zip(cols, wrap, fgs, fts)):
                val = row_data[c_idx] if c_idx < len(row_data) else ""
                tk.Label(rf, text=f"  {val}" if val else "", bg=bg, fg=fg,
                         font=ft, width=w, anchor="w", padx=4,
                         wraplength=wr, justify="left").pack(side="left", anchor="n")

        note = tk.Frame(self._assum_frame, bg=C["bg"], pady=10)
        note.pack(fill="x")
        tk.Label(note,
                 text=("  ℹ  Click  ✏ Edit Assumptions  to modify any value. "
                       "Changes save to landed_cost_assumptions.json and apply instantly.\n"
                       "  All rates from FY 2024-25 tariff orders. Verify before use."),
                 bg=C["bg"], fg=C["border"],
                 font=("Segoe UI", 8, "italic"),
                 justify="left", wraplength=900).pack(anchor="w")

    # ─────────────────────────────────────────────────────────
    #  CALCULATION
    # ─────────────────────────────────────────────────────────
    def calculate(self, _=None):
        try:
            bc  = float(self.v_bcost.get())
            lmc = float(self.v_lmc.get())
        except ValueError:
            self.lbl_val.config(text="ERR", fg=C["red"])
            self.lbl_case.config(text="⚠  Enter numeric values for Base Cost and LMC")
            return

        voltage = ">33kV" if self.v_voltage.get() == ">33kV" else "≤33kV"
        lcoe, breakdown, err = compute_landed_cost(
            self.v_ptype.get(), bc,
            self.v_inj_st.get(), self.v_inj_conn.get(),
            self.v_wth_st.get(), self.v_wth_conn.get(),
            voltage, self.v_mode.get(), lmc, self.v_oadr.get(),
            self._rates)

        if err:
            self.lbl_val.config(text="N/A", fg=C["red"])
            self.lbl_case.config(text=f"⚠  {err}")
            self._clear_tbl(); return

        self.lbl_val.config(text=f"{lcoe:.4f}", fg=C["success"])
        self.lbl_case.config(
            text=(f"{self.v_inj_st.get()} [{self.v_inj_conn.get()}]  →  "
                  f"{self.v_wth_st.get()} [{self.v_wth_conn.get()}]"
                  f"    ·  {self.v_mode.get()}  ·  {self.v_oadr.get()}"
                  f"  ·  {self.v_ptype.get()}  ·  {voltage}"))
        self._render_tbl(bc, lcoe, breakdown)

    def _clear_tbl(self):
        for w in self.tbl.winfo_children(): w.destroy()
        tk.Label(self.tbl, text="No results — adjust parameters.",
                 bg=C["bg"], fg=C["subtext"],
                 font=("Segoe UI", 10, "italic")).pack(pady=30)

    def _render_tbl(self, bc, lcoe, breakdown):
        for w in self.tbl.winfo_children(): w.destroy()

        hdr = tk.Frame(self.tbl, bg=C["border"], pady=5)
        hdr.pack(fill="x")
        for txt, w, a in [("  Charge Component",42,"w"),
                           ("Rs./kWh",12,"e"),("Ps./kWh",12,"e"),("Share %",9,"e")]:
            tk.Label(hdr, text=txt, bg=C["border"], fg=C["accent"],
                     font=("Segoe UI",9,"bold"), width=w, anchor=a, padx=6).pack(side="left")

        row_bgs = [C["card"], "#192D42"]
        for idx, (name, val) in enumerate(breakdown.items()):
            bg = row_bgs[idx % 2]
            pct = val / lcoe * 100 if lcoe else 0
            pct_fg = C["accent"] if pct < 15 else C["accent2"] if pct < 40 else C["red"]
            row = tk.Frame(self.tbl, bg=bg, pady=3)
            row.pack(fill="x")
            tk.Label(row, text=f"  {name}", bg=bg, fg=C["text"],
                     font=("Segoe UI",9), anchor="w", width=42, padx=2).pack(side="left")
            tk.Label(row, text=f"{val:.4f}", bg=bg, fg=C["text"],
                     font=("Segoe UI",9,"bold"), anchor="e", width=12, padx=6).pack(side="left")
            tk.Label(row, text=f"{val*100:.2f}", bg=bg, fg=C["subtext"],
                     font=("Segoe UI",9), anchor="e", width=12, padx=6).pack(side="left")
            tk.Label(row, text=f"{pct:.1f}%", bg=bg, fg=pct_fg,
                     font=("Segoe UI",9), anchor="e", width=9, padx=8).pack(side="left")

        tk.Frame(self.tbl, bg=C["accent"], height=1).pack(fill="x", pady=4)
        bc_rs = bc / 100
        self._srow("Base Power Cost",    bc_rs,      bc,          bc_rs/lcoe*100,  C["subtext"])
        self._srow("TOTAL LANDED COST",  lcoe,       lcoe*100,    100.0,           C["success"], True)

        foot = tk.Frame(self.tbl, bg=C["bg"], pady=6)
        foot.pack(fill="x")
        tk.Label(foot,
                 text="  Formula: Landed Cost = Σ(mults × tariff rates) + Base Cost  "
                      "·  24 route cases verified against Excel Main Sheet",
                 bg=C["bg"], fg=C["border"],
                 font=("Segoe UI",7,"italic"), justify="left").pack(anchor="w")

    def _srow(self, label, rs, ps, pct, fg, bold=False):
        fnt = ("Segoe UI",10,"bold") if bold else ("Segoe UI",9)
        row = tk.Frame(self.tbl, bg=C["panel"], pady=3)
        row.pack(fill="x")
        tk.Label(row, text=f"  {label}", bg=C["panel"], fg=fg, font=fnt,
                 anchor="w", width=42, padx=2).pack(side="left")
        tk.Label(row, text=f"{rs:.4f}", bg=C["panel"], fg=fg, font=fnt,
                 anchor="e", width=12, padx=6).pack(side="left")
        tk.Label(row, text=f"{ps:.2f}", bg=C["panel"], fg=fg, font=fnt,
                 anchor="e", width=12, padx=6).pack(side="left")
        tk.Label(row, text=f"{pct:.1f}%", bg=C["panel"], fg=fg, font=fnt,
                 anchor="e", width=9, padx=8).pack(side="left")

    def reset(self):
        self.v_ptype.set("Thermal");   self.v_bcost.set("400")
        self.v_inj_st.set("MAHARASHTRA"); self.v_inj_conn.set("CTU")
        self.v_voltage.set(">33kV");   self.v_wth_st.set("GUJARAT")
        self.v_wth_conn.set("CTU");    self.v_mode.set("CAPTIVE")
        self.v_lmc.set(str(self._rates.get("gj_lmc_default", 30)))
        self.v_oadr.set("LTOA");       self.calculate()


if __name__ == "__main__":
    app = LandedCostApp()
    app.mainloop()
