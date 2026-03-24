"""
oa run test — 26 markets across 2 templates (Python format).

Usage:
    oa run test_runfile.py
    oa run test_runfile.py --dry-run
    oa run test_runfile.py -v
"""

DATAPATH = "../excel_test_data"

jobs = {
    # Argentina through Korea → template 1
    "test_runfile_template_1.pptx": {
        "Argentina":      f"{DATAPATH}/rpm_tracking_Argentina_(05_07).xlsx",
        "Australia":      f"{DATAPATH}/rpm_tracking_Australia_(05_15).xlsx",
        "Brazil":         f"{DATAPATH}/rpm_tracking_Brazil_(05_07).xlsx",
        "Colombia":       f"{DATAPATH}/rpm_tracking_Colombia_(05_07).xlsx",
        "Czech_Republic": f"{DATAPATH}/rpm_tracking_Czech_Republic_(05_07).xlsx",
        "France":         f"{DATAPATH}/rpm_tracking_France_(05_07).xlsx",
        "Germany":        f"{DATAPATH}/rpm_tracking_Germany_(05_07).xlsx",
        "Greece":         f"{DATAPATH}/rpm_tracking_Greece_(05_07).xlsx",
        "Indonesia":      f"{DATAPATH}/rpm_tracking_Indonesia_(05_07).xlsx",
        "Italy":          f"{DATAPATH}/rpm_tracking_Italy_(05_07).xlsx",
        "Japan":          f"{DATAPATH}/rpm_tracking_Japan_(05_07).xlsx",
        "Kazakhstan":     f"{DATAPATH}/rpm_tracking_Kazakhstan_(05_15).xlsx",
        "Korea":          f"{DATAPATH}/rpm_tracking_Korea_(05_07).xlsx",
    },

    # Malaysia through United_States → template 2
    "test_runfile_template_2.pptx": {
        "Malaysia":       f"{DATAPATH}/rpm_tracking_Malaysia_(05_07).xlsx",
        "Mexico":         f"{DATAPATH}/rpm_tracking_Mexico_(05_07).xlsx",
        "New_Zealand":    f"{DATAPATH}/rpm_tracking_New_Zealand_(05_07).xlsx",
        "Phillippines":   f"{DATAPATH}/rpm_tracking_Phillippines_(05_07).xlsx",
        "Poland":         f"{DATAPATH}/rpm_tracking_Poland_(05_07).xlsx",
        "Portugal":       f"{DATAPATH}/rpm_tracking_Portugal_(05_09).xlsx",
        "Romania":        f"{DATAPATH}/rpm_tracking_Romania_(05_07).xlsx",
        "Serbia":         f"{DATAPATH}/rpm_tracking_Serbia_(05_07).xlsx",
        "South_Africa":   f"{DATAPATH}/rpm_tracking_South_Africa_(05_07).xlsx",
        "Spain":          f"{DATAPATH}/rpm_tracking_Spain_(05_07).xlsx",
        "Taiwan":         f"{DATAPATH}/rpm_tracking_Taiwan_(05_07).xlsx",
        "United_Kingdom": f"{DATAPATH}/rpm_tracking_United_Kingdom_(05_07).xlsx",
        "United_States":  f"{DATAPATH}/rpm_tracking_United_States_(05_07).xlsx",
    },
}

default_output = "runfile_output/{name}.pptx"

steps = ["links", "tables", "deltas", "coloring", "charts"]

config = {}
