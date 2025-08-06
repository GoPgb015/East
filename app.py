"""
sales-pivot / app.py
Dark-theme dashboard – Schneider Electric palette
• Dashboard (6 charts) → Pivot-2 → Pivot-1
"""

from flask import Flask, render_template, request, send_file
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import io
import tempfile

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# ──────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────
ID_COLS   = ["Cluster", "Employee Responsible", "Helios Code"]
PIV1_DEC  = 2
EMP_SCALE = 1          # rupees → millions  (set to 1 ⇒ no scaling)
EMP_DEC   = 2
TOTAL_HDR = "Total (mn)"

# Schneider-Electric colours
SE_GREEN  = "#00B140"
LIGHT_GRN = "#6AFFB4"
MID_GRN   = "#7BCB8C"
DARK_GRN  = "#00582B"
GREY      = "#A8A8A8"

EMPLOYEE_CLUSTER_MAP = {
    # CG
    "Pritesh": "CG", "Abinash": "CG", "Mayur": "CG",
    "TBH": "CG", "Arti": "CG",
    # OD
    "Arun": "OD", "Abhishek": "OD", "Bodhis": "OD", "Mahesh": "OD",
    "MD FAZLE MURSHED": "OD",          # ← NEW
    # B+R
    "Rahul": "B+R", "Harsh": "B+R", "Vikash": "B+R", "Nagender": "B+R",
    # JSR
    "Sunny": "JSR",
    "Priyanka Auddy": "JSR",           # ← NEW
}

DEFAULT_CLUSTER = "OD"

TABLE_CLASS = "table table-bordered table-hover display nowrap"

# ──────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────
def plot_div(fig, height=400):
    return pio.to_html(
        fig, full_html=False, include_plotlyjs="cdn",
        config={"displaylogo": False}, default_height=height
    )

def dark_template(fig):
    fig.update_layout(template="plotly_dark",
                      paper_bgcolor="#000", plot_bgcolor="#000",
                      font_color="#e0e0e0")
    return fig

# ──────────────────────────────────────────────────────────────────
# ROUTES
# ──────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    # 1. read Excel files
    ob_file = request.files.get("ob_file")
    sales_file = request.files.get("file")
    targets_file = request.files.get("targets_file")

    if not ob_file or ob_file.filename == "":
        return "⚠️ No OB file selected.", 400
    if not sales_file or sales_file.filename == "":
        return "⚠️ No sales file selected.", 400
    if not targets_file or targets_file.filename == "":
        return "⚠️ No targets file selected.", 400

    try:
        df_sales = pd.read_excel(io.BytesIO(sales_file.read()), sheet_name=0)
        df_ob = pd.read_excel(io.BytesIO(ob_file.read()), sheet_name=0)
        # Read targets from the template using pd.ExcelFile to read multiple sheets
        with pd.ExcelFile(targets_file) as xls:
            emp_targets_df = pd.read_excel(xls, sheet_name="Employee Targets")
            bu_targets_df = pd.read_excel(xls, sheet_name="BU Targets")
        # Validate and clean target values
        emp_targets_df["Target"] = pd.to_numeric(emp_targets_df["Target"], errors="coerce")
        bu_targets_df["Target"] = pd.to_numeric(bu_targets_df["Target"], errors="coerce")
        # Create dictionaries for easy lookup, filling NaNs with 0
        emp_targets = emp_targets_df.set_index("Employee Responsible")["Target"].fillna(0).to_dict()
        bu_targets = bu_targets_df.set_index("BU")["Target"].fillna(0).to_dict()
    except Exception as exc:
        return f"⚠️ Could not read Excel – {exc}", 400

    for df in [df_sales, df_ob]:
        df.columns = df.columns.str.strip()
        def get_cluster(employee_name_from_file):
            if isinstance(employee_name_from_file, str):
                for map_key, cluster in EMPLOYEE_CLUSTER_MAP.items():
                    if map_key in employee_name_from_file:
                        return cluster
            return DEFAULT_CLUSTER
        df['Cluster'] = df['Employee Responsible'].apply(get_cluster)

    missing = [c for c in ID_COLS if c not in df_sales.columns]
    if missing:
        return f"⚠️ Missing column(s) in sales: {', '.join(missing)}", 400
    missing_ob = [c for c in ID_COLS if c not in df_ob.columns]
    if missing_ob:
        return f"⚠️ Missing column(s) in OB: {', '.join(missing_ob)}", 400

    # ────────────────────────────────
    # ❶ HARD-WIRE THE SALES/OB COLUMN
    # ────────────────────────────────
    if "MINR-2025" not in df_sales.columns:
        return "⚠️ Column 'MINR-2025' not found in the sales sheet.", 400
    if "MINR-2025" not in df_ob.columns:
        return "⚠️ Column 'MINR-2025' not found in the OB sheet.", 400
    sales_col = "MINR-2025"
    ob_col = "MINR-2025"
    # ────────────────────────────────

    # ───────────────────────────────────────────────────────────────
    # PIVOT-2  (Employee totals & KPIs)
    # ───────────────────────────────────────────────────────────────
    # Group by cluster/employee for both sales and OB
    grp_sales = (
        df_sales.groupby(["Cluster", "Employee Responsible"])[sales_col]
        .sum()
        .reset_index()
    )
    grp_ob = (
        df_ob.groupby(["Cluster", "Employee Responsible"])[ob_col]
        .sum()
        .reset_index()
    )
    # Merge sales and OB on Cluster + Employee
    merged = pd.merge(grp_ob, grp_sales, on=["Cluster", "Employee Responsible"], how="outer", suffixes=("_ob", "_sales")).fillna(0)

    rows, cluster_rows, emp_rows = [], [], []
    for clust, sub in merged.groupby("Cluster", sort=False):
        subtotal_ob = sub[ob_col + "_ob"].sum()
        subtotal_sales = sub[sales_col + "_sales"].sum()
        total_ob_mn = round(subtotal_ob / EMP_SCALE, EMP_DEC)
        total_sales_mn = round(subtotal_sales / EMP_SCALE, EMP_DEC)
        cl_target = emp_targets.get(
            clust, sub["Employee Responsible"].map(emp_targets).sum(min_count=1)
        )
        cl_target = round(cl_target, EMP_DEC) if cl_target is not None else None
        cl_target_mn = round(cl_target / EMP_SCALE, EMP_DEC) if cl_target is not None else None
        remaining_ob_mn = round(cl_target_mn - total_ob_mn, EMP_DEC) if cl_target_mn is not None else None
        remaining_sales_mn = round(cl_target_mn - total_sales_mn, EMP_DEC) if cl_target_mn is not None else None
        achv_ob_pct = round(total_ob_mn / cl_target_mn * 100, 2) if cl_target_mn else None
        achv_sales_pct = round(total_sales_mn / cl_target_mn * 100, 2) if cl_target_mn else None
        subtotal_row = {
            "Cluster": clust, "Employee Responsible": f"{clust} Total",
            "OB Total (mn)": total_ob_mn, "OB Remaining": remaining_ob_mn, "OB Achiev %": achv_ob_pct,
            "Sales Total (mn)": total_sales_mn, "Sales Remaining": remaining_sales_mn, "Sales Achiev %": achv_sales_pct,
            "Target": cl_target_mn
        }
        rows.append(subtotal_row)
        cluster_rows.append(subtotal_row)
        for _, r in sub.iterrows():
            name = r["Employee Responsible"]
            tgt = emp_targets.get(name, 0)
            tgt_mn = round(tgt / EMP_SCALE, EMP_DEC) if tgt > 0 else 0
            ob_total = round(r[ob_col + "_ob"] / EMP_SCALE, EMP_DEC)
            sales_total = round(r[sales_col + "_sales"] / EMP_SCALE, EMP_DEC)
            ob_rem = round(tgt_mn - ob_total, EMP_DEC) if tgt_mn > 0 else 0
            sales_rem = round(tgt_mn - sales_total, EMP_DEC) if tgt_mn > 0 else 0
            ob_ach = round((ob_total / tgt_mn * 100), 2) if tgt_mn > 0 else 0
            sales_ach = round((sales_total / tgt_mn * 100), 2) if tgt_mn > 0 else 0
            emp_row = {
                "Name": name,
                "OB Total (mn)": ob_total, "OB Remaining": ob_rem, "OB Achiev %": ob_ach,
                "Sales Total (mn)": sales_total, "Sales Remaining": sales_rem, "Sales Achiev %": sales_ach,
                "Target": tgt_mn
            }
            emp_rows.append(emp_row)
            rows.append({"Cluster": clust, "Employee Responsible": name, **emp_row | {}})
    pivot2 = pd.DataFrame(rows)
    pivot2.drop(columns=['Name'], inplace=True, errors='ignore')
    def highlight_subtotal_row(row):
        is_subtotal = 'Total' in str(row['Employee Responsible'])
        return ['background-color: #00582B'] * len(row) if is_subtotal else [''] * len(row)
    styler = pivot2.style.apply(highlight_subtotal_row, axis=1).format({
        "OB Total (mn)": '{:,.2f}', "OB Remaining": '{:,.2f}', "OB Achiev %": '{:.2f}',
        "Sales Total (mn)": '{:,.2f}', "Sales Remaining": '{:,.2f}', "Sales Achiev %": '{:.2f}',
        'Target': '{:,.2f}'
    })
    pivot2_html = styler.to_html(classes=TABLE_CLASS, index=False, border=0, na_rep="")

    # ───────────────────────────────────────────────────────────────
    # BU Wise Performance Table
    # ───────────────────────────────────────────────────────────────
    def get_bu_from_helios(helios_code):
        if isinstance(helios_code, str):
            if helios_code.startswith("PP"):
                return "PP"
            elif helios_code.startswith("HD"):
                return "H&D"
            elif helios_code.startswith("ID"):
                return "IND"
            elif helios_code.startswith("SP") or helios_code.startswith("PS"):
                return "SP"
            elif helios_code.startswith("DP") or helios_code.startswith("DE"):
                return "DE"
        return "Other"
    df_sales["BU"] = df_sales["Helios Code"].apply(get_bu_from_helios)
    df_ob["BU"] = df_ob["Helios Code"].apply(get_bu_from_helios)
    bu_sales_df = df_sales.groupby("BU", as_index=False)[sales_col].sum()
    bu_sales_df.rename(columns={sales_col: "Sales"}, inplace=True)
    bu_ob_df = df_ob.groupby("BU", as_index=False)[ob_col].sum()
    bu_ob_df.rename(columns={ob_col: "OB"}, inplace=True)
    bu_performance_rows = []
    for bu_name, bu_target in bu_targets.items():
        sales_row = bu_sales_df[bu_sales_df["BU"] == bu_name]
        ob_row = bu_ob_df[bu_ob_df["BU"] == bu_name]
        sales = sales_row["Sales"].iloc[0] if not sales_row.empty else 0
        ob = ob_row["OB"].iloc[0] if not ob_row.empty else 0
        remaining_sales = bu_target - sales
        remaining_ob = bu_target - ob
        ach_sales = f"{round(sales / bu_target * 100, 2)}%" if bu_target else "0.00%"
        ach_ob = f"{round(ob / bu_target * 100, 2)}%" if bu_target else "0.00%"
        sales_mn = round(sales / EMP_SCALE, EMP_DEC)
        ob_mn = round(ob / EMP_SCALE, EMP_DEC)
        bu_target_mn = round(bu_target / EMP_SCALE, EMP_DEC)
        remaining_sales_mn = round(remaining_sales / EMP_SCALE, EMP_DEC)
        remaining_ob_mn = round(remaining_ob / EMP_SCALE, EMP_DEC)
        bu_performance_rows.append({
            "BU": bu_name,
            "OB (mn)": ob_mn,
            "OB Remaining (mn)": remaining_ob_mn,
            "ACH OB (%)": ach_ob,
            "Sales (mn)": sales_mn,
            "Sales Remaining (mn)": remaining_sales_mn,
            "ACH Sales (%)": ach_sales,
            "Target (mn)": bu_target_mn
        })
    bu_df = pd.DataFrame(bu_performance_rows)
    bu_table_html = bu_df.style.format({
        "OB (mn)": '{:,.2f}', "OB Remaining (mn)": '{:,.2f}', "ACH OB (%)": '{}',
        "Sales (mn)": '{:,.2f}', "Sales Remaining (mn)": '{:,.2f}', "ACH Sales (%)": '{}',
        "Target (mn)": '{:,.2f}'
    }).to_html(classes=TABLE_CLASS, index=False, border=0, na_rep="")

    # ─── Dashboard charts (6 graphs) ────────────────────────────────
    # Use only sales data for charts
    clusters  = [r["Cluster"] for r in cluster_rows]
    totals    = [r["Sales Total (mn)"] for r in cluster_rows]
    targets   = [r["Target"]  for r in cluster_rows]
    remain    = [r["Sales Remaining"] for r in cluster_rows]
    achv_pct  = [r["Sales Achiev %"]  for r in cluster_rows]
    dash_divs = []
    # 1. Totals vs Targets
    fig1 = go.Figure([
        go.Bar(name="Total (mn)", x=clusters, y=totals,  marker_color=SE_GREEN),
        go.Bar(name="Target",     x=clusters, y=targets, marker_color=MID_GRN)
    ])
    dark_template(fig1).update_layout(
        barmode="group", title="Cluster Total vs Target", legend_title_text=""
    )
    dash_divs.append(plot_div(fig1))
    # 2. Achievement %
    fig2 = go.Figure(go.Bar(
        x=clusters, y=achv_pct, text=achv_pct, textposition="outside",
        marker_color=SE_GREEN
    ))
    dark_template(fig2).update_layout(title="Achievement %", yaxis_title="%",
                                      xaxis_tickangle=-15)
    dash_divs.append(plot_div(fig2))
    # 3. Remaining to Target
    fig3 = go.Figure([
        go.Bar(name="Achieved",  x=clusters, y=totals,  marker_color=SE_GREEN),
        go.Bar(name="Remaining", x=clusters, y=remain, marker_color=LIGHT_GRN)
    ])
    dark_template(fig3).update_layout(
        barmode="stack", title="Remaining to Target", legend_title_text=""
    )
    dash_divs.append(plot_div(fig3))
    # 4. Pie share of totals
    fig4 = go.Figure(go.Pie(
        labels=clusters, values=totals,
        marker_colors=[SE_GREEN, MID_GRN, LIGHT_GRN, DARK_GRN, GREY][: len(clusters)],
        hole=0.4
    ))
    dark_template(fig4).update_layout(title="Share of Total Sales (mn)")
    dash_divs.append(plot_div(fig4, height=380))
    # 5. Scatter Target vs Achievement %
    fig5 = go.Figure(go.Scatter(
        x=targets, y=achv_pct, mode="markers+text",
        text=clusters, textposition="top center",
        marker=dict(size=12, color=SE_GREEN,
                    line=dict(color=DARK_GRN, width=1.5))
    ))
    dark_template(fig5).update_layout(
        title="Target vs Achievement %", xaxis_title="Target (mn)",
        yaxis_title="Achiev %"
    )
    dash_divs.append(plot_div(fig5))
    # 6. Employee performance bar
    emp_df = pd.DataFrame(emp_rows).dropna(subset=["Target"])
    emp_df = emp_df.sort_values("Sales Achiev %", ascending=False).head(12)  # top-12
    fig6 = go.Figure([
        go.Bar(name="Achieved",  x=emp_df["Name"], y=emp_df["Sales Total (mn)"],
               marker_color=SE_GREEN),
        go.Bar(name="Remaining", x=emp_df["Name"], y=emp_df["Sales Remaining"],
               marker_color=LIGHT_GRN)
    ])
    dark_template(fig6).update_layout(
        barmode="stack", title="Top Employee Performance (mn)",
        xaxis_tickangle=-25, legend_title_text=""
    )
    dash_divs.append(plot_div(fig6))
    # ───────────────────────────────────────────────────────────────
    # PIVOT-1  (Helios × Cluster | Employee)
    # ───────────────────────────────────────────────────────────────
    # Only sales data for pivot-1
    pivot1 = (
        df_sales.pivot_table(
            index="Helios Code",
            columns=["Cluster", "Employee Responsible"],
            values=sales_col, aggfunc="sum", fill_value=0
        )
    )
    total_row = pivot1.sum(numeric_only=True).to_frame().T
    total_row.index = ["Total"]
    pivot1 = pd.concat([pivot1, total_row])
    pivot1.index.name = "Helios Code"
    pivot1.columns = [f"{cl} | {emp}" for cl, emp in pivot1.columns]
    pivot1 = pivot1.reset_index()
    numeric_cols = pivot1.columns.drop('Helios Code')
    format_dict = {col: '{:,.2f}' for col in numeric_cols}
    pivot1_html = pivot1.style.format(format_dict, na_rep="0").to_html(
        classes=TABLE_CLASS, index=False, border=0
    )
    return render_template(
        "result.html",
        graphs=dash_divs,
        table2=pivot2_html,
        bu_table=bu_table_html,
        table1=pivot1_html
    )

# --------------------- template generator --------------------------
@app.route("/generate_target_template", methods=["POST"])
def generate_target_template():
    file = request.files.get("file")
    if not file: return "⚠️ No file selected.", 400
    try:
        df = pd.read_excel(io.BytesIO(file.read()), sheet_name=0)
    except Exception as exc:
        return f"⚠️ Could not read Excel – {exc}", 400
    df.columns = df.columns.str.strip()
    employees = sorted(df["Employee Responsible"].dropna().unique())
    emp_df = pd.DataFrame({"Employee Responsible": employees, "Target": [None]*len(employees)})
    bu_df  = pd.DataFrame({"BU": ["PP","H&D","IND","SP","PS","DE"], "Target": [None]*6})
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name) as xw:
            emp_df.to_excel(xw, "Employee Targets", index=False)
            bu_df.to_excel(xw, "BU Targets", index=False)
        tmp.seek(0)
        return send_file(tmp.name, as_attachment=True, download_name="target_template.xlsx")

# ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
