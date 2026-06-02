from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


BASE_COLUMNS = [
    "Doc Pcte",
    "Fech Doc",
    "Convenio",
    "Regimen",
    "APB",
    "Saldo Apb",
    "Fecha Radicado",
    "Tipo Contrato",
    "MES",
]

MONTH_ORDER = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]

DATA_DIR = Path("dashboard_data")
LATEST_FILE = DATA_DIR / "BASE_RADICACION_ULTIMA_CARGA.xlsx"


st.set_page_config(
    page_title="Dashboard Radicacion",
    layout="wide",
    initial_sidebar_state="expanded",
)


def css() -> None:
    st.markdown(
        """
        <style>
        :root {
          --bg:#071823; --panel:#102733; --panel2:#0c202b; --line:#244652;
          --ink:#edf7f8; --muted:#9ab0b8; --blue:#0bb7b9; --amber:#f7b718;
        }
        .stApp { background:#071823; color:var(--ink); }
        [data-testid="stSidebar"] { background:#0c202b; border-right:1px solid var(--line); }
        h1, h2, h3 { color:#edf7f8; letter-spacing:0; }
        .muted { color:var(--muted); font-size:13px; }
        .kpi {
          background:linear-gradient(135deg,#073b43,#0c202b);
          border:1px solid #244652;
          padding:16px;
          min-height:105px;
        }
        .kpi span { display:block; color:#fff; font-size:12px; font-weight:800; text-transform:uppercase; }
        .kpi strong { display:block; color:#fff; font-size:25px; margin-top:8px; }
        .kpi em { display:block; color:#ff3154; font-style:normal; font-weight:700; margin-top:6px; }
        .panel {
          background:#102733;
          border:1px solid #244652;
          padding:14px;
          min-height:230px;
        }
        .bar-row {
          display:grid;
          grid-template-columns:minmax(170px, 1.4fr) 2fr 140px;
          gap:10px;
          align-items:center;
          min-height:30px;
        }
        .bar-label { overflow:hidden; white-space:nowrap; text-overflow:ellipsis; font-size:13px; }
        .bar-track { height:17px; background:#324b52; border:1px solid #405e66; }
        .bar-fill { height:100%; background:linear-gradient(90deg,#f7b718,#0bb7b9); }
        .bar-value { text-align:right; font-variant-numeric:tabular-nums; font-size:13px; font-weight:800; }
        .rank-row {
          display:grid;
          grid-template-columns:34px 1fr auto;
          gap:10px;
          align-items:center;
          padding:11px 0;
          border-bottom:1px solid rgba(255,255,255,.08);
        }
        .rank-row b {
          width:28px; height:28px; border-radius:50%;
          display:grid; place-items:center;
          background:#f7b718; color:#061823;
        }
        .rank-row span { overflow:hidden; white-space:nowrap; text-overflow:ellipsis; font-weight:700; }
        .rank-row strong { font-size:13px; }
        .month-chart {
          height:245px;
          display:flex;
          align-items:end;
          gap:10px;
          border-bottom:1px solid #244652;
          padding:18px 8px 0;
        }
        .month-col { flex:1; display:flex; flex-direction:column; align-items:center; min-width:45px; }
        .month-value {
          color:#9ab0b8;
          font-size:11px;
          margin-bottom:6px;
          writing-mode:vertical-rl;
          transform:rotate(180deg);
          max-height:78px;
        }
        .month-bar { width:72%; min-width:24px; background:linear-gradient(180deg,#f7b718,#0bb7b9); }
        .month-label { margin-top:8px; color:#d5e4e7; font-weight:800; }
        .stDownloadButton button, .stButton button {
          background:#f7b718 !important;
          color:#071823 !important;
          border:0 !important;
          font-weight:800 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def money(value: float) -> str:
    return "$" + f"{float(value or 0):,.0f}".replace(",", ".")


def number(value: float) -> str:
    return f"{float(value or 0):,.0f}".replace(",", ".")


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def load_dataframe(file_or_path) -> pd.DataFrame:
    df = pd.read_excel(file_or_path, sheet_name="BASE")
    missing = [col for col in BASE_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError("Faltan columnas en la hoja BASE: " + ", ".join(missing))

    df = df[BASE_COLUMNS].copy()
    df["Saldo Apb"] = pd.to_numeric(df["Saldo Apb"], errors="coerce").fillna(0)
    df["Fecha Radicado"] = pd.to_datetime(df["Fecha Radicado"], errors="coerce")
    df["Fech Doc"] = pd.to_datetime(df["Fech Doc"], errors="coerce", dayfirst=True)
    for col in ["Doc Pcte", "Convenio", "Regimen", "Tipo Contrato", "MES"]:
        df[col] = df[col].map(normalize_text)
    df["APB"] = df["APB"].map(normalize_text)
    df["MES"] = df["MES"].str.lower()
    return df


@st.cache_data(show_spinner=False)
def load_from_disk(path: str, mtime: float) -> pd.DataFrame:
    return load_dataframe(path)


def get_data() -> pd.DataFrame | None:
    if "df" in st.session_state:
        return st.session_state["df"]
    if LATEST_FILE.exists():
        return load_from_disk(str(LATEST_FILE), LATEST_FILE.stat().st_mtime)
    return None


def save_uploaded(uploaded_file) -> pd.DataFrame:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    payload = uploaded_file.getvalue()
    LATEST_FILE.write_bytes(payload)
    df = load_dataframe(BytesIO(payload))
    st.session_state["df"] = df
    load_from_disk.clear()
    return df


def group_sum(df: pd.DataFrame, field: str, top: int | None = None) -> list[dict]:
    grouped = df.groupby(field, dropna=False)["Saldo Apb"].sum().sort_values(ascending=False).reset_index()
    if top:
        grouped = grouped.head(top)
    return [{"label": str(row[field]) or "Sin dato", "value": float(row["Saldo Apb"])} for _, row in grouped.iterrows()]


def month_summary(df: pd.DataFrame) -> list[dict]:
    rows = []
    for month in MONTH_ORDER:
        piece = df[df["MES"] == month]
        if len(piece) or month in set(df["MES"].dropna()):
            rows.append({"label": month.capitalize(), "value": float(piece["Saldo Apb"].sum()), "count": int(piece["Doc Pcte"].count())})
    return rows


def apply_filters(df: pd.DataFrame, filters: dict[str, str]) -> pd.DataFrame:
    out = df
    if filters.get("mes"):
        out = out[out["MES"] == filters["mes"]]
    if filters.get("regimen"):
        out = out[out["Regimen"] == filters["regimen"]]
    if filters.get("contrato"):
        out = out[out["Tipo Contrato"] == filters["contrato"]]
    if filters.get("apb"):
        out = out[out["APB"] == filters["apb"]]
    if filters.get("convenio"):
        needle = filters["convenio"].lower()
        out = out[out["Convenio"].str.lower().str.contains(needle, na=False)]
    return out


def build_summary(df: pd.DataFrame) -> dict:
    total = float(df["Saldo Apb"].sum())
    docs = int(df["Doc Pcte"].count())
    convenios = int(df["Convenio"].replace("", pd.NA).dropna().nunique())
    apbs = int(df["APB"].replace("", pd.NA).dropna().nunique())
    return {
        "kpis": {
            "saldo_total": total,
            "documentos": docs,
            "convenios": convenios,
            "apb": apbs,
            "promedio_documento": total / docs if docs else 0,
        },
        "months": month_summary(df),
        "top_convenios": group_sum(df, "Convenio", 10),
        "top_apb": group_sum(df, "APB", 10),
        "regimen": group_sum(df, "Regimen"),
        "tipo_contrato": group_sum(df, "Tipo Contrato"),
    }


def bars(items: list[dict], limit: int = 10) -> str:
    rows = items[:limit]
    if not rows:
        return "<p class='muted'>Sin datos</p>"
    max_value = max(float(item["value"]) for item in rows) or 1
    html_rows = []
    for item in rows:
        width = max(4, min(100, float(item["value"]) / max_value * 100))
        html_rows.append(
            f"<div class='bar-row'>"
            f"<div class='bar-label' title='{item['label']}'>{item['label']}</div>"
            f"<div class='bar-track'><div class='bar-fill' style='width:{width:.1f}%'></div></div>"
            f"<div class='bar-value'>{money(item['value'])}</div>"
            f"</div>"
        )
    return "".join(html_rows)


def month_chart(items: list[dict]) -> str:
    if not items:
        return "<p class='muted'>Sin datos por mes</p>"
    max_value = max(float(item["value"]) for item in items) or 1
    cols = []
    for item in items:
        height = max(8, min(165, float(item["value"]) / max_value * 165))
        cols.append(
            f"<div class='month-col'>"
            f"<div class='month-value'>{money(item['value'])}</div>"
            f"<div class='month-bar' style='height:{height:.0f}px'></div>"
            f"<div class='month-label'>{item['label'][:3]}</div>"
            f"</div>"
        )
    return "<div class='month-chart'>" + "".join(cols) + "</div>"


def kpi_card(label: str, value: str, foot: str = "") -> None:
    st.markdown(f"<div class='kpi'><span>{label}</span><strong>{value}</strong><em>{foot}</em></div>", unsafe_allow_html=True)


def make_excel_report(df: pd.DataFrame, summary: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws["A1"] = "Reporte de Radicacion"
    ws["A1"].font = Font(size=18, bold=True, color="1F4E78")
    ws["A2"] = f"Generado: {datetime.now():%Y-%m-%d %H:%M}"
    ws["A4"] = "Indicador"
    ws["B4"] = "Valor"
    for idx, (label, value, fmt) in enumerate(
        [
            ("Saldo total", summary["kpis"]["saldo_total"], "$#,##0"),
            ("Documentos", summary["kpis"]["documentos"], "#,##0"),
            ("Convenios", summary["kpis"]["convenios"], "#,##0"),
            ("APB", summary["kpis"]["apb"], "#,##0"),
            ("Promedio documento", summary["kpis"]["promedio_documento"], "$#,##0"),
        ],
        start=5,
    ):
        ws[f"A{idx}"] = label
        ws[f"B{idx}"] = value
        ws[f"B{idx}"].number_format = fmt

    tables = [("D4", "Mes", summary["months"]), ("G4", "Top Convenios", summary["top_convenios"]), ("J4", "Top APB", summary["top_apb"])]
    for anchor, title, rows in tables:
        col = ws[anchor].column
        row = ws[anchor].row
        ws.cell(row=row, column=col, value=title)
        ws.cell(row=row + 1, column=col, value="Categoria")
        ws.cell(row=row + 1, column=col + 1, value="Saldo")
        for idx, item in enumerate(rows, start=row + 2):
            ws.cell(row=idx, column=col, value=item["label"])
            ws.cell(row=idx, column=col + 1, value=item["value"]).number_format = "$#,##0"

    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ["A4", "D4", "G4", "J4"]:
        ws[cell].font = Font(bold=True, color="FFFFFF")
        ws[cell].fill = fill
    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 18

    line = LineChart()
    line.title = "Saldo por mes"
    line.add_data(Reference(ws, min_col=5, min_row=5, max_row=5 + len(summary["months"])), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=4, min_row=6, max_row=5 + len(summary["months"])))
    ws.add_chart(line, "A12")

    bar = BarChart()
    bar.type = "bar"
    bar.title = "Top convenios"
    bar.add_data(Reference(ws, min_col=8, min_row=5, max_row=15), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=7, min_row=6, max_row=15))
    ws.add_chart(bar, "H12")

    base = wb.create_sheet("BASE")
    base.append(BASE_COLUMNS)
    for row in df.itertuples(index=False):
        base.append([x.to_pydatetime() if hasattr(x, "to_pydatetime") else x for x in row])
    ref = f"A1:{get_column_letter(len(BASE_COLUMNS))}{base.max_row}"
    tab = Table(displayName="Tabla_Base_Radicacion", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    base.add_table(tab)
    base.freeze_panes = "A2"
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def pdf_text(value: object) -> str:
    return ("" if value is None else str(value)).encode("latin-1", "replace").decode("latin-1")


def pdf_money(value: float) -> str:
    return "$" + f"{float(value or 0):,.0f}"


def pdf_table(title: str, rows: list[list], widths: list[float] | None = None) -> list:
    styles = getSampleStyleSheet()
    elements = [Paragraph(pdf_text(title), styles["Heading3"])]
    table = PdfTable([[pdf_text(cell) for cell in row] for row in rows], colWidths=widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b3b43")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#c8d3d7")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef5f6")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
            ]
        )
    )
    elements.extend([table, Spacer(1, 0.18 * inch)])
    return elements


def make_pdf_report(df: pd.DataFrame, summary: dict, filters: dict[str, str]) -> bytes:
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(letter),
        leftMargin=0.35 * inch,
        rightMargin=0.35 * inch,
        topMargin=0.35 * inch,
        bottomMargin=0.35 * inch,
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Reporte de Radicacion", styles["Title"]),
        Paragraph(
            pdf_text(
                f"Generado: {datetime.now():%Y-%m-%d %H:%M} | Filtros: "
                + (", ".join(f"{k}: {v}" for k, v in filters.items() if v) or "Todos")
            ),
            styles["Normal"],
        ),
        Spacer(1, 0.14 * inch),
    ]
    k = summary["kpis"]
    story.extend(
        pdf_table(
            "Indicadores",
            [
                ["Indicador", "Valor"],
                ["Saldo total", pdf_money(k["saldo_total"])],
                ["Documentos", number(k["documentos"])],
                ["Convenios", number(k["convenios"])],
                ["APB", number(k["apb"])],
                ["Promedio documento", pdf_money(k["promedio_documento"])],
            ],
            [2.2 * inch, 2.0 * inch],
        )
    )
    story.extend(pdf_table("Saldo por mes", [["Mes", "Saldo", "Documentos"]] + [[i["label"], pdf_money(i["value"]), number(i["count"])] for i in summary["months"]], [1.5 * inch, 1.8 * inch, 1.2 * inch]))
    story.extend(pdf_table("Top convenios", [["Convenio", "Saldo"]] + [[i["label"][:70], pdf_money(i["value"])] for i in summary["top_convenios"]], [5.0 * inch, 1.6 * inch]))
    story.extend(pdf_table("Top APB", [["APB", "Saldo"]] + [[i["label"], pdf_money(i["value"])] for i in summary["top_apb"]], [1.6 * inch, 1.8 * inch]))
    detail = df.sort_values("Saldo Apb", ascending=False).head(35)
    rows = [["Documento", "Fecha", "Convenio", "Regimen", "APB", "Contrato", "Mes", "Saldo"]]
    for _, row in detail.iterrows():
        rows.append(
            [
                row["Doc Pcte"],
                row["Fecha Radicado"].strftime("%Y-%m-%d") if not pd.isna(row["Fecha Radicado"]) else "",
                row["Convenio"][:38],
                row["Regimen"][:32],
                row["APB"],
                row["Tipo Contrato"],
                str(row["MES"]).capitalize(),
                pdf_money(row["Saldo Apb"]),
            ]
        )
    story.extend(pdf_table("Detalle principal", rows, [0.85 * inch, 0.75 * inch, 2.2 * inch, 1.8 * inch, 0.65 * inch, 0.75 * inch, 0.55 * inch, 1.1 * inch]))
    doc.build(story)
    return output.getvalue()


def admin_panel() -> None:
    st.sidebar.markdown("### Administracion")
    admin_password = st.secrets.get("ADMIN_PASSWORD", "tu-clave-segura")
    typed = st.sidebar.text_input("Clave para cargar datos", type="password")
    if not typed:
        st.sidebar.caption("La vista publica puede filtrar y descargar, pero no cargar datos.")
        return
    if typed != admin_password:
        st.sidebar.error("Clave incorrecta. Revisa el valor ADMIN_PASSWORD en los secrets de Streamlit.")
        st.sidebar.caption("La vista publica puede filtrar y descargar, pero no cargar datos.")
        return
    st.sidebar.success("Clave correcta. Ya puedes cargar datos.")
    uploaded = st.sidebar.file_uploader("Cargar BASE_RADICACION actualizada", type=["xlsx", "xls"])
    if uploaded:
        try:
            df = save_uploaded(uploaded)
            st.sidebar.success(f"Base cargada: {len(df):,} registros")
            st.rerun()
        except Exception as exc:
            st.sidebar.error(str(exc))


def render_dashboard(df_raw: pd.DataFrame) -> None:
    st.sidebar.markdown("### Filtros")
    months_present = set(df_raw["MES"].dropna().tolist())
    month_options = ["Todos"] + [m.capitalize() for m in MONTH_ORDER if m in months_present]
    regimen_options = ["Todos"] + sorted([x for x in df_raw["Regimen"].dropna().unique().tolist() if x])
    contrato_options = ["Todos"] + sorted([x for x in df_raw["Tipo Contrato"].dropna().unique().tolist() if x])
    apb_options = ["Todos"] + sorted([x for x in df_raw["APB"].dropna().unique().tolist() if x])

    mes_label = st.sidebar.selectbox("Mes", month_options)
    regimen = st.sidebar.selectbox("Regimen", regimen_options)
    contrato = st.sidebar.selectbox("Tipo contrato", contrato_options)
    apb = st.sidebar.selectbox("APB", apb_options)
    convenio = st.sidebar.text_input("Buscar convenio")

    filters = {
        "mes": "" if mes_label == "Todos" else mes_label.lower(),
        "regimen": "" if regimen == "Todos" else regimen,
        "contrato": "" if contrato == "Todos" else contrato,
        "apb": "" if apb == "Todos" else apb,
        "convenio": convenio.strip(),
    }
    df = apply_filters(df_raw, filters)
    summary = build_summary(df)
    k = summary["kpis"]

    st.title("Dashboard de Radicacion")
    st.markdown("<p class='muted'>Vista publica: las personas pueden filtrar, navegar y descargar reportes. Solo quien tenga clave puede cargar datos.</p>", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Saldo total", money(k["saldo_total"]), "con filtros activos")
    with c2:
        kpi_card("Documentos", number(k["documentos"]), "registros filtrados")
    with c3:
        kpi_card("Convenios", number(k["convenios"]), "entidades distintas")
    with c4:
        kpi_card("Promedio doc.", money(k["promedio_documento"]), "saldo / documentos")

    excel_bytes = make_excel_report(df, summary)
    pdf_bytes = make_pdf_report(df, summary, filters)
    d1, d2 = st.columns([1, 1])
    with d1:
        st.download_button("Descargar PDF", pdf_bytes, "reporte_radicacion.pdf", "application/pdf")
    with d2:
        st.download_button("Descargar Excel", excel_bytes, "reporte_radicacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Resumen", "Convenios", "APB", "Regimen y contrato", "Base"])
    with tab1:
        left, right = st.columns([2, 1])
        with left:
            st.markdown("<div class='panel'><h3>Saldo por mes</h3>" + month_chart(summary["months"]) + "</div>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<div class='panel'><h3>Top convenios</h3>" + bars(summary["top_convenios"], 10) + "</div>", unsafe_allow_html=True)
        with right:
            rank = []
            for idx, item in enumerate(summary["top_convenios"][:5], start=1):
                rank.append(
                    f"<div class='rank-row'><b>{idx}</b>"
                    f"<span>{item['label']}</span><strong>{money(item['value'])}</strong></div>"
                )
            st.markdown("<div class='panel'><h3>Top 5 entidades</h3>" + "".join(rank) + "</div>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<div class='panel'><h3>Tipo de contrato</h3>" + bars(summary["tipo_contrato"], 10) + "</div>", unsafe_allow_html=True)
    with tab2:
        st.markdown("<div class='panel'><h3>Convenios por saldo</h3>" + bars(summary["top_convenios"], 10) + "</div>", unsafe_allow_html=True)
    with tab3:
        st.markdown("<div class='panel'><h3>APB por saldo</h3>" + bars(summary["top_apb"], 10) + "</div>", unsafe_allow_html=True)
    with tab4:
        left, right = st.columns(2)
        with left:
            st.markdown("<div class='panel'><h3>Regimen</h3>" + bars(summary["regimen"], 20) + "</div>", unsafe_allow_html=True)
        with right:
            st.markdown("<div class='panel'><h3>Tipo de contrato</h3>" + bars(summary["tipo_contrato"], 20) + "</div>", unsafe_allow_html=True)
    with tab5:
        view = df.sort_values("Saldo Apb", ascending=False).head(500).copy()
        view["Saldo Apb"] = view["Saldo Apb"].map(money)
        st.dataframe(view, use_container_width=True, height=520)


def main() -> None:
    css()
    admin_panel()
    df = get_data()
    if df is None:
        st.title("Dashboard de Radicacion")
        st.info("Aun no hay una base cargada. Ingresa la clave de administracion en el panel lateral y carga el archivo BASE_RADICACION.xlsx.")
        return
    render_dashboard(df)


if __name__ == "__main__":
    main()
