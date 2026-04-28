import requests
import pandas as pd
import sqlite3
import win32com.client
from dotenv import load_dotenv
import warnings
import os
from openpyxl import load_workbook
from datetime import date

TARGET_EXCEL = r"C:\Users\GastonVecchio\Documents\Stock\Stock depositos.xlsx"
EMAIL_RECEIVER = "gaston.vecchio@grupolargentina.com"
PRODUCT_COL   = "Familia"               # Column name for product names
STOCK_COL     = "Cobertura" 

today = date.today()

conn = sqlite3.connect(r"C:\Users\GastonVecchio\Documents\Code\Python\Stocks\inventory.db")
cursor = conn.cursor()


def _migrate_db():
    existing = [row[1] for row in cursor.execute("PRAGMA table_info(stockouts)").fetchall()]
    if "cobertura" not in existing:
        cursor.execute("ALTER TABLE stockouts ADD COLUMN cobertura REAL")
    if "status" not in existing:
        cursor.execute("ALTER TABLE stockouts ADD COLUMN status TEXT DEFAULT 'crash'")
    if "date_resolved" not in existing:
        cursor.execute("ALTER TABLE stockouts ADD COLUMN date_resolved TEXT")
    conn.commit()

_migrate_db()

procedure = input("Que queres hacer?\n1-Actualizar stocks\n2-Alertar faltantes\n3-Ver KPIs de stockouts\n")


def actualizar_stock():
    load_dotenv()
    
    METABASE_URL = "https://metabase-new.grupol.ar"
    USERNAME = os.getenv("USERNAME")
    PASSWORD = os.getenv("PASSWORD")
    print("Porfavor espera mientras se actualizan los stocks...")
    
    try:
    # Authenticate
        
        session = requests.Session()
        token = session.post(f"{METABASE_URL}/api/session", json={
            "username": USERNAME,
            "password": PASSWORD
        }).json()["id"]

        # Download the report
        response = session.post(
            f"{METABASE_URL}/api/card/201/query/xlsx",
            headers={"X-Metabase-Session": token}
        )

        # Save to a temp file and read it
        with open("temp_report.xlsx", "wb") as f:
            f.write(response.content)
        
    except:
        print("Error on Metabase connection")

    # Load data from the downloaded file
    df = pd.read_excel("temp_report.xlsx")

    stock_ciu = df[df["CtroDistrib"].str.contains("CIUDADELA")]
    stock_moreno = df[(df["Empresa"].str.contains("GRUPO L")) & ((df["CtroDistrib"].str.contains("CDM03 - MORENO 3")) | (df["CtroDistrib"].str.contains("CDM02 - MORENO 2")) | (df["CtroDistrib"].str.contains("CD MORENO")))]
    stock_atlantico = df[df["Empresa"].str.contains("SERVICIOS ATLANTICO SA")]

    with pd.ExcelWriter(TARGET_EXCEL, engine="openpyxl", mode="w") as writer:
        stock_ciu.to_excel(writer, sheet_name="STOCK CIUDADELA", index=False)
        stock_moreno.to_excel(writer, sheet_name="STOCK MORENO", index=False)
        stock_atlantico.to_excel(writer, sheet_name="STOCK ATLANTICO", index=False)

    ATLANTICO_PATH = r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables ATLANTICO.xlsx"
    VO_PATH = r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables VO.xlsx"

    #write_to_excel(ATLANTICO_PATH, "Stock", stock_atlantico, 0)
    #write_to_excel(VO_PATH, "Stock", stock_moreno, 0)
    #write_to_excel(VO_PATH, "Stock", stock_ciu, start_row=len(stock_moreno) + 2)

def write_to_excel(target_path, sheet_name, df, start_row=1):
    
    wb = load_workbook(target_path)
    ws = wb[sheet_name]

    if start_row == 1:  # only clear on first write
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=15):
            for cell in row:
                cell.value = None

    for i, (_, row) in enumerate(df.iterrows()):
        for j, value in enumerate(row.iloc[:15]):
            ws.cell(row=start_row + i, column=j + 1, value=value)

    wb.save(target_path)

def _sync_stockouts(df, company):
    df = df.copy()
    df[STOCK_COL] = pd.to_numeric(df[STOCK_COL], errors="coerce")
    df = df.dropna(subset=[STOCK_COL])

    open_db = pd.read_sql(
        "SELECT * FROM stockouts WHERE company = ? AND status != 'resolved'",
        conn, params=(company,)
    )
    open_products = set(open_db["product_name"].tolist())
    crashed = df[df[STOCK_COL] < 0.50]
    crashed_products = set(crashed[PRODUCT_COL].tolist())

    # Products that recovered: were open, now >= 50%
    for product in open_products - crashed_products:
        cursor.execute(
            "UPDATE stockouts SET status = 'resolved', date_resolved = ? "
            "WHERE product_name = ? AND company = ? AND status != 'resolved'",
            (str(today), product, company)
        )

    for _, row in crashed.iterrows():
        product = row[PRODUCT_COL]
        cob = row[STOCK_COL]
        new_status = "broke" if cob <= 0 else "crash"
        notes = None if pd.isna(row.get("Estado")) else row["Estado"]

        if product not in open_products:
            # New crash: insert
            cursor.execute(
                "INSERT INTO stockouts (product_name, date_of_stockout, notes, company, cobertura, status) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (product, str(today), notes, company, cob, new_status)
            )
        elif new_status == "broke":
            # Escalate existing crash to broke
            cursor.execute(
                "UPDATE stockouts SET status = 'broke', cobertura = ? "
                "WHERE product_name = ? AND company = ? AND status = 'crash'",
                (cob, product, company)
            )

    conn.commit()


def alertar_faltantes():

    def send_email(html_body, subject):
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_RECEIVER
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Send()

    def build_rows_html(low_df):
        rows = ""
        for _, row in low_df.iterrows():
            product = row[PRODUCT_COL]
            cobertura = f"{row[STOCK_COL] * 100:.2f}%"
            estado = "Pendiente" if pd.isna(row.get("Estado")) else row["Estado"]
            broke_style = ' style="background-color:#ffe0e0;"' if row[STOCK_COL] <= 0 else ""
            rows += f"""
            <tr{broke_style}>
                <td style="padding: 8px; border: 1px solid #ddd;">{product}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{cobertura}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{estado}</td>
            </tr>
        """
        return rows

    EMAIL_TEMPLATE = """
        <p>Este es un mail automatico, comprar los siguientes productos con poco stock:</p>
        <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
            <thead>
                <tr style="background-color: #4472C4; color: white;">
                    <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
                    <th style="padding: 10px; border: 1px solid #ddd;">Cobertura</th>
                    <th style="padding: 10px; border: 1px solid #ddd;">Estado</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>"""

    print("Leyendo stock y necesidades de deposito...")

    VO_STOCK = pd.read_excel(r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables VO.xlsx", "Necesidad por familia", header=1)
    VO_STOCK[STOCK_COL] = pd.to_numeric(VO_STOCK[STOCK_COL], errors="coerce")
    low_stock_VO_df = VO_STOCK[VO_STOCK[STOCK_COL] < 0.50].dropna(subset=[STOCK_COL])
    _sync_stockouts(low_stock_VO_df, "VO")
    send_email(EMAIL_TEMPLATE.format(rows=build_rows_html(low_stock_VO_df)), "⚠️ Alerta productos con stock menor a 50% VO")

    ATLANTICO_STOCK = pd.read_excel(r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables ATLANTICO.xlsx", "Necesidad por familia")
    ATLANTICO_STOCK[STOCK_COL] = pd.to_numeric(ATLANTICO_STOCK[STOCK_COL], errors="coerce")
    low_stock_ATLANTICO_df = ATLANTICO_STOCK[ATLANTICO_STOCK[STOCK_COL] < 0.50].dropna(subset=[STOCK_COL])
    _sync_stockouts(low_stock_ATLANTICO_df, "Atlantico")
    send_email(EMAIL_TEMPLATE.format(rows=build_rows_html(low_stock_ATLANTICO_df)), "⚠️ Alerta productos con stock menor a 50% ATLANTICO")
    

def mostrar_kpis():
    df_db = pd.read_sql("SELECT * FROM stockouts", conn)
    if df_db.empty:
        print("No hay datos de stockouts registrados.")
        return

    df_db["date_of_stockout"] = pd.to_datetime(df_db["date_of_stockout"])
    df_db["date_resolved"] = pd.to_datetime(df_db["date_resolved"])
    today_ts = pd.Timestamp(today)

    open_so = df_db[df_db["status"] != "resolved"].copy()
    open_so["dias_abierto"] = (today_ts - open_so["date_of_stockout"]).dt.days

    resolved = df_db[df_db["status"] == "resolved"].copy()
    resolved["dias_para_resolver"] = (resolved["date_resolved"] - resolved["date_of_stockout"]).dt.days
    avg_days = resolved["dias_para_resolver"].mean() if not resolved.empty else None

    print("\n=== STOCKOUTS ABIERTOS ===")
    if open_so.empty:
        print("  No hay stockouts abiertos.")
    else:
        for _, row in open_so.sort_values("dias_abierto", ascending=False).iterrows():
            label = "QUEBRADO" if row["status"] == "broke" else "CRASH"
            cob = f"{row['cobertura']*100:.1f}%" if pd.notna(row.get("cobertura")) else "N/A"
            print(f"  [{label}] {row['product_name']} ({row['company']}) | Cobertura: {cob} | {row['dias_abierto']} dias sin resolver")

    print("\n=== KPI: PROMEDIO DIAS PARA RESOLVER ===")
    if avg_days is not None:
        print(f"  Promedio historico: {avg_days:.1f} dias ({len(resolved)} casos resueltos)")
    else:
        print("  Sin datos historicos de resolucion aun.")

    print(f"\n=== RESUMEN ===")
    broke_count = len(open_so[open_so["status"] == "broke"])
    crash_count = len(open_so[open_so["status"] == "crash"])
    print(f"  Abiertos: {len(open_so)} (crash: {crash_count}, quebrados: {broke_count}) | Resueltos: {len(resolved)}")
    if not open_so.empty:
        print(f"  Mas critico: {open_so.sort_values('dias_abierto', ascending=False).iloc[0]['product_name']} ({int(open_so['dias_abierto'].max())} dias)")


if (procedure == "1"):
    actualizar_stock()
elif (procedure == "2"):
    alertar_faltantes()
elif (procedure == "3"):
    mostrar_kpis()

