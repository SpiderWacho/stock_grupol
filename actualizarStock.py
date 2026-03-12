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

conn = sqlite3.connect(r"C:\Users\GastonVecchio\Documents\Code\inventory.db")
cursor = conn.cursor()

procedure = input("Que queres hacer?\n1-Actualizar stocks\n2-Alertar faltantes\n")


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

def alertar_faltantes():

    def send_email(html_body, subject):
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_RECEIVER
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Send()

    print("Leyendo stock y necesidades de deposito...")
    ## Actualizar archivos de necesidad

    
    VO_STOCK = pd.read_excel(r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables VO.xlsx","Necesidad por familia",header=1) 
    VO_STOCK["Cobertura"] = pd.to_numeric(VO_STOCK["Cobertura"], errors="coerce")
    low_stock_VO_df = VO_STOCK[VO_STOCK["Cobertura"] < 0.50]
    low_stock_VO_df = low_stock_VO_df.dropna(subset=["Cobertura"])

    rows_html = ""
    for _, row in low_stock_VO_df.iterrows():
        product = row[PRODUCT_COL]
        cobertura = f"{row[STOCK_COL] * 100:.2f}%"
        estado = "Pendiente" if pd.isna(row["Estado"]) else row["Estado"]

        rows_html += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">{product}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{cobertura}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{estado}</td>
            </tr>
        """

    html_body = f"""
    <p>Este es un mail automatico, comprar los siguientes productos con poco stock:</p>
    <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
        <thead>
            <tr style="background-color: #4472C4; color: white;">
                <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Cobertura</th>
                <th style="padding: 10px; border: 1px solid #ddd;">Estado</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>
    """

    send_email(html_body, "⚠️ Alerta productos con stock menor a 50% VO")

    # Build email ATLANTICO
    ATLANTICO_STOCK = pd.read_excel(r"C:\Users\GastonVecchio\Grupo L\Abastecimiento Online - Documentos (1)\10. Operador Descartables\COMPRAS DESCARTABLES MORENO Y CIUDADELA\Seguimiento stock descartables ATLANTICO.xlsx", "Necesidad por familia")
    ATLANTICO_STOCK["Cobertura"] = pd.to_numeric(ATLANTICO_STOCK["Cobertura"], errors="coerce")
    low_stock_ATLANTICO_df = ATLANTICO_STOCK[ATLANTICO_STOCK["Cobertura"] < 0.50]      
    low_stock_ATLANTICO_df = low_stock_ATLANTICO_df.dropna(subset=["Cobertura"])                                     

    rows_html = ""
    for _, row in low_stock_ATLANTICO_df.iterrows():
        product = row[PRODUCT_COL]
        cobertura = f"{row[STOCK_COL] * 100:.2f}%"
        estado = "Pendiente" if pd.isna(row["Estado"]) else row["Estado"]
        rows_html += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">{product}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{cobertura}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">{estado}</td>
            </tr>
        """

    html_body = f"""
        <p>Este es un mail automatico, comprar los siguientes productos con poco stock:</p>
        <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
            <thead>
                <tr style="background-color: #4472C4; color: white;">
                    <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
                    <th style="padding: 10px; border: 1px solid #ddd;">Cobertura</th>
                    <th style="padding: 10px; border: 1px solid #ddd;">Estado</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        """

    # Send email ATLATNICO
    send_email(html_body, "⚠️ Alerta productos con stock menor a 50% ATLANTICO")

    def write_to_db():
        cursor.execute("SELECT product_name, date_of_stockout FROM stockouts WHERE product_name=?",(product,))
        rows = cursor.fetchall()


        for product_name, date_of_stockout in rows:
            stockout_date = date.fromisoformat(date_of_stockout)
            days_ago = (today - stockout_date).days
            print(f"{product_name} ran out of stock {days_ago} days ago")
            if (days_ago < 1):
                cursor.execute("""
                    INSERT INTO stockouts (product_name, date_of_stockout, note)
                    VALUES (?, ?, ?)
                    """, (product, today, estado)          )
        
        conn.commit()


if (procedure == "1"):
    actualizar_stock()
elif (procedure == "2"):
    alertar_faltantes()

