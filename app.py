import streamlit as st
import pandas as pd
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import io
import datetime

XLS_PATH = "FAVT-V41.xlsx"


# =========================================================
# Excel loader (CRÍTICO: data_only=True)
# =========================================================
def load_excel():
    wb = openpyxl.load_workbook(XLS_PATH, data_only=True)

    # -------------------------
    # PREÇO BAIRROS
    # -------------------------
    ws = wb["PREÇO BAIRROS"]

    headers = [ws.cell(row=4, column=c).value for c in range(1, 8)]
    rows = []
    for r in range(5, 93):
        row = [ws.cell(row=r, column=c).value for c in range(1, 8)]
        if all(v is None for v in row):
            continue
        rows.append(row)

    preco_df = pd.DataFrame(rows, columns=headers)

    # -------------------------
    # FACTORES
    # -------------------------
    f1 = wb["Factor1_H"]
    f1_df = pd.DataFrame(
        [[f1.cell(r, c).value for c in range(1, 6)] for r in range(11, 30)]
    )

    f2 = wb["Factor2_Elev"]
    f2_df = pd.DataFrame(
        [[f2.cell(r, c).value for c in range(1, 5)] for r in range(11, 30)]
    )

    vistas = wb["N_Vistas"]
    vistas_df = pd.DataFrame(
        [[vistas.cell(r, c).value for c in range(2, 4)] for r in range(3, 10)]
    )

    infra = wb["Factor5_infra"]
    infra_df = pd.DataFrame(
        [[infra.cell(r, c)]()]()
