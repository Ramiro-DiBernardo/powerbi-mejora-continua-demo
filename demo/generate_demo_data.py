"""
Generate a synthetic dataset for the "Mejora Continua 2026" Power BI report.

The workbook keeps the exact structure the Power Query expects
(sheet "BASE", title in row 1, headers in row 2), but every value is fake:
generic sites, units, workstations and owners. No company data is used.

Usage:
    python generate_demo_data.py            # writes mejora_continua_demo.xlsx
Then point the RutaExcel parameter in Power BI to the generated file.
"""
import random
from datetime import date, timedelta

from openpyxl import Workbook

random.seed(2026)

N_ROWS = 110
SITES = ["Planta Norte", "Planta Sur", "Centro Logístico"]
UNITS = {
    "Laboratorio": ["Recepción", "Reparación", "Prensado", "OQC", "Embalaje"],
    "Depósito": ["Picking", "Despacho", "Almacenamiento"],
    "Mantenimiento": ["Taller", "Pañol"],
    "Áreas comunes": ["Comedor", "Vestuarios", "Oficinas"],
    "Equipos médicos": ["Recepción", "Funcional", "OQC"],
}
PROBLEMS = [
    ("Layout del puesto de {p} ineficiente", "Rediseño del layout del puesto de {p}"),
    ("Falta señalización en {p}", "Señalización y demarcación de {p}"),
    ("Herramientas sin ubicación fija en {p}", "Implementación de tablero de sombras en {p}"),
    ("Tiempos de espera altos en {p}", "Balanceo de tareas en {p}"),
    ("Riesgo de daño de equipos en {p}", "Protecciones y bandejas para {p}"),
    ("Material almacenado a piso en {p}", "Estanterías y picking ordenado en {p}"),
    ("Registro manual en papel en {p}", "Digitalización del registro de {p}"),
    ("Falta de puntos de carga en {p}", "Estación de carga en {p}"),
    ("Iluminación insuficiente en {p}", "Mejora de iluminación en {p}"),
    ("Cables sueltos en {p}", "Ordenamiento de cableado en {p}"),
]
OWNERS = [f"Responsable {i}" for i in range(1, 9)]
STATUS_WEIGHTS = {"IMPLEMENTADA": 0.55, "EN PROCESO": 0.2, "NO INICIADO": 0.25}

HEADERS = ["N°", "FECHA RELEVO", "AÑO", "MES", "DIA", "DESCRIPCIÓN DE MEJORAS", "ACCIÓN",
           "IMPLEMENTADO", "SITIO", "UNIDAD", "PUESTO", "STATUS IDOC",
           "RESPONSABLE\nCARGA IDOC", "STATUS", "IMPACTO \nECONÓMICO"]
MONTHS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def main(path="mejora_continua_demo.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "BASE"
    ws.append(["FOLLOW UP MEJORAS (DEMO DATA)"])
    ws.append(HEADERS)

    start = date(2026, 1, 5)
    for n in range(1, N_ROWS + 1):
        unit = random.choice(list(UNITS))
        post = random.choice(UNITS[unit])
        problem, action = random.choice(PROBLEMS)
        found = start + timedelta(days=random.randint(0, 280))
        status = random.choices(list(STATUS_WEIGHTS), weights=STATUS_WEIGHTS.values())[0]
        done = found + timedelta(days=random.randint(5, 45)) if status == "IMPLEMENTADA" else None
        if done and done > date(2026, 10, 31):
            done, status = None, "EN PROCESO"
        ws.append([
            n, found, found.year, MONTHS[found.month - 1], found.day,
            problem.format(p=post if post.isupper() else post.lower()), action.format(p=post if post.isupper() else post.lower()),
            done, random.choice(SITES), unit, post,
            "IMPLEMENTADA" if status == "IMPLEMENTADA" else random.choice(["PENDIENTE", "CARGADA"]),
            random.choice(OWNERS), status,
            random.choices(["APLICA", "NO APLICA"], weights=[0.3, 0.7])[0],
        ])
    for row in ws.iter_rows(min_row=3, min_col=2, max_col=8):
        for cell in row:
            if isinstance(cell.value, date):
                cell.number_format = "dd/mm/yyyy"
    wb.save(path)
    print(f"Wrote {N_ROWS} fake rows to {path}")


if __name__ == "__main__":
    main()
