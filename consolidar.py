"""Consolida las ventas de todas las cajas en un solo Excel con resumen.

Uso: dejar los archivos de las cajas (.xlsx o .csv) en la carpeta `entrada` y ejecutar
`consolidar.bat`. El resultado queda en `salida/ventas_consolidadas.xlsx`.

Qué hace:
  - Reconoce cada archivo aunque las columnas tengan otros nombres o haya títulos arriba.
  - Unifica fechas (texto dd/mm/aaaa, aaaa-mm-dd o fecha de Excel) y montos ("15.000" o 15000).
  - Lleva cada producto a su nombre de `catalogo.csv`, aunque esté escrito distinto.
  - Elimina renglones duplicados y filas vacías o de totales.
  - Recalcula cada total como cantidad x precio y deja en la hoja "Revisar" los que no cerraban,
    los duplicados y los productos que no están en el catálogo.
"""

import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
import xlsxwriter

CARPETA = Path(__file__).resolve().parent
ENTRADA, SALIDA = CARPETA / "entrada", CARPETA / "salida"

SINONIMOS = {
    "ticket": {"ticket", "nro ticket", "comprobante", "nro comprobante"},
    "fecha": {"fecha", "fecha venta", "dia"},
    "producto": {"producto", "articulo", "descripcion"},
    "cantidad": {"cantidad", "cant", "unidades"},
    "precio": {"precio unitario", "precio", "p unit"},
    "total": {"total", "importe", "subtotal"},
    "vendedor": {"vendedor", "cajero"},
}


def sin_acentos(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


def nombre_columna(texto) -> str | None:
    limpio = re.sub(r"[^a-z0-9]+", " ", sin_acentos(texto)).strip()
    return next((campo for campo, nombres in SINONIMOS.items() if limpio in nombres), None)


def clave_producto(nombre: str) -> str:
    """Clave para comparar nombres de producto escritos de distintas formas."""
    s = sin_acentos(nombre)
    s = re.sub(r"(\d)\s*(lts?|litros?)\b", r"\1l", s)
    s = re.sub(r"(\d)\s*(grs?|gramos?)\b", r"\1g", s)
    s = re.sub(r"\bkilos?\b", "kg", s)
    s = re.sub(r"\bx\s*(\d+)", r"\1", s)
    palabras = re.sub(r"[^a-z0-9]+", " ", s).split()
    vacias = {"x", "de", "en", "unidad"}
    palabras = [p[:-1] if len(p) > 3 and p.endswith("s") else p for p in palabras if p not in vacias]
    return "".join(palabras)


def leer_catalogo() -> dict[str, tuple[str, str]]:
    catalogo = pd.read_csv(CARPETA / "catalogo.csv", sep=";", dtype=str).fillna("")
    claves = {}
    for _, fila in catalogo.iterrows():
        for nombre in [fila["producto"], *filter(None, fila["otros_nombres"].split("|"))]:
            claves[clave_producto(nombre)] = (fila["producto"], fila["categoria"])
    return claves


def leer_archivo(ruta: Path) -> pd.DataFrame:
    if ruta.suffix.lower() == ".csv":
        crudo = pd.read_csv(ruta, sep=None, engine="python", header=None, dtype=object, encoding="utf-8-sig")
    else:
        crudo = pd.read_excel(ruta, header=None, dtype=object)
    # El encabezado es la primera fila donde se reconocen al menos 4 columnas
    for i, fila in crudo.iterrows():
        campos = [nombre_columna(v) if pd.notna(v) else None for v in fila]
        if sum(c is not None for c in campos) >= 4:
            datos = crudo.iloc[i + 1:].copy()
            datos.columns = [c or f"_{n}" for n, c in enumerate(campos)]
            datos = datos[[c for c in datos.columns if not c.startswith("_")]]
            datos["caja"] = ruta.stem.replace("caja", "Caja ")
            return datos.dropna(how="all", subset=["ticket", "producto", "cantidad"])
    raise ValueError(f"No se reconocieron las columnas de {ruta.name}")


def a_numero(valor) -> float:
    if isinstance(valor, (int, float)):
        return float(valor)
    return float(str(valor).replace(".", "").replace(",", ".").strip())


def a_fecha(valor) -> pd.Timestamp:
    if isinstance(valor, pd.Timestamp) or hasattr(valor, "year"):
        return pd.Timestamp(valor)
    texto = str(valor).strip()
    formato = "%d/%m/%Y" if "/" in texto else "%Y-%m-%d"
    return pd.to_datetime(texto[:10], format=formato)


def consolidar() -> tuple[pd.DataFrame, pd.DataFrame]:
    catalogo = leer_catalogo()
    archivos = sorted(p for p in ENTRADA.iterdir() if p.suffix.lower() in {".xlsx", ".csv"})
    if not archivos:
        raise SystemExit(f"No hay archivos .xlsx ni .csv en {ENTRADA}")
    ventas = pd.concat([leer_archivo(p) for p in archivos], ignore_index=True)

    # Filas de totales: sin ticket y con "total" en la descripción
    ventas = ventas[ventas["ticket"].notna()].copy()
    revisar = []

    ventas["fecha"] = ventas["fecha"].map(a_fecha)
    ventas["ticket"] = ventas["ticket"].map(lambda v: int(a_numero(v)))
    for col in ["cantidad", "precio", "total"]:
        ventas[col] = ventas[col].map(a_numero)
    ventas["producto_original"] = ventas["producto"].astype(str).str.strip()

    duplicados = ventas.duplicated(subset=["caja", "ticket", "producto_original", "cantidad", "precio", "total"])
    for _, f in ventas[duplicados].iterrows():
        revisar.append(["Duplicado eliminado", f["caja"], f["ticket"], f["fecha"], f["producto_original"],
                        "El mismo renglón estaba cargado dos veces."])
    ventas = ventas[~duplicados].copy()

    calculado = ventas["cantidad"] * ventas["precio"]
    for _, f in ventas[ventas["total"] != calculado].iterrows():
        revisar.append(["Total no cierra", f["caja"], f["ticket"], f["fecha"], f["producto_original"],
                        f"Cargado: {f['total']:,.0f} · Cantidad x precio: {f['cantidad'] * f['precio']:,.0f}"
                        .replace(",", ".")])
    ventas["total"] = calculado

    reconocido = ventas["producto_original"].map(lambda n: catalogo.get(clave_producto(n)))
    for _, f in ventas[reconocido.isna()].iterrows():
        revisar.append(["Producto no reconocido", f["caja"], f["ticket"], f["fecha"], f["producto_original"],
                        "Agregarlo a catalogo.csv (o como otro nombre de un producto existente)."])
    ventas["producto"] = reconocido.map(lambda r: r[0] if r else None).fillna(ventas["producto_original"])
    ventas["categoria"] = reconocido.map(lambda r: r[1] if r else "Sin categoría")
    if "vendedor" not in ventas:
        ventas["vendedor"] = None
    ventas["vendedor"] = ventas["vendedor"].fillna("—")

    ventas = ventas.sort_values(["fecha", "caja", "ticket"]).reset_index(drop=True)
    revisar = pd.DataFrame(revisar, columns=["Tipo", "Caja", "Ticket", "Fecha", "Producto", "Detalle"])
    return ventas, revisar


def py(valor):
    """Valor de pandas/numpy a tipo de Python, para escribirlo en Excel."""
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    return valor.item() if hasattr(valor, "item") else valor


def escribir(ventas: pd.DataFrame, revisar: pd.DataFrame, ruta: Path) -> None:
    with xlsxwriter.Workbook(str(ruta)) as libro:
        enc = libro.add_format({"bold": True, "bg_color": "#14271E", "font_color": "#FFFFFF"})
        gs = libro.add_format({"num_format": '"Gs." #,##0'})
        num = libro.add_format({"num_format": "#,##0"})
        fecha = libro.add_format({"num_format": "dd/mm/yyyy"})
        titulo = libro.add_format({"bold": True, "font_size": 16, "font_color": "#14271E"})
        sub = libro.add_format({"bold": True, "font_size": 12, "font_color": "#14271E"})
        kpi = libro.add_format({"bold": True, "font_size": 15, "num_format": '"Gs." #,##0', "font_color": "#14271E"})
        kpi_num = libro.add_format({"bold": True, "font_size": 15, "num_format": "#,##0", "font_color": "#14271E"})
        gris = libro.add_format({"font_color": "#52655A"})

        resumen = libro.add_worksheet("Resumen")

        # Hoja Ventas: tabla de Excel con todas las ventas
        hoja = libro.add_worksheet("Ventas")
        cols = [("Fecha", "fecha", fecha, 12), ("Caja", "caja", None, 9), ("Ticket", "ticket", None, 9),
                ("Producto", "producto", None, 26), ("Categoría", "categoria", None, 13),
                ("Cantidad", "cantidad", num, 10), ("Precio", "precio", gs, 13), ("Total", "total", gs, 14),
                ("Vendedor", "vendedor", None, 11)]
        datos = [[py(r[c]) for _, c, _, _ in cols] for _, r in ventas.iterrows()]
        hoja.add_table(0, 0, len(datos), len(cols) - 1, {
            "data": datos, "name": "Ventas", "style": "Table Style Light 15",
            "columns": [{"header": h, "header_format": enc, "format": f} for h, _, f, _ in cols],
        })
        for i, (_, _, _, ancho) in enumerate(cols):
            hoja.set_column(i, i, ancho)
        hoja.freeze_panes(1, 0)

        # Hoja Resumen: todo con fórmulas sobre la tabla Ventas
        resumen.set_column("A:A", 24)
        resumen.set_column("B:B", 17)
        resumen.set_column("C:C", 3)
        resumen.set_column("D:D", 24)
        resumen.set_column("E:E", 17)
        resumen.write("A1", "Ventas consolidadas — octubre 2026", titulo)
        resumen.write("A2", "Minimercado de ejemplo (datos ficticios) · 3 cajas", gris)

        tickets = ventas.groupby(["caja", "ticket"]).ngroups
        resumen.write("A4", "Total vendido", gris)
        resumen.write_formula("A5", "=SUM(Ventas[Total])", kpi)
        resumen.write("B4", "Tickets", gris)
        resumen.write("B5", tickets, kpi_num)
        resumen.write("D4", "Ticket promedio", gris)
        resumen.write_formula("D5", "=A5/B5", kpi)
        resumen.write("E4", "Unidades", gris)
        resumen.write_formula("E5", "=SUM(Ventas[Cantidad])", kpi_num)

        fila = 7
        resumen.write(fila, 0, "Por caja", sub)
        cajas = sorted(ventas["caja"].unique())
        for i, caja in enumerate(cajas, start=fila + 1):
            resumen.write(i, 0, caja)
            resumen.write_formula(i, 1, f"=SUMIFS(Ventas[Total],Ventas[Caja],A{i + 1})", gs)

        resumen.write(fila, 3, "Por categoría", sub)
        categorias = ventas.groupby("categoria")["total"].sum().sort_values(ascending=False).index
        for i, cat in enumerate(categorias, start=fila + 1):
            resumen.write(i, 3, cat)
            resumen.write_formula(i, 4, f"=SUMIFS(Ventas[Total],Ventas[Categoría],D{i + 1})", gs)
        fin_cat = fila + len(categorias)

        fila = max(fila + len(cajas), fin_cat) + 2
        resumen.write(fila, 0, "10 productos más vendidos", sub)
        top = ventas.groupby("producto")["total"].sum().sort_values(ascending=False).head(10).index
        inicio_top = fila + 1
        for i, prod in enumerate(top, start=inicio_top):
            resumen.write(i, 0, prod)
            resumen.write_formula(i, 1, f"=SUMIFS(Ventas[Total],Ventas[Producto],A{i + 1})", gs)

        fila = inicio_top + len(top) + 1
        resumen.write(fila, 0, "Por día", sub)
        dias = sorted(ventas["fecha"].unique())
        inicio_dias = fila + 1
        for i, dia in enumerate(dias, start=inicio_dias):
            resumen.write_datetime(i, 0, pd.Timestamp(dia).to_pydatetime(), fecha)
            resumen.write_formula(i, 1, f"=SUMIFS(Ventas[Total],Ventas[Fecha],A{i + 1})", gs)
        fin_dias = inicio_dias + len(dias) - 1

        g_dias = libro.add_chart({"type": "column"})
        g_dias.add_series({"name": "Ventas por día", "categories": ["Resumen", inicio_dias, 0, fin_dias, 0],
                           "values": ["Resumen", inicio_dias, 1, fin_dias, 1], "fill": {"color": "#14271E"},
                           "gap": 40})
        g_dias.set_title({"name": "Ventas por día"})
        g_dias.set_legend({"none": True})
        g_dias.set_x_axis({"num_format": "dd", "date_axis": True})
        g_dias.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#E3E8E5"}}})
        g_dias.set_size({"width": 720, "height": 300})
        resumen.insert_chart("G4", g_dias)

        g_cat = libro.add_chart({"type": "bar"})
        g_cat.add_series({"name": "Por categoría", "categories": ["Resumen", 8, 3, fin_cat, 3],
                          "values": ["Resumen", 8, 4, fin_cat, 4], "fill": {"color": "#3E9A66"},
                          "data_labels": {"value": True, "num_format": "#,##0"}})
        g_cat.set_title({"name": "Ventas por categoría"})
        g_cat.set_legend({"none": True})
        g_cat.set_y_axis({"reverse": True})
        g_cat.set_x_axis({"num_format": "#,##0", "visible": False, "major_gridlines": {"visible": False}})
        g_cat.set_size({"width": 720, "height": 280})
        resumen.insert_chart("G20", g_cat)
        resumen.activate()

        # Hoja Revisar: lo que necesita una mirada humana
        rev = libro.add_worksheet("Revisar")
        if len(revisar):
            datos_rev = [[py(r[c]) for c in revisar.columns] for _, r in revisar.iterrows()]
            rev.add_table(0, 0, len(datos_rev), len(revisar.columns) - 1, {
                "data": datos_rev, "name": "Revisar", "style": "Table Style Light 10",
                "columns": [{"header": c, "header_format": enc, "format": fecha if c == "Fecha" else None}
                            for c in revisar.columns],
            })
        else:
            rev.write("A1", "No hay nada para revisar.")
        for col, ancho in enumerate([22, 9, 9, 12, 26, 60]):
            rev.set_column(col, col, ancho)

        # Hoja de instrucciones
        guia = libro.add_worksheet("Cómo se usa")
        guia.set_column("A:A", 100)
        pasos = [
            ("Cómo se usa", titulo),
            ("1. Copiar los archivos de las cajas (.xlsx o .csv) en la carpeta «entrada».", None),
            ("2. Hacer doble clic en «consolidar.bat».", None),
            ("3. Abrir «salida/ventas_consolidadas.xlsx».", None),
            ("", None),
            ("La hoja «Revisar» muestra los duplicados eliminados, los totales que no cerraban", None),
            ("(se recalculan como cantidad x precio) y los productos que no están en «catalogo.csv».", None),
            ("Para un producto nuevo, agregar una línea en «catalogo.csv» y volver a ejecutar.", None),
        ]
        for i, (texto, fmt) in enumerate(pasos):
            guia.write(i, 0, texto, fmt)


def main() -> int:
    SALIDA.mkdir(exist_ok=True)
    ventas, revisar = consolidar()
    ruta = SALIDA / "ventas_consolidadas.xlsx"
    escribir(ventas, revisar, ruta)
    print(f"{len(ventas)} renglones de {ventas['caja'].nunique()} cajas -> {ruta}")
    print("Para revisar:", revisar["Tipo"].value_counts().to_dict() if len(revisar) else "nada")
    return 0


if __name__ == "__main__":
    sys.exit(main())
