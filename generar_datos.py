"""Genera los archivos de ejemplo de la muestra: las ventas de octubre de 2026 de las 3 cajas de un
minimercado ficticio. Cada caja exporta distinto y con los errores típicos de las planillas reales:

  - caja1.xlsx: fechas como texto dd/mm/aaaa, nombres de producto escritos de varias formas,
    filas duplicadas, filas vacías y algunos totales mal tipeados.
  - caja2.csv: separado por punto y coma, otros nombres de columna, montos con punto de miles.
  - caja3.xlsx: dos filas de título arriba del encabezado y una fila de totales al final.

Los datos son inventados (semilla fija, para que el resultado sea siempre el mismo).
"""

import csv
import datetime as dt
import random
from pathlib import Path

import openpyxl

ENTRADA = Path(__file__).resolve().parent / "entrada"
rnd = random.Random(2026)

# (nombre canónico, categoría, precio en guaraníes, formas en que lo escriben las cajas)
PRODUCTOS = [
    ("Gaseosa cola 2 L", "Bebidas", 13000, ["Gaseosa cola 2L", "GASEOSA COLA 2 L", "gaseosa cola 2lt"]),
    ("Agua mineral 2 L", "Bebidas", 6000, ["Agua mineral 2L", "AGUA MINERAL 2 L"]),
    ("Cerveza 1 L", "Bebidas", 10500, ["Cerveza 1L", "cerveza 1 lt", "CERVEZA 1 L"]),
    ("Jugo en polvo", "Bebidas", 2500, ["Jugo en polvo", "jugo polvo"]),
    ("Leche entera 1 L", "Lácteos", 7500, ["Leche entera 1L", "LECHE ENTERA 1 L", "leche 1 lt"]),
    ("Yogur 1 L", "Lácteos", 11000, ["Yogur 1L", "YOGURT 1 L"]),
    ("Queso Paraguay kg", "Lácteos", 42000, ["Queso Paraguay kg", "queso py kg"]),
    ("Yerba mate 500 g", "Almacén", 15500, ["Yerba mate 500g", "YERBA 500 G", "yerba mate 500 gr"]),
    ("Arroz 1 kg", "Almacén", 8000, ["Arroz 1kg", "ARROZ 1 KG", "arroz 1 kilo"]),
    ("Fideos 500 g", "Almacén", 5000, ["Fideos 500g", "FIDEO 500 G"]),
    ("Aceite de girasol 900 ml", "Almacén", 16500, ["Aceite girasol 900ml", "ACEITE 900 ML"]),
    ("Azúcar 1 kg", "Almacén", 7000, ["Azucar 1kg", "AZÚCAR 1 KG", "azucar 1 kilo"]),
    ("Harina 1 kg", "Almacén", 6000, ["Harina 1kg", "HARINA 1 KG"]),
    ("Sal fina 500 g", "Almacén", 3000, ["Sal fina 500g", "SAL 500 G"]),
    ("Galletitas dulces", "Almacén", 5500, ["Galletitas dulces", "galletita dulce"]),
    ("Huevos (docena)", "Frescos", 16000, ["Huevos docena", "HUEVO X 12"]),
    ("Pan (kg)", "Frescos", 12000, ["Pan kg", "PAN X KG"]),
    ("Chipa (unidad)", "Frescos", 3000, ["Chipa", "CHIPA UNIDAD"]),
    ("Tomate (kg)", "Frescos", 9000, ["Tomate kg", "TOMATE X KG"]),
    ("Banana (kg)", "Frescos", 6500, ["Banana kg", "banana x kilo"]),
    ("Detergente 500 ml", "Limpieza", 8500, ["Detergente 500ml", "DETERGENTE 500 ML"]),
    ("Jabón en polvo 800 g", "Limpieza", 19000, ["Jabon en polvo 800g", "JABÓN POLVO 800 G"]),
    ("Lavandina 1 L", "Limpieza", 6500, ["Lavandina 1L", "LAVANDINA 1 LT"]),
    ("Papel higiénico x4", "Limpieza", 11500, ["Papel higienico x4", "PAPEL HIG. X4"]),
]
VENDEDORES = ["Ana", "Luis", "Rosa", "Carlos"]


def tickets(caja: int):
    """Tickets de un mes para una caja: (n.º de ticket, fecha, renglones)."""
    numero = 1000 * caja
    for dia in range(1, 32):
        fecha = dt.date(2026, 10, dia)
        cantidad = rnd.randint(18, 32) + (10 if fecha.weekday() >= 5 else 0)
        for _ in range(cantidad):
            numero += 1
            renglones = []
            for producto in rnd.sample(PRODUCTOS, rnd.randint(1, 4)):
                nombre, _, precio, formas = producto
                renglones.append((rnd.choice(formas), rnd.randint(1, 3), precio))
            yield numero, fecha, renglones


def caja1():
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Ticket", "Fecha", "Producto", "Cantidad", "Precio unitario", "Total"])
    for numero, fecha, renglones in tickets(1):
        for nombre, cant, precio in renglones:
            total = cant * precio
            if rnd.random() < 0.006:  # total mal tipeado
                total += rnd.choice([1000, -1000, 10000])
            fila = [numero, fecha.strftime("%d/%m/%Y"), nombre, cant, precio, total]
            hoja.append(fila)
            if rnd.random() < 0.01:  # el mismo renglón cargado dos veces
                hoja.append(fila)
        if rnd.random() < 0.02:
            hoja.append([None] * 6)
    libro.save(ENTRADA / "caja1.xlsx")


def caja2():
    with (ENTRADA / "caja2.csv").open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["nro_ticket", "fecha_venta", "articulo", "cant", "precio", "importe", "vendedor"])
        for numero, fecha, renglones in tickets(2):
            vendedor = rnd.choice(VENDEDORES)
            for nombre, cant, precio in renglones:
                w.writerow([numero, fecha.isoformat(), nombre, cant, f"{precio:,}".replace(",", "."),
                            f"{cant * precio:,}".replace(",", "."), vendedor])


def caja3():
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["REPORTE DE VENTAS - CAJA 3"])
    hoja.append(["Período: octubre 2026"])
    hoja.append(["Comprobante", "Día", "Descripción", "Unidades", "P. Unit.", "Subtotal", "Cajero"])
    suma = 0
    for numero, fecha, renglones in tickets(3):
        cajero = rnd.choice(VENDEDORES)
        for nombre, cant, precio in renglones:
            hoja.append([numero, fecha, nombre, cant, precio, cant * precio, cajero])
            suma += cant * precio
    hoja.append([None, None, "TOTAL GENERAL", None, None, suma, None])
    libro.save(ENTRADA / "caja3.xlsx")


if __name__ == "__main__":
    ENTRADA.mkdir(exist_ok=True)
    caja1()
    caja2()
    caja3()
    print("Archivos de ejemplo creados en", ENTRADA)
