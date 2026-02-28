from openpyxl import Workbook, load_workbook
import os

ARCHIVO = "inventario.xlsx"

PRECIOS = {
    "Manzana": 0.50,
    "Banana": 0.30,
    "Naranja": 0.40,
    "Pera": 0.60,
    "Mango": 1.00,
    "Uva": 0.80,
    "Sandia": 3.00,
    "Melon": 2.50,
    "Kiwi": 0.90,
    "Piña": 1.50,
    "Fresa": 0.70
}

CANTIDAD_MAXIMA = 50
CANTIDAD_MINIMA = 1


def crear_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    ws.append(["Fruta", "Cantidad", "Precio"])
    wb.save(ARCHIVO)


def abrir_excel():
    if not os.path.exists(ARCHIVO):
        crear_excel()
    return load_workbook(ARCHIVO)


def buscar_fruta(ws, fruta):
    for fila in ws.iter_rows(min_row=2, values_only=False):
        if fila[0].value == fruta:
            return fila
    return None


def agregar_fruta():
    wb = abrir_excel()
    ws = wb["Inventario"]

    fruta = input("Nombre de la fruta: ").capitalize()

    if fruta not in PRECIOS:
        print(" Fruta no permitida.")
        return

    try:
        cantidad = int(input("Cantidad a agregar: "))
        if cantidad <= 0:
            raise ValueError
    except ValueError:
        print(" Cantidad inválida.")
        return

    fila = buscar_fruta(ws, fruta)

    if fila:
        cantidad_actual = fila[1].value
        nueva = cantidad_actual + cantidad

        if nueva > CANTIDAD_MAXIMA:
            print(" Supera el máximo permitido.")
            return

        fila[1].value = nueva
    else:
        if cantidad < CANTIDAD_MINIMA:
            print(" Debe iniciar con mínimo 25.")
            return
        if cantidad > CANTIDAD_MAXIMA:
            print(" Supera el máximo permitido.")
            return

        ws.append([fruta, cantidad, PRECIOS[fruta]])

    wb.save(ARCHIVO)
    print(" Inventario actualizado en Excel.")


def vender_fruta():
    wb = abrir_excel()
    ws = wb["Inventario"]

    fruta = input("Fruta vendida: ").capitalize()
    fila = buscar_fruta(ws, fruta)

    if not fila:
        print(" Fruta no encontrada.")
        return

    try:
        cantidad = int(input("Cantidad vendida: "))
        if cantidad <= 0:
            raise ValueError
    except ValueError:
        print(" Cantidad inválida.")
        return

    cantidad_actual = fila[1].value
    nueva = cantidad_actual - cantidad

    if nueva < CANTIDAD_MINIMA:
        print(" No puede bajar del mínimo 25.")
        return

    fila[1].value = nueva
    wb.save(ARCHIVO)
    print(" Venta registrada en Excel.")


def mostrar_inventario():
    wb = abrir_excel()
    ws = wb["Inventario"]

    print("\n INVENTARIO ACTUAL:")
    for fila in ws.iter_rows(values_only=True):
        print(fila)


def menu():
    while True:
        print("\n--- SISTEMA INVENTARIO (EXCEL EN VIVO) ---")
        print("1. Agregar fruta")
        print("2. Vender fruta")
        print("3. Mostrar inventario")
        print("4. Salir")

        opcion = input("Opción: ")

        match opcion:
            case "1":
                agregar_fruta()
            case "2":
                vender_fruta()
            case "3":
                mostrar_inventario()
            case "4":
                break
            case _:
                print("Opción inválida")


if __name__ == "__main__":
    menu()