import pandas as pd
import os

ARCHIVO = "inventario.xlsx"

def cargar_inventario():
    if os.path.exists(ARCHIVO):
        return pd.read_excel(ARCHIVO)
    else:
        df = pd.DataFrame(columns=["Fruta", "Cantidad", "Precio"])
        df.to_excel(ARCHIVO, index=False)
        return df

def guardar_inventario(df):
    df.to_excel(ARCHIVO, index=False)

def agregar_fruta(df):
    fruta = input("Nombre de la fruta: ").capitalize()
    cantidad = int(input("Cantidad: "))
    precio = float(input("Precio por unidad: "))

    if fruta in df["Fruta"].values:
        df.loc[df["Fruta"] == fruta, "Cantidad"] += cantidad
    else:
        nueva = pd.DataFrame([[fruta, cantidad, precio]], columns=df.columns)
        df = pd.concat([df, nueva], ignore_index=True)

    guardar_inventario(df)
    print("Fruta agregada correctamente.")
    return df

def vender_fruta(df):
    fruta = input("Fruta vendida: ").capitalize()
    cantidad = int(input("Cantidad vendida: "))

    if fruta in df["Fruta"].values:
        df.loc[df["Fruta"] == fruta, "Cantidad"] -= cantidad
        guardar_inventario(df)
        print("Venta registrada.")
    else:
        print("La fruta no existe en el inventario.")

    return df

def mostrar_inventario(df):
    print("\nInventario actual:")
    print(df)

def menu():
    df = cargar_inventario()

    while True:
        print("\n--- SISTEMA DE INVENTARIO FRUTERÍA ---")
        print("1. Agregar fruta")
        print("2. Vender fruta")
        print("3. Mostrar inventario")
        print("4. Salir")

        opcion = input("Seleccione una opción: ")

        if opcion == "1":
            df = agregar_fruta(df)
        elif opcion == "2":
            df = vender_fruta(df)
        elif opcion == "3":
            mostrar_inventario(df)
        elif opcion == "4":
            break
        else:
            print("Opción inválida")

if __name__ == "__main__":
    menu()