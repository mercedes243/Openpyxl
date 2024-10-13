import openpyxl
from openpyxl import Workbook

def cargar_o_crear_excel(nombre_archivo):
    try:
        libro = openpyxl.load_workbook(nombre_archivo)
    except FileNotFoundError:
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Gastos"
        hoja.append(["Fecha", "Descripción", "Monto"])
        libro.save(nombre_archivo)
    return libro

def ingresar_gastos():
    
    lista_gastos = []
    while True:
        fecha = input("Ingrese la fecha del gasto (YYYY-MM-DD): ")
        descripcion = input("Ingrese la descripción del gasto: ")
        monto = validar_monto()
        lista_gastos.append((fecha, descripcion, monto))

        agregar_mas = input("¿Desea ingresar un gasto adicional? (s/n): ").lower()
        if agregar_mas != 's':
            break
    return lista_gastos

def validar_monto():
    while True:
        try:
            monto = float(input("Ingrese el monto del gasto: "))
            return monto
        except ValueError:
            print("Error: el monto debe ser un número.")

def guardar_gastos_en_excel(gastos, nombre_archivo):
    libro = cargar_o_crear_excel(nombre_archivo)
    hoja = libro["Gastos"]
    
    for gasto in gastos:
        hoja.append(gasto)
    
    libro.save(nombre_archivo)

def generar_resumen(gastos):
    total_gastos = len(gastos)
    gasto_mas_caro = max(gastos, key=lambda x: x[2])
    gasto_mas_barato = min(gastos, key=lambda x: x[2])
    monto_total = sum(gasto[2] for gasto in gastos)
    
    print("\nResumen de Gastos:")
    print(f"Total de gastos: {total_gastos}")
    print(f"Gasto más caro: {gasto_mas_caro[1]} el {gasto_mas_caro[0]} por {gasto_mas_caro[2]}")
    print(f"Gasto más barato: {gasto_mas_barato[1]} el {gasto_mas_barato[0]} por {gasto_mas_barato[2]}")
    print(f"Monto total de gastos: {monto_total}")

def main():
    nombre_archivo = "informe_gastos.xlsx"
    print("Bienvenido al programa de gestión de gastos personales.")
    gastos = ingresar_gastos()
    guardar_gastos_en_excel(gastos, nombre_archivo)
    generar_resumen(gastos)
    print(f"\nLos datos de los gastos se han guardado en el archivo {nombre_archivo}.")

if __name__ == "__main__":
    main()

