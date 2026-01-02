import pandas as pd

#Esta es la ruta de archivos
INPUT_FILE = "data/ventas_ejemplo.xlsx"
OUTPUT_file = "output/reporte_generado.xlsx"

def generar_reporte():
  #Esta parte lee el archivo de excel
  df = pd.read_excel(INPUT_FILE)

#Esta parte elimina las filas incompletas
df = df.dropna()

#Aca se asegura los tipos numericos
df["Cantidad"] = pd.to_numeric(df["Cantidad"])
df["Precio"] = pd.to_numeric(df["Precio"])

#Aca se calcula tl total por fila
df["Total"] = df["Cantidad"]) * df["Precio"])

#Aca se crea el resumen por producto
resumen = (
  df.groupby("Producto")["Total"].sum().reset_index()
)

# Guardar el reporte en un nuevo Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalle limpio", index=False)
        resumen.to_excel(writer, sheet_name="Resumen por producto", index=False)

    print("Reporte generado correctamente.")


if __name__ == "__main__":
    generar_reporte()
