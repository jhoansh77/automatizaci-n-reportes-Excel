import pandas as pd

# Rutas de archivos
INPUT_FILE = "data/ventas_ejemplo.xlsx"
OUTPUT_FILE = "output/reporte_generado.xlsx"


def generar_reporte():
    # Leer el archivo Excel
    df = pd.read_excel(INPUT_FILE)

    # Eliminar filas incompletas
    df = df.dropna()

    # Asegurar tipos numéricos
    df["Cantidad"] = pd.to_numeric(df["Cantidad"])
    df["Precio"] = pd.to_numeric(df["Precio"])

    # Calcular total por fila
    df["Total"] = df["Cantidad"] * df["Precio"]

    # Crear resumen por producto
    resumen = (
        df.groupby("Producto")["Total"]
        .sum()
        .reset_index()
    )

    # Guardar el reporte en un nuevo Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalle limpio", index=False)
        resumen.to_excel(writer, sheet_name="Resumen por producto", index=False)

    print("Reporte generado correctamente.")


if __name__ == "__main__":
    generar_reporte()
