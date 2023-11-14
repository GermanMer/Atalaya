import re

def extraer_valores(lineas):
    valores = []
    patron = re.compile(r'\b\d{1,3}(?:\.\d{3})*(?:,\d+)\b')

    for linea in lineas:
        resultado = patron.search(linea)
        if resultado:
            valor = resultado.group()
            valores.append(float(valor.replace('.', '').replace(',', '.')))

    return valores

# Abre el archivo en modo de lectura
with open('Extracto BNA.txt', 'r') as archivo:
    # Define los patrones de búsqueda con expresiones regulares
    patron_comisiones = re.compile(r'\d{2}/\d{2}/\d{2}.*?COM')
    patron_iva = re.compile(r'\d{2}/\d{2}/\d{2}(?!.*RETEN\.)\s.*?I\.V\.A\.')
    patron_reten_iva = re.compile(r'\d{2}/\d{2}/\d{2}.*?RETEN\. I\.V\.A\.')

    # Inicializa listas para almacenar las líneas encontradas
    lineas_comisiones = []
    lineas_iva = []
    lineas_reten_iva = []

    # Itera sobre cada línea en el archivo
    for linea in archivo:
        # Busca el patrón de COM en la línea
        resultado_comisiones = patron_comisiones.search(linea)
        # Busca el patrón de I.V.A. en la línea
        resultado_iva = patron_iva.search(linea)
        # Busca el patrón de RETEN. I.V.A. en la línea
        resultado_reten_iva = patron_reten_iva.search(linea)

        # Clasifica la línea en el grupo correspondiente
        if resultado_comisiones:
            lineas_comisiones.append(linea.strip())
        elif resultado_iva:
            lineas_iva.append(linea.strip())
        elif resultado_reten_iva:
            lineas_reten_iva.append(linea.strip())

# Extrae valores de la lista de COM
valores_comisiones = extraer_valores(lineas_comisiones)
# Imprime las líneas encontradas para COM
print("COMISIONES:")
for linea in lineas_comisiones:
    print(linea)
# Suma los valores de COM
suma_comisiones = sum(valores_comisiones)
print("Total de COMISIONES:", suma_comisiones)

# Extrae valores de la lista de I.V.A.
valores_iva = extraer_valores(lineas_iva)
# Imprime valores de COM
# Imprime las líneas encontradas para I.V.A.
print("\nI.V.A.:")
for linea in lineas_iva:
    print(linea)
# Suma los valores de I.V.A.
suma_iva = sum(valores_iva)
print("Total de I.V.A.:", suma_iva)

# Extrae valores de la lista de RETEN. I.V.A.
valores_reten_iva = extraer_valores(lineas_reten_iva)
# Imprime las líneas encontradas para RETEN. I.V.A.
print("\nRETENCIONES I.V.A.:")
for linea in lineas_reten_iva:
    print(linea)
# Suma los valores de RETEN. I.V.A.
suma_reten_iva = sum(valores_reten_iva)
print("Total de RETENCIONES I.V.A.:", suma_reten_iva)

# Agregar una espera para que la ventana no se cierre automáticamente en Windows
print("\n")
input("Presiona Enter para salir...")
