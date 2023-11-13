import openpyxl

# Ruta del archivo de Excel
archivo_excel = "Flexline-Unabsorbed Calculation (Forecast) FY24 Q124 BID1 No ERSA (Macros habilitadas).xlsm"

# Lista de valores que quieres buscar
valores_a_buscar = [
    "ARC43268-G-SVC",
    "ARC69033-E-SVC",
    "ARCR09010-K-SVC",
    "ARCR09930-M-SVC",
    "ARCR13208-AA-SVC",
    "ARCR14870-M-SVC",
    "ARCR16292-AG-SVC",
    "ARCR18793-T-SVC",
    "ARCR40420-01-K-SVC",
    "ARCR40420-03-02",
    "ARCR72290105-G",
    "ARCR72290144-E",
    "ARCR72290146-C",
    "BX92316384-10",
    "BX92316388-12",
    "BX92436973-11",
    "DR-MS26321",
    "DR-MS30364",
    "DR-MS30366",
    "DR-MS30372",
    "DR-MS30379",
    "DR-MS30680",
    "DR-MS30740",
    "DR-MS30742",
    "DR-MS31088",
    "DR-MS31254",
    "DR-MS31256",
    "DR-MS31311",
    "DR-MS32107",
    "DR-MS32109",
    "DR-MS32422",
    "DR-MS32776",
    "DR-MS32881",
    "DR-MS32981",
    "DR-MS33121",
    "DR-MS33991",
    "DR-MS34220",
    "DR-MS34381",
    "DR-MS34601",
    "DR-MS34621",
    "DR-MS34631",
    "DR-MS34681",
    "DR-MS34691",
    "DR-MS34711",
    "DR-MS34821",
    "DR-MS40251",
    "DR-MS40605",
    "DR-MU23388",
    "DR-MU23389",
    "DR-MU23390",
    "DR-MU23391",
    "DR-MU25466",
    "DR-MU26011",
    "DR-MU26023",
    "DR-MU26373",
    "DR-MU27191",
    "DR-MU27192",
    "DR-MS41201",
    "DR-MU60791",
    "DR-MS34861",
    "DR-MS40107",
    "DR-MS40108",
    "DR-MS40633",
    "DR-MS41171",
    "DR-MS41181",
    "DR-MS41221",
    "DR-MS41231",
    "DR-MS41241",
    "DR-MS41251",
    "FR-180010",
    "FR-180196",
    "FR-180232-01",
    "FR-180350",
    "FR-180409",
    "FR-180424",
    "FR-190019",
    "FR-190234-232",
    "FR-190246",
    "FR-190332",
    "FR-190500",
    "FR-190522",
    "FR-190710",
    "FR-190816",
    "FR-190821-D",
    "FR-190937",
    "FR-190951-106",
    "FR-191061-D",
    "FR-191117",
    "FR-191142",
    "FR-180198",
    "MDT-11800094",
    "MDT-1180152",
    "MDT-1180170",
    "U-1000081-RB",
    "U-1000085-RB-JA",
    "U-1000770-RD-JA",
    "U-1000873-RA",
    "U-2040-0208-R11",
    "U-2040-0219-31-JA",
    "U-2040-0219-51-JA",
    "U-2040-0220-R3",
    "U-2040-0266-R4",
    "U-2040-0275-R7R",
    "BCDTC10008452",
    "BCDTC10013054",
    "BCDTC10013065",
    "BCDTC10013069",
    "BCDTC10013114",
    "BCDTC10013117",
    "BCDTC10017344",
    "BCDTC10018763",
    "ZM400.053",
    "ZM400.054",
    "ZM400.055",
    "ZMATS3200",
    "ZMATS5000"
]

# Listas para almacenar valores encontrados y no encontrados
encontrados = []
no_encontrados = []

# Abrir el archivo de Excel
workbook = openpyxl.load_workbook(archivo_excel)
sheet = workbook["Forecast"]

# Iterar a través de la lista de valores a buscar
for valor in valores_a_buscar:
    encontrado = False
    for row in sheet.iter_rows(values_only=True):
        for cell_value in row:
            if cell_value == valor:
                encontrado = True
                break
    if encontrado:
        encontrados.append(valor)
    else:
        no_encontrados.append(valor)

# Cerrar el archivo de Excel
workbook.close()

# Mostrar los resultados
print("Valores encontrados:")
for valor in encontrados:
    print(valor)

print("Valores no encontrados:")
for valor in no_encontrados:
    print(valor)

print(f"Cantidad encontrados: {len(encontrados)}")
print(f"Cantidad no encontrados: {len(no_encontrados)}")