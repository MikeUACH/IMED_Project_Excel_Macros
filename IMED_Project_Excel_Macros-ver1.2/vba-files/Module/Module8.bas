Attribute VB_Name = "Module8"
Dim ArchivoDestinoPath As String
Sub ObtenerYColocarTabsUnabFlex(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim archivoOrigen As Workbook
    
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    
    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("AllocationTotal")
    Set hojaDestino = ArchivoDestino.Sheets("Sheet1")
    
    ' Obtiene los valores de la tabla antes de ser actualizada
    Dim TotalFlexlineBID2 As Variant
    Dim AllocationUCBID2 As Variant
    Dim AllocationTotalBID2 As Variant
    
    TotalFlexlineBID2 = hojaDestino.Range("D3:O12")
    AllocationUCBID2 = hojaDestino.Range("D17:O26")
    AllocationTotalBID2 = hojaDestino.Range("D31:O40")
    
    ' Obtiene los valores correspondientes al mes siguiente en cada hoja de origen
    Dim TotalFlexline As Variant
    Dim AllocationUC As Variant
    Dim AllocationTotal As Variant
    
    ' Buscar la coincidencia en hojaOrigen
    TotalFlexline = hojaOrigen.Range("D59:O69")

    ' Buscar la coincidencia en hojaOrigen
    AllocationUC = hojaOrigen.Range("D73:O83")
    
    ' Buscar la coincidencia en hojaOrigen
    AllocationTotal = hojaOrigen.Range("D86:O96")
    
    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    hojaDestino.Range("Z3:AK12").Value = TotalFlexlineBID2
    hojaDestino.Range("Z17:AK26").Value = AllocationUCBID2
    hojaDestino.Range("Z31:AK40").Value = AllocationTotalBID2
    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    hojaDestino.Range("D3:O12").Value = TotalFlexline
    hojaDestino.Range("D17:O26").Value = AllocationUC
    hojaDestino.Range("D31:O40").Value = AllocationTotal
    
End Sub


























