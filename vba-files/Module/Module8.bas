Attribute VB_Name = "Module8"
Dim archivoOrigenPath As String ' Variable global para almacenar la ruta del archivo origen
Dim ArchivoDestinoPath As String
Sub ObtenerYColocarTabsUnabFlex()
    Dim archivoOrigen As Workbook
    
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    ' Verifica si ya se ha seleccionado un archivo de origen previamente
    If archivoOrigenPath = "" Then
        ' Abre el cuadro de diálogo de selección de archivo con un título personalizado para el archivo de origen
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de origen(Flexline-Unabsorbed Calculation)")
        ' Verifica si se seleccionó un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If ArchivoDestinoPath = "" Then
        ' Abre el cuadro de diálogo de selección de archivo
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de destino(Variance BID2 Vs BID 3)")
        ' Verifica si se seleccionó un archivo
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If
    
    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    ' Define la hoja de cï¿½lculo en el archivo de origen
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
    
    ' Coloca los valores obtenidos en celdas especï¿½ficas de tu hoja de cï¿½lculo principal
    hojaDestino.Range("Z3:AK12").value = TotalFlexlineBID2
    hojaDestino.Range("Z17:AK26").value = AllocationUCBID2
    hojaDestino.Range("Z31:AK40").value = AllocationTotalBID2
    ' Coloca los valores obtenidos en celdas especï¿½ficas de tu hoja de cï¿½lculo principal
    hojaDestino.Range("D3:O12").value = TotalFlexline
    hojaDestino.Range("D17:O26").value = AllocationUC
    hojaDestino.Range("D31:O40").value = AllocationTotal
    
    archivoOrigen.Close SaveChanges:=False
End Sub


























