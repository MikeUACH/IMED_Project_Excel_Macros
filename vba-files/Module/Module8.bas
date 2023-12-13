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
    AllocationTotalMERZ = hojaOrigen.Range("D87:O87")
    AllocationTotalMED = hojaOrigen.Range("D88:O88")
    AllocationTotalSNN = hojaOrigen.Range("D89:O89")
    AllocationTotalBSC = hojaOrigen.Range("D90:O90")
    AllocationTotalDRA = hojaOrigen.Range("D91:O91")
    AllocationTotalZIM = hojaOrigen.Range("D92:O92")
    AllocationTotalVAR = hojaOrigen.Range("D93:O93")
    AllocationTotalBD = hojaOrigen.Range("D94:O94")
    AllocationTotalCUT = hojaOrigen.Range("D95:O95")
    AllocationTotalDEX = hojaOrigen.Range("D96:O96")

    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    hojaDestino.Range("Z3:AK12").Value = TotalFlexlineBID2
    hojaDestino.Range("Z17:AK26").Value = AllocationUCBID2
    hojaDestino.Range("Z31:AK40").Value = AllocationTotalBID2
    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    hojaDestino.Range("D3:O12").Value = TotalFlexline
    hojaDestino.Range("D17:O26").Value = AllocationUC


    hojaDestino.Range("D31:O31").Value = AllocationTotalMERZ
    hojaDestino.Range("D32:O32").Value = AllocationTotalMED
    hojaDestino.Range("D33:O33").Value = AllocationTotalSNN
    hojaDestino.Range("D34:O34").Value = AllocationTotalBSC
    hojaDestino.Range("D35:O35").Value = AllocationTotalDRA
    hojaDestino.Range("D36:O36").Value = AllocationTotalZIM
    hojaDestino.Range("D37:O37").Value = AllocationTotalVAR
    hojaDestino.Range("D38:O38").Value = AllocationTotalBD
    hojaDestino.Range("D39:O39").Value = AllocationTotalCUT
End Sub


























