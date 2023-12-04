Attribute VB_Name = "Module2"
Dim archivoOrigenPath As String
Public ArchivoDestinoPathBU As String

Sub UpdNonMatMarginBU(ByVal ArchivoDestinoPathBU As String)
    Dim ArchivoDestino As Workbook
    Dim TotalFlexline As Variant
    Dim hojaOrigen As Worksheet
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If archivoOrigenPath = "" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de origen(Unabsorbed Flexline)")
        ' Verifica si se seleccion� un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If

    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPathBU)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("AllocationTotal")
    
    ' Buscar la coincidencia en hojaOrigen
    TotalFlexline = hojaOrigen.Range("D59:O69")
    
    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    ArchivoDestino.Sheets("Non Mat Margin").Range("D168:O178").Value = TotalFlexline
    archivoOrigen.Close SaveChanges:=False
End Sub





