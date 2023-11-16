Attribute VB_Name = "Module2"
Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
Dim archivoOrigenPath As String

Sub ObtenerYColocarTotalFlexline()
    Dim ArchivoDestino As Workbook
    Dim TotalFlexline As Variant
    Dim hojaOrigen As Worksheet
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If archivoOrigenPath = "" Then
        ' Abre el cuadro de diálogo de selección de archivo
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de origen(Unabsorbed Flexline)")
        ' Verifica si se seleccionó un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If ArchivoDestinoPath = "" Then
        ' Abre el cuadro de diálogo de selección de archivo
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo de destino(BU Scenario Flexline)")
        ' Verifica si se seleccionó un archivo
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If

    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de cálculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("AllocationTotal")
    
    ' Buscar la coincidencia en hojaOrigen
    TotalFlexline = hojaOrigen.Range("D59:O69")
    
    ' Coloca los valores obtenidos en celdas específicas de tu hoja de cálculo principal
    ArchivoDestino.Sheets("Non Mat Margin").Range("D168:O178").Value = TotalFlexline
    archivoOrigen.Close SaveChanges:=False
End Sub





