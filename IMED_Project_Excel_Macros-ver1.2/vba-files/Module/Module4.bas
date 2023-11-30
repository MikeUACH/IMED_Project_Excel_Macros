Attribute VB_Name = "Module4"
Dim archivoOrigenPath As String ' Variable global para almacenar la ruta del archivo origen
Dim ArchivoDestinoPath As String

Sub Inutil()
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    Dim rangoOrigen As Range
    Dim rangoDestino As Range
    Dim rangoNuevosNumeros As Range
    
    Dim NuevoRango As Range
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If archivoOrigenPath = "" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo de origen(BU Scenario Flexline)")
        ' Verifica si se seleccion� un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If ArchivoDestinoPath = "" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de destino(Unabsorbed Flexline)")
        ' Verifica si se seleccion� un archivo
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    
    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaDestino = ArchivoDestino.Sheets("AllocationTotal")
    
    ' Define la hoja de destino en tu libro actual
    Set hojaOrigen = archivoOrigen.Sheets("Income Statement")
    
    ' Define el rango de celdas de origen (D70:O70) y destino (D34:O34)
    Set rangoOrigen = hojaDestino.Range("D71:O71")
    Set rangoDestino = hojaOrigen.Range("D34:O34")
    
    rangoOrigen.Copy
    rangoDestino.PasteSpecial xlPasteValues

    Set rangoNuevosNumeros = hojaOrigen.Range("D34:O34")
    Set NuevoRango = hojaDestino.Range("D72:O72")
    
    ' Pega los valores del calculo de las celdas D37:O37 hacia las celdas D34:O34
    rangoNuevosNumeros.Copy
    rangoDestino.PasteSpecial xlPasteValues

    rangoDestino.Copy
    NuevoRango.PasteSpecial xlPasteValues
    archivoOrigen.Close SaveChanges:=False
End Sub





