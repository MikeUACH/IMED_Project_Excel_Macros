Attribute VB_Name = "Module3"
Dim archivoOrigenPath As String
Sub UpdWCellTabBU(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim ArchivoDestino As Workbook
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim rangoTabla As Range
    Dim celdaInicio As Range
    Dim rangoDatos As Range
    Dim filaInicio As Long
    
    ' Usa el archivo seleeccionado como ArchivoDestino
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    
    ' Define la hoja de c�lculo en el archivo de origen (en este caso, la hoja "IMED")
    Set hojaOrigen = archivoOrigen.Sheets("IMED")
    
    ' Define el rango de la tabla excluyendo las celdas A34:M35 (ajusta los n�meros de filas y columnas seg�n tu tabla)
    Set rangoTabla = Union(hojaOrigen.Range("A1:M33"), hojaOrigen.Range("A36").Resize(13, 36))
    
    ' Buscar la palabra "NUEVO Forecast" en la tabla
    On Error Resume Next
    Set celdaInicio = rangoTabla.Find("NUEVO Forecast")
    On Error GoTo 0
    
    ' Verificar si se encontr� la palabra
    If Not celdaInicio Is Nothing Then
        ' Obtener la fila de inicio donde se encontr� la palabra
        filaInicio = celdaInicio.Row
        
        ' Definir el rango de datos a copiar excluyendo la fila donde se encuentra la palabra (36 columnas por 13 filas)
        Set rangoDatos = hojaOrigen.Range("A" & filaInicio + 1 & ":AL" & filaInicio + 13)
        
        ' Copiar los datos al archivo de trabajo actual (ajusta el rango de destino seg�n tus necesidades)
        ArchivoDestino.Sheets("WCStaff Format").Range("B3").Resize(13, 33).Value = rangoDatos.Value
    Else
        MsgBox "La palabra 'NUEVO Forecast' no se encontro en la tabla.", vbExclamation
    End If
    archivoOrigen.Close SaveChanges:=False
End Sub


