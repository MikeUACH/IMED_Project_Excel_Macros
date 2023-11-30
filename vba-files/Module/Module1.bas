Attribute VB_Name = "Module1"
' Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
' Dim archivoOrigenPath As String


' Public RangoHojaOrigen() As Variant ' Variable para almacenar los rangos de la hoja de origen
' Public RangoHojaDestino As Range ' Variable para almacenar el rango de la hoja de destino


' ' Redimensionar el arreglo para almacenar los rangos
' ReDim RangoHojaOrigen(1 To 7, 1 To 7)

' ' Procesar cada turno
' For turno = 1 To 7
'     Dim turnoNombre As String
'     turnoNombre = nombresTurnos(turno - 1)
    
'     ' Obtener rango de celdas B9:B15 de la hoja de configuraci?n
'     Dim rangoTurno As Range
'     Set rangoTurno = ThisWorkbook.Sheets("hojaConfiguracion").Range("B9:B15")
    
'     ' Iterar a trav?s de cada celda del rango y asignar sus valores
'     Dim i As Integer
'     i = 1 ' Iniciar desde la primera posici?n en la matriz
'     For Each celda In rangoTurno
'         RangoHojaOrigen(turno, i) = celda.Value
'         i = i + 1 ' Mover a la siguiente posici?n en la matriz
'     Next celda
    
'     ' Definir el rango de la hoja de destino
'     Set RangoHojaDestino = ThisWorkbook.Sheets("hojaConfiguracion").Range("B16")
    
    
' Next turno

Sub ObtenerYColocarShifts()
    Dim ArchivoDestino As Workbook
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaConfiguracion As Worksheet
    Dim turno As Integer
    
    archivoOrigenPath = "C:\Users\3762091\Desktop\trabajo\Proyecto Excel\Excel\NUEVOS\08 DL Breakdown FY24 Q1 BID 4 Rev F.xlsx"
    ' Verifica si ya se ha seleccionado un archivo de origen previamente
    If archivoOrigenPath = "" Then
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de origen (DL Breakdown)")
        If archivoOrigenPath = "Falso" Then
            Exit Sub
        End If
    End If

    ' Verifica si ya se ha seleccionado un archivo de destino previamente
    If ArchivoDestinoPath = "" Then
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xls), *.xlsx", , "Selecciona el archivo de destino (BU Scenario Flexline)")
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub
        End If
    End If

    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)

    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("IMED DL Breakdow")
    
    ' Define la hoja de configuraci�n en el archivo de origen
    Set hojaConfiguracion = ThisWorkbook.Sheets("hojaConfiguracion")

    ' Definir el rango de destino fijo
    Dim rangoDestino As String
    rangoDestino = hojaConfiguracion.Range("B16").Value

    ' Definir los nombres de los turnos
    Dim nombresTurnos() As String
    nombresTurnos = Split("FirstShift,SecondShift,ThirdShift,FourTwentyShift,FourTwentyOneShift,FourTwentyTwoShift,FourTwentyThreeShift", ",")

    ' Procesar cada turno
    For turno = 1 To 7
        Dim turnoNombre As String
        turnoNombre = nombresTurnos(turno - 1)

        ' Obtener el rango de origen de la hoja de configuraci�n
        Dim rangoOrigen As String
        rangoOrigen = hojaConfiguracion.Range("B" & turno + 8).Value

        ' Copiar datos del rango de origen al rango de destino
        ArchivoDestino.Sheets("Sheet1").Range(rangoDestino).Value = hojaOrigen.Range(rangoOrigen).Value
    Next turno

    ' Cerrar el archivo de origen sin guardar cambios
    archivoOrigen.Close SaveChanges:=False
End Sub

