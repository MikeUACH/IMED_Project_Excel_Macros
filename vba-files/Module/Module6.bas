Attribute VB_Name = "Module6"
Sub ActualizarTABRateCalcFlex(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim valores(1 To 7, 1 To 4) As Double
    
    ' Abre el archivo de origen seleccionado (usa la ruta almacenada)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("Income Statement")

    ' Resto del c�digo para obtener y colocar valores aqu�...
    ' Obtiene los valores de Q1, Q2, Q3 y Q4 directamente desde las celdas
    valores(1, 1) = hojaOrigen.Range("T10").Value
    valores(1, 2) = hojaOrigen.Range("U10").Value
    valores(1, 3) = hojaOrigen.Range("V10").Value
    valores(1, 4) = hojaOrigen.Range("W10").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(2, 1) = hojaOrigen.Range("T11").Value
    valores(2, 2) = hojaOrigen.Range("U11").Value
    valores(2, 3) = hojaOrigen.Range("V11").Value
    valores(2, 4) = hojaOrigen.Range("W11").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(3, 1) = hojaOrigen.Range("T14").Value
    valores(3, 2) = hojaOrigen.Range("U14").Value
    valores(3, 3) = hojaOrigen.Range("V14").Value
    valores(3, 4) = hojaOrigen.Range("W14").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(4, 1) = hojaOrigen.Range("T15").Value
    valores(4, 2) = hojaOrigen.Range("U15").Value
    valores(4, 3) = hojaOrigen.Range("V15").Value
    valores(4, 4) = hojaOrigen.Range("W15").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(5, 1) = hojaOrigen.Range("T16").Value
    valores(5, 2) = hojaOrigen.Range("U16").Value
    valores(5, 3) = hojaOrigen.Range("V16").Value
    valores(5, 4) = hojaOrigen.Range("W16").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(6, 1) = hojaOrigen.Range("T23").Value
    valores(6, 2) = hojaOrigen.Range("U23").Value
    valores(6, 3) = hojaOrigen.Range("V23").Value
    valores(6, 4) = hojaOrigen.Range("W23").Value
    
    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(7, 1) = hojaOrigen.Range("T12").Value
    valores(7, 2) = hojaOrigen.Range("U12").Value
    valores(7, 3) = hojaOrigen.Range("V12").Value
    valores(7, 4) = hojaOrigen.Range("W12").Value
    ' Cierra el archivo de origen sin guardar cambios
    ' archivoOrigen.Close SaveChanges:=False

    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    For i = 1 To 7 ' Cambia el l�mite para i a 6
        For j = 1 To 4 ' Cambia el l�mite para j a 4
            ArchivoDestino.Sheets("Rate Calculation").Cells(i + 2, j + 31).Value = valores(i, j)
        Next j
    Next i
End Sub







