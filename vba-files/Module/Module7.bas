Attribute VB_Name = "Module7"
Sub ActualizarPercentageTABFlexline(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaOrigen2 As Worksheet
    Dim ArchivoDestino As Workbook
    
    Dim trimestre As Integer
    Dim mesActual As Integer
    Dim anoFiscal As Integer
    
    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de c?lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("Non Mat Margin")
    Set hojaOrigen2 = archivoOrigen.Sheets("WCStaff Format")

    ' Obtén el número del mes actual
    mesActual = Month(Date)
    
    ' Determina el año fiscal actual
    If mesActual >= 9 Then
        ' Si el mes actual es septiembre o posterior, estás en el año fiscal actual
        anoFiscal = Year(Date) + 1
    Else
        ' Si el mes actual es anterior a septiembre, estás en el año fiscal anterior
        anoFiscal = Year(Date) - 1
    End If

    ' Calcula el trimestre actual
    trimestre = (mesActual - 1) \ 3 + 1
    Debug.Print "trimestre valor: " & trimestre
    ' Obtiene los valores correspondientes al trimestre
    Select Case trimestre
        Case 4
            ' Trimestre 1 (enero a marzo)
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("D115:F115")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = Application.WorksheetFunction.Sum(hojaOrigen2.Range("C37:E37")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("D126:F126")) / 3
        Case 3
            ' Trimestre 2 (abril a junio)
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("G115:I115")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = Application.WorksheetFunction.Sum(hojaOrigen2.Range("F37:H37")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("G126:I126")) / 3
        Case 2
            ' Trimestre 3 (julio a septiembre)
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("J115:L115")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = Application.WorksheetFunction.Sum(hojaOrigen2.Range("I37:K37")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("J126:L126")) / 3
        Case 1
            ' Trimestre 4 (octubre a diciembre)
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("M115:O115")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = Application.WorksheetFunction.Sum(hojaOrigen2.Range("L37:N37")) / 3
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = Application.WorksheetFunction.Sum(hojaOrigen.Range("M126:O126")) / 3
    End Select
End Sub

