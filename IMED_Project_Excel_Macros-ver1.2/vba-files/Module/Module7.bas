Attribute VB_Name = "Module7"
Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
Dim archivoOrigenPath As String
Sub ActualizarPercentageTABFlexline()
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    
    Dim trimestre1 As Double
    Dim trimestre2 As Double
    Dim trimestre3 As Double
    Dim trimestre4 As Double
    
    Dim trimestreWCStaff1 As Double
    Dim trimestreWCStaff2 As Double
    Dim trimestreWCStaff3 As Double
    Dim trimestreWCStaff4 As Double
    
    Dim trimestreSQFT1 As Double
    Dim trimestreSQFT2 As Double
    Dim trimestreSQFT3 As Double
    Dim trimestreSQFT4 As Double
    
    ' Verifica si ya se ha seleccionado un archivo de origen previamente
    If archivoOrigenPath = "" Then
        ' Abre el cuadro de di?logo de selecci?n de archivo con un t?tulo personalizado para el archivo de origen
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo de origen(BU Scenario Flexline)")
        ' Verifica si se seleccion? un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccion? un archivo, sale del procedimiento
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If ArchivoDestinoPath = "" Then
        ' Abre el cuadro de di?logo de selecci?n de archivo
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de destino(Unabsorbed Flexline)")
        ' Verifica si se seleccion? un archivo
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub ' Si no se seleccion? un archivo, sale del procedimiento
        End If
    End If
    
    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de c?lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("Non Mat Margin")
    Set hojaOrigen2 = archivoOrigen.Sheets("WCStaff Format")

    ' Obtén el número del mes actual
    Dim mesActual As Integer
    mesActual = Month(Date)

    ' Determina el trimestre actual
    Select Case mesActual
        Case 1 To 3
            ' Trimestre 1 (enero a marzo)
            ' Obtiene los valores correspondientes al trimestre
            trimestre1 = Application.WorksheetFunction.Sum(hojaOrigen.Range("D115:F115"))
            trimestreWCStaff1 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("C37:E37"))
            trimestreSQFT1 = Application.WorksheetFunction.Sum(hojaOrigen.Range("D126:F126"))
            ' Calcula el promedio de cada trimestre dividiendo por 3
            trimestre1 = trimestre1 / 3
            trimestreWCStaff1 = trimestreWCStaff1 / 3
            trimestreSQFT1 = trimestreSQFT1 / 3
            ' Coloca los valores obtenidos en celdas espec?ficas de tu hoja de c?lculo principal
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = trimestre1
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = trimestreWCStaff1
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = trimestreSQFT1
        Case 4 To 6
            ' Trimestre 2 (abril a junio)
            ' Obtiene los valores correspondientes al trimestre
            trimestre2 = Application.WorksheetFunction.Sum(hojaOrigen.Range("G115:I115"))
            trimestreWCStaff2 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("F37:H37"))
            trimestreSQFT2 = Application.WorksheetFunction.Sum(hojaOrigen.Range("G126:I126"))
            ' Calcula el promedio de cada trimestre dividiendo por 3
            trimestre2 = trimestre2 / 3
            trimestreWCStaff2 = trimestreWCStaff2 / 3
            trimestreSQFT2 = trimestreSQFT2 / 3
            ' Coloca los valores obtenidos en celdas espec?ficas de tu hoja de c?lculo principal
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = trimestre2
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = trimestreWCStaff2
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = trimestreSQFT2
        Case 7 To 9
            ' Trimestre 3 (julio a septiembre)
            ' Obtiene los valores correspondientes al trimestre
            trimestre3 = Application.WorksheetFunction.Sum(hojaOrigen.Range("J115:L115"))
            trimestreWCStaff3 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("I37:K37"))
            trimestreSQFT3 = Application.WorksheetFunction.Sum(hojaOrigen.Range("J126:L126"))
            ' Calcula el promedio de cada trimestre dividiendo por 3
            trimestre3 = trimestre3 / 3
            trimestreWCStaff3 = trimestreWCStaff3 / 3
            trimestreSQFT3 = trimestreSQFT3 / 3
            ' Coloca los valores obtenidos en celdas espec?ficas de tu hoja de c?lculo principal
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = trimestre3
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = trimestreWCStaff3
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = trimestreSQFT3
        Case 10 To 12
            ' Trimestre 4 (octubre a diciembre)
            ' Obtiene los valores correspondientes al trimestre
            trimestre4 = Application.WorksheetFunction.Sum(hojaOrigen.Range("M115:O115"))
            trimestreWCStaff4 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("L37:N37"))
            trimestreSQFT4 = Application.WorksheetFunction.Sum(hojaOrigen.Range("M126:O126"))
            ' Calcula el promedio de cada trimestre dividiendo por 3
            trimestre4 = trimestre4 / 3
            trimestreWCStaff4 = trimestreWCStaff4 / 3
            trimestreSQFT4 = trimestreSQFT4 / 3
            ' Coloca los valores obtenidos en celdas espec?ficas de tu hoja de c?lculo principal
            ArchivoDestino.Sheets("Percentage").Range("D3").Value = trimestre4
            ArchivoDestino.Sheets("Percentage").Range("D5").Value = trimestreWCStaff4
            ArchivoDestino.Sheets("Percentage").Range("D7").Value = trimestreSQFT4
    End Select
     
    archivoOrigen.Close SaveChanges:=False
End Sub

