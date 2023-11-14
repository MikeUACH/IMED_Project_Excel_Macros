Attribute VB_Name = "Module7"
Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
Dim archivoOrigenPath As String
Sub ActualizarPercentageTAB()
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
        ' Abre el cuadro de diálogo de selección de archivo con un título personalizado para el archivo de origen
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo de origen(BU Scenario Flexline)")
        ' Verifica si se seleccionó un archivo
        If archivoOrigenPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo previamente
    If ArchivoDestinoPath = "" Then
        ' Abre el cuadro de diálogo de selección de archivo
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo de destino(Unabsorbed Flexline)")
        ' Verifica si se seleccionó un archivo
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub ' Si no se seleccionó un archivo, sale del procedimiento
        End If
    End If
    
    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de cï¿½lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("Non Mat Margin")
    Set hojaOrigen2 = archivoOrigen.Sheets("WCStaff Format")
    
    ' Obtiene los valores correspondientes al trimestre 1 (por ejemplo, de enero a marzo)
    trimestre1 = Application.WorksheetFunction.Sum(hojaOrigen.Range("D115:F115"))
    
    ' Obtiene los valores correspondientes al trimestre 2 (por ejemplo, de abril a junio)
    trimestre2 = Application.WorksheetFunction.Sum(hojaOrigen.Range("G115:I115"))
    
    ' Obtiene los valores correspondientes al trimestre 3 (por ejemplo, de julio a septiembre)
    trimestre3 = Application.WorksheetFunction.Sum(hojaOrigen.Range("J115:L115"))
    
    ' Obtiene los valores correspondientes al trimestre 4 (por ejemplo, de octubre a diciembre)
    trimestre4 = Application.WorksheetFunction.Sum(hojaOrigen.Range("M115:O115"))
    
    ' Calcula el promedio de cada trimestre dividiendo por 3
    trimestre1 = trimestre1 / 3
    trimestre2 = trimestre2 / 3
    trimestre3 = trimestre3 / 3
    trimestre4 = trimestre4 / 3
    
    ' Obtiene los valores correspondientes al trimestre
    trimestreWCStaff1 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("C37:E37"))
    trimestreWCStaff2 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("F37:H37"))
    trimestreWCStaff3 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("I37:K37"))
    trimestreWCStaff4 = Application.WorksheetFunction.Sum(hojaOrigen2.Range("L37:N37"))
    
    ' Calcula el promedio de cada trimestre dividiendo por 3
    trimestreWCStaff1 = trimestreWCStaff1 / 3
    trimestreWCStaff2 = trimestreWCStaff2 / 3
    trimestreWCStaff3 = trimestreWCStaff3 / 3
    trimestreWCStaff4 = trimestreWCStaff4 / 3
    
    ' Obtiene los valores correspondientes al trimestre
    trimestreSQFT1 = Application.WorksheetFunction.Sum(hojaOrigen.Range("D126:F126"))
    trimestreSQFT2 = Application.WorksheetFunction.Sum(hojaOrigen.Range("G126:I126"))
    trimestreSQFT3 = Application.WorksheetFunction.Sum(hojaOrigen.Range("J126:L126"))
    trimestreSQFT4 = Application.WorksheetFunction.Sum(hojaOrigen.Range("M126:O126"))
    
    ' Calcula el promedio de cada trimestre dividiendo por 3
    trimestreSQFT1 = trimestreSQFT1 / 3
    trimestreSQFT2 = trimestreSQFT2 / 3
    trimestreSQFT3 = trimestreSQFT3 / 3
    trimestreSQFT4 = trimestreSQFT4 / 3
    
    
    ' Coloca los valores obtenidos en celdas especï¿½ficas de tu hoja de cï¿½lculo principal
    ArchivoDestino.Sheets("Percentage").Range("D3").value = trimestre1
    ArchivoDestino.Sheets("Percentage").Range("D25").value = trimestre2
    ArchivoDestino.Sheets("Percentage").Range("D47").value = trimestre3
    ArchivoDestino.Sheets("Percentage").Range("D69").value = trimestre4
    
    ' Coloca los valores obtenidos en celdas especï¿½ficas de tu hoja de cï¿½lculo principal
    ArchivoDestino.Sheets("Percentage").Range("D5").value = trimestreWCStaff1
    ArchivoDestino.Sheets("Percentage").Range("D27").value = trimestreWCStaff2
    ArchivoDestino.Sheets("Percentage").Range("D49").value = trimestreWCStaff3
    ArchivoDestino.Sheets("Percentage").Range("D71").value = trimestreWCStaff4
    
    ' Coloca los valores obtenidos en celdas especï¿½ficas de tu hoja de cï¿½lculo principal
    ArchivoDestino.Sheets("Percentage").Range("D7").value = trimestreSQFT1
    ArchivoDestino.Sheets("Percentage").Range("D29").value = trimestreSQFT2
    ArchivoDestino.Sheets("Percentage").Range("D51").value = trimestreSQFT3
    ArchivoDestino.Sheets("Percentage").Range("D73").value = trimestreSQFT4
    archivoOrigen.Close SaveChanges:=False
End Sub




