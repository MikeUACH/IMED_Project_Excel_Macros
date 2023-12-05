VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    ' Lï¿½gica para seleccionar automï¿½ticamente el archivo destino
    ArchivoDestinoPathSC = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo BU Scenario Flexline")
    
    ' Verifica si se seleccionï¿½ un archivo
    If ArchivoDestinoPathSC = "Falso" Then
        Exit Sub
    End If

    ArchivoOrigenPathUF = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo Unabsorbed Flexline")
    
    ' Verifica si se seleccionï¿½ un archivo
    If ArchivoOrigenPathUF = "Falso" Then
        Exit Sub
    End If

    UpdWCstaffShiftTabsBU ArchivoDestinoPathSC  ' Llama al mï¿½dulo 1
    UpdNonMatMarginBU ArchivoDestinoPathSC, ArchivoOrigenPathUF     ' Llama al mï¿½dulo 2
    UpdWCellTabBU ArchivoDestinoPathSC   ' Llama al mï¿½dulo 3
    ActualizarPercentageTABFlexline ArchivoDestinoPathSC, ArchivoOrigenPathUF ' Llama al mï¿½dulo 6
    ActualizarTABRateCalcFlex ArchivoDestinoPathSC, ArchivoOrigenPathUF  ' Llama al mï¿½dulo 7
    ObtenerYColocarTabsUnabFlex ArchivoOrigenPathUF ' Llama al mï¿½dulo 8
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acciónn realizada en archivos BU y Flexline, también se generó el reporte"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Worksheet_Activate()
    ' Lï¿½gica para seleccionar automï¿½ticamente el archivo destino
    ArchivoDestinoPathSC = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo BU Scenario Flexline")
    
    ' Verifica si se seleccionï¿½ un archivo
    If ArchivoDestinoPathSC = "Falso" Then
        Exit Sub
    End If

    ArchivoOrigenPathUF = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo Unabsorbed Flexline")
    
    ' Verifica si se seleccionï¿½ un archivo
    If ArchivoOrigenPathUF = "Falso" Then
        Exit Sub
    End If

    UpdWCstaffShiftTabsBU ArchivoDestinoPathSC  ' Llama al mï¿½dulo 1
    UpdNonMatMarginBU ArchivoDestinoPathSC, ArchivoOrigenPathUF     ' Llama al mï¿½dulo 2
    UpdWCellTabBU ArchivoDestinoPathSC   ' Llama al mï¿½dulo 3
    ActualizarPercentageTABFlexline ArchivoDestinoPathSC, ArchivoOrigenPathUF ' Llama al mï¿½dulo 6
    ActualizarTABRateCalcFlex ArchivoDestinoPathSC, ArchivoOrigenPathUF  ' Llama al mï¿½dulo 7
    ObtenerYColocarTabsUnabFlex ArchivoOrigenPathUF ' Llama al mï¿½dulo 8
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acciónn realizada en archivos BU y Flexline, también se generó el reporte"
    wsRegistro.Columns("A:B").AutoFit
End Sub
