VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Inicio"
   ClientHeight    =   9690.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ObtenerYColocarTabsUnabFlex
End Sub

Private Sub Label1_Click()
    ' Cambia el color de fondo del Label al hacer clic en �l
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta seg�n sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ' L�gica para seleccionar autom�ticamente el archivo destino
    ArchivoDestinoPathSC = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo de destino (BU Scenario Flexline)")
    
    ' Verifica si se seleccion� un archivo
    If ArchivoDestinoPathSC = "Falso" Then
        Exit Sub
    End If
    UpdWCstaffShiftTabsBU ArchivoDestinoPathSC  ' Llama al m�dulo 1
    ' UpdNonMatMarginBU ArchivoDestinoPathBU    ' Llama al m�dulo 2
    ' UpdWCellTabBU ArchivoDestinoPathBU   ' Llama al m�dulo 3
    ' ActualizarPercentageTABFlexline ' Llama al m�dulo 6
    ' ActualizarTABRateCalcFlex   ' Llama al m�dulo 7
    ' ObtenerYColocarTabsUnabFlex ' Llama al m�dulo 8
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acci�n realizada en archivos BU y Flexline, tambi�n se gener� el reporte"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label7_Click()
    ' Cambia el color de fondo del Label al hacer clic en �l
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta seg�n sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    SeleccionarArchivoDestino
    UpdWCstaffShiftTabsBU ' Llama al m�dulo 1
    UpdNonMatMarginBU   ' Llama al m�dulo 2
    UpdWCellTabBU   ' Llama al m�dulo 3
    ActualizarPercentageTABFlexline ' Llama al m�dulo 6
    ActualizarTABRateCalcFlex   ' Llama al m�dulo 7
    ObtenerYColocarTabsUnabFlex ' Llama al m�dulo 8
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acci�n realizada en archivos BU y Flexline, tambi�n se genero el reporte"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label8_Click()
    ' Cambia el color de fondo del Label al hacer clic en �l
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta seg�n sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    SeleccionarArchivoDestino
    UpdWCstaffShiftTabsBU ' Llama al m�dulo 1
    UpdNonMatMarginBU   ' Llama al m�dulo 2
    UpdWCellTabBU   ' Llama al m�dulo 3
    ActualizarPercentageTABFlexline ' Llama al m�dulo 6
    ActualizarTABRateCalcFlex   ' Llama al m�dulo 7
    ObtenerYColocarTabsUnabFlex ' Llama al m�dulo 8
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acci�n realizada en archivos BU y Flexline, tambi�n se genero el reporte"
    wsRegistro.Columns("A:B").AutoFit
End Sub
