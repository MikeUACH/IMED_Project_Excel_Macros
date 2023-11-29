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
    ' Cambia el color de fondo del Label al hacer clic en él
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarShifts
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Shift Tabs WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label7_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarShifts
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Shift Tabs WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label8_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label1.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label1.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarShifts
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Shift Tabs WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Label2_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label2.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label2.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarPercentageTAB
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Percentage tabs en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label9_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label2.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label2.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarPercentageTAB
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Percentage tabs en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label10_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label2.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label2.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarPercentageTAB
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Percentage tabs en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Label3_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label3.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label3.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarTABRateCalc
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Rate Calculation en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label12_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label3.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label3.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarTABRateCalc
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Rate Calculation en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label11_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label3.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label3.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ActualizarTABRateCalc
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Rate Calculation en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Label5_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label5.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label5.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    RealizarOperaciones
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Income Statement en BU_Scenario_Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label14_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label5.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label5.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    RealizarOperaciones
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Income Statement en BU_Scenario_Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label13_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label5.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label5.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    RealizarOperaciones
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Income Statement en BU_Scenario_Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Label4_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label4.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label4.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarTotalFlexline
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Non Mat Margin en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label16_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label4.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label4.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarTotalFlexline
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Non Mat Margin en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub
Private Sub Label15_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label4.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label4.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarTotalFlexline
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en Non Mat Margin en Unabsorbed- Flexline Calculation"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub Label6_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label6.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label6.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarWCStaffFormat
    
    ' Registrar la acción en la hoja de registro
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
    ' Ocultar todas las columnas que no contienen registros
    Dim allColumns As Range
    Set allColumns = wsRegistro.UsedRange.EntireColumn
    Dim column As Range
    
    For Each column In allColumns
        If Application.WorksheetFunction.CountA(column) = 0 Then
            column.Hidden = True
        Else
            column.Hidden = False
        End If
    Next column
End Sub
Private Sub Label18_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label6.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label6.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarWCStaffFormat
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
    ' Ocultar todas las columnas que no contienen registros
    Dim allColumns As Range
    Set allColumns = wsRegistro.UsedRange.EntireColumn
    Dim column As Range
    
    For Each column In allColumns
        If Application.WorksheetFunction.CountA(column) = 0 Then
            column.Hidden = True
        Else
            column.Hidden = False
        End If
    Next column
End Sub
Private Sub Label17_Click()
    ' Cambia el color de fondo del Label al hacer clic en él
    Label6.BackColor = RGB(240, 240, 240) ' Gris claro al hacer clic en el Label
    Dim StartTime As Double
    StartTime = Timer ' Guarda el tiempo de inicio
    
    Do
        DoEvents ' Permite que otros eventos se procesen
    Loop While Timer < StartTime + 0.099 ' Espera 1 segundo (ajusta según sea necesario)
    
    Label6.BackColor = RGB(255, 255, 255) ' Vuelve a color blanco
    ObtenerYColocarWCStaffFormat
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acción realizada en WCStaff Format en BU Scenario Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub

