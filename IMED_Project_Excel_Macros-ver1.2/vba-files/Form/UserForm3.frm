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
Dim PathBU As String
Dim PathDL As String
Dim PathWC As String
Dim PathFlex As String
Dim PathVariance As String
Private Sub UserForm_Initialize()
    ActualizarEstadoBoton
End Sub
' Hacer un If que haga que seleccione archivo con nombre similar
Private Sub btnSeleccionarBU_Click()
    If PathBU = "" Or PathBU = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathBU = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo BU Scenario Flexline")
        ' Verifica si se seleccion� un archivo
        If PathBU = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    Debug.Print "Path BU: " & PathBU
    ActualizarEstadoBoton
End Sub

Private Sub btnSeleccionarDL_Click()
    If PathDL = "" Or PathDL = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathDL = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo DL Breakdown")
        ' Verifica si se seleccion� un archivo
        If PathDL = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    
    ActualizarEstadoBoton
    Debug.Print "Path DL: " & PathDL
End Sub

Private Sub btnSeleccionarWC_Click()
    If PathWC = "" Or PathWC = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathWC = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo WC Staff")
        ' Verifica si se seleccion� un archivo
        If PathWC = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    ActualizarEstadoBoton
    Debug.Print "Path WC: " & PathWC
End Sub

Private Sub btnSeleccionarFlex_Click()
    If PathFlex = "" Or PathFlex = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathFlex = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo Flexline Unabsorbed-Calculation")
        ' Verifica si se seleccion� un archivo
        If PathFlex = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    
    ActualizarEstadoBoton
    Debug.Print "Path Flex: " & PathFlex
End Sub

Private Sub btnSeleccionarVariance_Click()
    If PathVariance = "" Or PathVariance = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathVariance = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona el archivo Variance BID")
        ' Verifica si se seleccion� un archivo
        If PathVariance = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
    End If
    
    ActualizarEstadoBoton
    Debug.Print "Path Variance: " & PathVariance
End Sub
    
Private Sub btnActualizar_Click()
    ' Hacer un If para verificar que se han seleccionado los archivos necesarios
    UpdWCstaffShiftTabsBU PathDL, PathBU  ' Llama al m�dulo 1
    UpdNonMatMarginBU PathFlex, PathBU    ' Llama al m�dulo 2
    UpdWCellTabBU PathWC, PathBU   ' Llama al m�dulo 3
    ActualizarPercentageTABFlexline PathBU, PathFlex ' Llama al m�dulo 6
    ActualizarTABRateCalcFlex PathBU, PathFlex  ' Llama al m�dulo 7
    
    
    Dim wsRegistro As Worksheet
    Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
    Dim lastRow As Long
    lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
    wsRegistro.Cells(lastRow, 1).Value = Now
    wsRegistro.Cells(lastRow, 2).Value = "Acci�n realizada en archivos BU y Flexline"
    wsRegistro.Columns("A:B").AutoFit
End Sub

Private Sub btnGenerarReporte_Click()
   ObtenerYColocarTabsUnabFlex PathFlex, PathVariance ' Llama al m�dulo 8
End Sub

Private Sub btnBorrarUbicacionBU_Click()
    If Len(PathBU) > 0 Then
        PathBU = "False"
        If PathBU = "False" Then
            Dim comprobarBU As Boolean
            comprobarBU = True
        End If

        If comprobarBU = True Then
            PathBU = ""
            MsgBox "Se ha borrado con éxito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path BU: " & PathBU
    End If

    If Len(PathBU) = 0 And comprobarBU = False Then
        MsgBox "No hay ningÃºn archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicacionDL_Click()
    If Len(PathDL) > 0 Then
        PathDL = "False"
        If PathDL = "False" Then
            Dim comprobarDL As Boolean
            comprobarDL = True
        End If

        If comprobarDL = True Then
            PathDL = ""
            MsgBox "Se ha borrado con éxito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path DL: " & PathDL
    End If

    If Len(PathDL) = 0 And comprobarDL = False Then
        MsgBox "No hay ningÃºn archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicacionWC_Click()
    If Len(PathWC) > 0 Then
        PathWC = "False"
        If PathWC = "False" Then
            Dim comprobarWC As Boolean
            comprobarWC = True
        End If

        If comprobarWC = True Then
            PathWC = ""
            MsgBox "Se ha borrado con éxito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path WC: " & PathWC
    End If

    If Len(PathWC) = 0 And comprobarWC = False Then
        MsgBox "No hay ningÃºn archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicacionFlex_Click()
    If Len(PathFlex) > 0 Then
        PathFlex = "False"
        If PathFlex = "False" Then
            Dim comprobarFlex As Boolean
            comprobarFlex = True
        End If

        If comprobarFlex = True Then
            PathFlex = ""
            MsgBox "Se ha borrado con éxito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path Flex: " & PathFlex
    End If

    If Len(PathFlex) = 0 And comprobarFlex = False Then
        MsgBox "No hay ningÃºn archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicacionVariance_Click()
    If Len(PathVariance) > 0 Then
        PathVariance = "False"
        If PathVariance = "False" Then
            Dim comprobarVariance As Boolean
            comprobarVariance = True
        End If

        If comprobarVariance = True Then
            PathVariance = ""
            MsgBox "Se ha borrado con éxito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path Variance: " & PathVariance
    End If

    If Len(PathVariance) = 0 And comprobarVariance = False Then
        MsgBox "No hay ningÃºn archivo seleccionado"
    End If
End Sub

Private Sub ActualizarEstadoBoton()
    If Len(PathBU) = 0 Or PathBU = "False" Then
        btnSeleccionarBU.BackColor = RGB(255, 172, 172)
        txtNotSelectedBU.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarBU.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoBU As String
        nombreArchivoBU = Mid(PathBU, InStrRev(PathBU, "\") + 1)
        txtNotSelectedBU.Caption = "Seleccionado: " & nombreArchivoBU
    End If
    
    If Len(PathDL) = 0 Or PathDL = "False" Then
        btnSeleccionarDL.BackColor = RGB(255, 172, 172)
        txtNotSelectedDL.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarDL.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoDL As String
        nombreArchivoDL = Mid(PathDL, InStrRev(PathDL, "\") + 1)
        txtNotSelectedDL.Caption = "Seleccionado: " & nombreArchivoDL
    End If
    
    If Len(PathWC) = 0 Or PathWC = "False" Then
        btnSeleccionarWC.BackColor = RGB(255, 172, 172)
        txtNotSelectedWC.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarWC.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoWC As String
        nombreArchivoWC = Mid(PathWC, InStrRev(PathWC, "\") + 1)
        txtNotSelectedWC.Caption = "Seleccionado: " & nombreArchivoWC
    End If
    
    If Len(PathFlex) = 0 Or PathFlex = "False" Then
        btnSeleccionarFlex.BackColor = RGB(255, 172, 172)
        txtNotSelectedFX.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarFlex.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoFlex As String
        nombreArchivoFlex = Mid(PathFlex, InStrRev(PathFlex, "\") + 1)
        txtNotSelectedFX.Caption = "Seleccionado: " & nombreArchivoFlex
    End If
    
    If Len(PathVariance) = 0 Or PathVariance = "False" Then
        btnSeleccionarVariance.BackColor = RGB(255, 172, 172)
        txtNotSelectedVariance.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarVariance.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoVariance As String
        nombreArchivoVariance = Mid(PathVariance, InStrRev(PathVariance, "\") + 1)
        txtNotSelectedVariance.Caption = "Seleccionado: " & nombreArchivoVariance
    End If
End Sub
