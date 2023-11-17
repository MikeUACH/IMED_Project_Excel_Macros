VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9960.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20295
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBorrarDestinoSTWF_Click()
    Dim ArchivoDestino As Workbook
    Dim hojaDestino As Worksheet

    ' Verifica si ya se ha seleccionado un archivo de destino previamente
    If ArchivoDestinoPath = "" Then
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de destino (BU Scenario Flexline)")
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub
        End If
    End If

    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)

    ' Define la hoja de cálculo en el archivo de destino
    Set hojaDestino = ArchivoDestino.Sheets("Sheet1")
    
    ' Establece la variable RealizarAsignacion en False
    RealizarAsignacion = False
    
    ' Verifica si la variable RangoHojaDestino está definida
    If RangoHojaDestino <> "" Then
        ' Borra el contenido del rango en la hoja de destino
        hojaDestino.Range(RangoHojaDestinoGlobal).ClearContents
        hojaDestino.Range(RangoHojaDestino).ClearContents
        txtDestinoSTWF.Value = ""
        ' Reinicia la variable RangoHojaDestino a una cadena vacía en el módulo
        RangoHojaDestinoGlobal = ""
        ' Asigna una cadena vacía a la variable global en el módulo
        RangoHojaDestino = ""
    Else
        MsgBox "La variable RangoHojaDestino no está definida. Ejecute primero la subrutina ObtenerYColocarShifts.", vbExclamation
    End If
    
    ' Mostrar el mensaje solo si se borró algo
    If RangoHojaDestino <> "" Then
        MsgBox "Se borró correctamente.", vbExclamation
    End If
End Sub

Private Sub btnSeleccionarDestinoSTWF_Click()
    Dim rangoSeleccionado As Range
    Dim ArchivoDestino As Workbook
    Dim hojaDestino As Worksheet

    ' Verifica si ya se ha seleccionado un archivo de destino previamente
    If ArchivoDestinoPath = "" Then
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de destino (BU Scenario Flexline)")
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub
        End If
    End If
    
    ' Define la hoja de configuración en el libro actual (ajusta el nombre según tus necesidades)
    Set hojaConfiguracion = ThisWorkbook.Sheets("hojaConfiguracion")
    
    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)

    ' Define la hoja de cálculo en el archivo de destino
    Set hojaDestino = ArchivoDestino.Sheets("Sheet1")
    ' Solicitar al usuario que seleccione un rango
    On Error Resume Next
    Set rangoSeleccionado = Application.InputBox("Selecciona el rango de datos que deseas utilizar", Type:=8)
    On Error GoTo 0
    
    ' Verificar si el usuario seleccionó un rango
    If Not rangoSeleccionado Is Nothing Then
        ' Guardar el rango en la celda B10 de la hoja de configuración
        hojaConfiguracion.Range("B10").Value = rangoSeleccionado.Address
        hojaConfiguracion.Columns("B:B").AutoFit ' Ajustar automáticamente el ancho de la columna B
        MsgBox "Rango seleccionado: " & rangoSeleccionado.Address
    Else
        MsgBox "Operación cancelada por el usuario.", vbInformation
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Llama a la subrutina para obtener y colocar los turnos
    
    ' Muestra el contenido de la variable RangoHojaOrigen en el TextBox
    With Me.txtOrigenSTWF
        ' Utiliza la función Join para concatenar los elementos del array en una cadena
        .Value = Join(RangoHojaOrigen, vbNewLine)
    End With
    With Me.txtDestinoSTWF
        .Value = RangoHojaDestino
    End With
End Sub
Private Sub CommandButton25_Click()
    Unload Me
    UserForm2.Show
End Sub

Private Sub Image1_Click()
    Unload Me
    UserForm3.Show
End Sub
