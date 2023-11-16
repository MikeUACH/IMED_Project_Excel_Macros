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
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de destino (Pruebas)")
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub
        End If
    End If
    
    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    ' Define la hoja de c?lculo en el archivo de destino
    Set hojaDestino = ArchivoDestino.Sheets("Sheet1")
    
    ' Verifica si la variable RangoHojaDestino está definida
    If RangoHojaDestino > 0 Then
        MsgBox "Se borro correctamente el rango", vbExclamation
        txtDestinoSTWF.Value = ""
        ' Borra el contenido del rango en la hoja de destino
        hojaDestino.Range(RangoHojaDestino).ClearContents
    Else
        MsgBox "La variable RangoHojaDestino no está definida. Ejecute primero la subrutina ObtenerYColocarShifts.", vbExclamation
    End If
    
    ' Cierra el archivo de destino sin guardar cambios
    ArchivoDestino.Close SaveChanges:=False
End Sub
Private Sub btnSeleccionarDestinoSTWF_Click()
    With Me.txtDestinoSTWF
    
    End With
End Sub

Private Sub UserForm_Initialize()
    ' Llama a la subrutina para obtener y colocar los turnos
    ObtenerYColocarShifts
    
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
