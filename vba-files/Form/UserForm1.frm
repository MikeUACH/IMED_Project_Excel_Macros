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
    ' Verifica si la variable RangoHojaDestino esta definida
    If RangoCeldaSTFW <> "" Then
        txtDestinoSTWF.Value = ""
        ' Reinicia la variable RangoHojaDestino a una cadena vac�a en el m�dulo
        ThisWorkbook.Sheets("hojaConfiguracion").Range("B16").ClearContents
        If RangoCeldaSTFW = "" Then
            MsgBox "Se borro correctamente.", vbExclamation
        End If
    Else
        MsgBox "El rango no est� definido.", vbExclamation
    End If
    
    ' Mostrar el mensaje solo si se borr� algo
    
End Sub

Private Sub btnSeleccionarDestinoSTWF_Click()
    Dim rangoSeleccionado As Range
    Dim ArchivoDestino As Workbook
    Dim hojaDestino As Worksheet

    ' Define la hoja de configuraci�n en el libro actual (ajusta el nombre seg�n tus necesidades)
    Set hojaConfiguracion = ThisWorkbook.Sheets("hojaConfiguracion")

    ' Solicitar al usuario que seleccione un rango
    On Error Resume Next
    Set rangoSeleccionado = Application.InputBox("Selecciona el rango de datos que deseas utilizar", Type:=8)
    On Error GoTo 0
    
    ' Verificar si el usuario seleccion� un rango
    If Not rangoSeleccionado Is Nothing Then
        ' Guardar el rango en la celda B10 de la hoja de configuraci�n
        hojaConfiguracion.Range("B16").Value = rangoSeleccionado.Address
        MsgBox "Rango seleccionado: " & rangoSeleccionado.Address
    Else
        MsgBox "Operaci�n cancelada por el usuario.", vbInformation
    End If
End Sub
Private Sub UserForm_Initialize()
    ' Obtener el rango de celdas B9:B15 desde la hoja de configuraci�n
    Dim rangoConfiguracion As Range
    Set rangoConfiguracion = ThisWorkbook.Sheets("hojaConfiguracion").Range("B9:B15")
    
    ' Obtener los valores del rango y asignarlos a un array
    Dim valoresConfiguracion() As Variant
    valoresConfiguracion = rangoConfiguracion.Value

    ' Mostrar los valores en txtOrigenSTWF
    With Me.txtOrigenSTWF
        ' Utilizar Join para concatenar los elementos del array en una cadena con saltos de l�nea
        .Value = Join(Application.WorksheetFunction.Transpose(valoresConfiguracion), vbCrLf)
    End With

    ' Obtener y mostrar el valor de RangoCeldaSTFW
    Dim RangoCeldaSTFW As String
    RangoCeldaSTFW = ThisWorkbook.Sheets("hojaConfiguracion").Range("B16").Value
    With Me.txtDestinoSTWF
        .Value = RangoCeldaSTFW
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
