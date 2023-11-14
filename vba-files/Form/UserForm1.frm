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
Private Sub UserForm1_Initialize()
    txtOrigenSTFW.Text = RangoHojaOrigen()
End Sub

Private Sub CommandButton25_Click()
    Unload Me
    UserForm2.Show
End Sub

Private Sub Image1_Click()
    Unload Me
    UserForm3.Show
End Sub
