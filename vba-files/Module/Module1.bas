Attribute VB_Name = "Module1"
Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
Dim archivoOrigenPath As String

Public RangoHojaOrigen() As String ' Variable para almacenar los rangos de la hoja de origen
Public RangoHojaDestino As String ' Variable para almacenar el rango de la hoja de destino

Sub ObtenerYColocarShifts()
    Dim ArchivoDestino As Workbook
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim turno As Integer
    
    archivoOrigenPath = "C:\Users\3762091\Desktop\trabajo\Proyecto Excel\Excel\NUEVOS\08 DL Breakdown FY24 Q1 BID 4 Rev F.xlsx"
    ' Verifica si ya se ha seleccionado un archivo de origen previamente
    If archivoOrigenPath = "" Then
        archivoOrigenPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de origen (DL Breakdown)")
        If archivoOrigenPath = "Falso" Then
            Exit Sub
        End If
    End If
    
    ' Verifica si ya se ha seleccionado un archivo de destino previamente
    If ArchivoDestinoPath = "" Then
        ArchivoDestinoPath = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de destino (BU Scenario Flexline)")
        If ArchivoDestinoPath = "Falso" Then
            Exit Sub
        End If
    End If
    
    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    ' Define la hoja de c?lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("IMED DL Breakdow")
    
    ' Definir los nombres de los turnos
    Dim nombresTurnos() As String
    nombresTurnos = Split("FirstShift,SecondShift,ThirdShift,FourTwentyShift,FourTwentyOneShift,FourTwentyTwoShift,FourTwentyThreeShift", ",")
    
    ' Redimensionar el arreglo para almacenar los rangos
    ReDim RangoHojaOrigen(1 To 7)
    
    ' Procesar cada turno
    For turno = 1 To 7
        Dim turnoNombre As String
        turnoNombre = nombresTurnos(turno - 1)
        
        ' Definir el rango de la hoja de origen
        RangoHojaOrigen(turno) = hojaOrigen.Range("S" & (45 + ((turno - 1) * 41)) & ":AD" & (81 + ((turno - 1) * 41))).Address
        
        ' Definir el rango de la hoja de destino
        RangoHojaDestino = ArchivoDestino.Sheets("Sheet1").Range("S" & (45 + ((turno - 1) * 41)) & ":AD" & (81 + ((turno - 1) * 41))).Address
        
        ' Asignar valores desde la hoja de origen a la hoja de destino
        ArchivoDestino.Sheets("Sheet1").Range(RangoHojaDestino).Value = hojaOrigen.Range(RangoHojaOrigen(turno)).Value
    Next turno
    
    ' Cerrar el archivo de origen sin guardar cambios
    archivoOrigen.Close SaveChanges:=False
End Sub

