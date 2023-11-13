Attribute VB_Name = "Module6"
Dim ArchivoDestinoPath As String ' Variable global para almacenar la ruta del archivo origen
Dim archivoOrigenPath As String

Sub ActualizarTABRateCalc()
    Dim valores(1 To 7, 1 To 4) As Double
    
    
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
    
    
    
    ' Abre el archivo de origen seleccionado (usa la ruta almacenada)
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    ' Define la hoja de cálculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("Income Statement")

    ' Resto del código para obtener y colocar valores aquí...
    ' Obtiene los valores de Q1, Q2, Q3 y Q4 directamente desde las celdas
    valores(1, 1) = hojaOrigen.Range("T10").Value
    valores(1, 2) = hojaOrigen.Range("U10").Value
    valores(1, 3) = hojaOrigen.Range("V10").Value
    valores(1, 4) = hojaOrigen.Range("W10").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(2, 1) = hojaOrigen.Range("T11").Value
    valores(2, 2) = hojaOrigen.Range("U11").Value
    valores(2, 3) = hojaOrigen.Range("V11").Value
    valores(2, 4) = hojaOrigen.Range("W11").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(3, 1) = hojaOrigen.Range("T14").Value
    valores(3, 2) = hojaOrigen.Range("U14").Value
    valores(3, 3) = hojaOrigen.Range("V14").Value
    valores(3, 4) = hojaOrigen.Range("W14").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(4, 1) = hojaOrigen.Range("T15").Value
    valores(4, 2) = hojaOrigen.Range("U15").Value
    valores(4, 3) = hojaOrigen.Range("V15").Value
    valores(4, 4) = hojaOrigen.Range("W15").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(5, 1) = hojaOrigen.Range("T16").Value
    valores(5, 2) = hojaOrigen.Range("U16").Value
    valores(5, 3) = hojaOrigen.Range("V16").Value
    valores(5, 4) = hojaOrigen.Range("W16").Value

    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(6, 1) = hojaOrigen.Range("T23").Value
    valores(6, 2) = hojaOrigen.Range("U23").Value
    valores(6, 3) = hojaOrigen.Range("V23").Value
    valores(6, 4) = hojaOrigen.Range("W23").Value
    
    ' Obtiene los valores de Q1, Q2, Q3 Y Q4 directamente desde las celdas'
    valores(7, 1) = hojaOrigen.Range("T12").Value
    valores(7, 2) = hojaOrigen.Range("U12").Value
    valores(7, 3) = hojaOrigen.Range("V12").Value
    valores(7, 4) = hojaOrigen.Range("W12").Value
    ' Cierra el archivo de origen sin guardar cambios
    ' archivoOrigen.Close SaveChanges:=False

    ' Coloca los valores obtenidos en celdas específicas de tu hoja de cálculo principal
    For I = 1 To 7 ' Cambia el límite para i a 6
        For j = 1 To 4 ' Cambia el límite para j a 4
            ArchivoDestino.Sheets("Rate Calculation").Cells(I + 2, j + 31).Value = valores(I, j)
        Next j
    Next I
    archivoOrigen.Close SaveChanges:=False
End Sub







