Sub generarplanilla()
    Dim wbNuevo As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim rango As Range
    Dim rutaArchivo As String
    Dim nuevaHojaNombre As String
    Dim hojaExiste As Boolean
    Dim contador As Integer

    ' Ruta del archivo creado previamente (Documentos)
    rutaArchivo = Environ("USERPROFILE") & "\Documents\" & ThisWorkbook.Sheets("Hoja1").Range("E24").Value & ".xlsx"

    ' Abre el archivo existente
    On Error Resume Next
    Set wbNuevo = Workbooks.Open(rutaArchivo)
    If wbNuevo Is Nothing Then
        MsgBox "No se pudo abrir el archivo especificado.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Establecer la hoja de origen
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")

    ' Selecciona el rango que se va a copiar
    Set rango = wsOrigen.Range("A1:H20")

  
    Set wsDestino = wbNuevo.Sheets.Add

   
    rango.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll

    
    nuevaHojaNombre = wsDestino.Range("A3").Value
    If nuevaHojaNombre = "" Then nuevaHojaNombre = "HojaSinNombre"

   
    hojaExiste = True
    contador = 1

    Do While hojaExiste
        hojaExiste = False
        For Each ws In wbNuevo.Sheets
            If ws.Name = nuevaHojaNombre Then
                hojaExiste = True
                nuevaHojaNombre = nuevaHojaNombre & contador
                contador = contador + 1
                Exit For
            End If
        Next ws
    Loop

    ' Renombrar la nueva hoja con el nombre especificado
    wsDestino.Name = nuevaHojaNombre

    ' Ajustar ancho de columnas y filas
    wsDestino.Columns("A:I").AutoFit
    wsDestino.Rows("1:100").AutoFit

    ' Insertar una nueva fila y copiar datos
    Dim filaOrigen As Long
    Dim filaDestino As Long
    filaOrigen = 1
    filaDestino = 6

    ' Inserta una nueva fila en la posición 5 y copia la fila 1 allí
    wsDestino.Rows(filaDestino).Insert Shift:=xlDown
    wsDestino.Rows(filaOrigen).Copy Destination:=wsDestino.Rows(filaDestino)

    ' Borra la fila original en la posición 1 para evitar duplicados
    wsDestino.Rows(filaOrigen).Delete


    Dim columnaNueva As Long
    Dim filaNombre As Long
    Dim nombreColumna As String
    columnaNueva = 3
    filaNombre = 5
    nombreColumna = "DTC"

    wsDestino.Columns(columnaNueva).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsDestino.Cells(filaNombre, columnaNueva).Value = nombreColumna

    Dim columnaNueva2 As Long
    Dim filaNombre2 As Long
    Dim nombreColumna2 As String
    columnaNueva2 = 4
    filaNombre2 = 5
    nombreColumna2 = "DV"

    wsDestino.Columns(columnaNueva2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsDestino.Cells(filaNombre2, columnaNueva2).Value = nombreColumna2

    ' Aplicar bordes a la tabla
    Dim rangoTabla As Range
    Set rangoTabla = wsDestino.Range("A5:J24")

    With rangoTabla.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' Añadir fórmulas
    For x = 7 To 24
        wsDestino.Range("E" & x).Formula = "=B" & x & "-C" & x & "-D" & x
    Next x
    
    
    wsDestino.Range("B1:E4").ClearContents

    Application.CutCopyMode = False

    
    wbNuevo.Save

    MsgBox "Datos copiados y modificados con éxito en el archivo existente."
End Sub

