Sub copiar_pegar(nombreBusqueda As String, fechaInicio As Date, fechaFin As Date)
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim filaDestino As Long
    Dim columnaBusqueda As Long
    Dim columnaBusquedaFecha As Long
    Dim filaBoton As Long
    Dim datos As Variant
    Dim resultado() As Variant
    Dim contador As Long
    Dim valorBusqueda As String

    Set wsOrigen = ThisWorkbook.Sheets("OT 2023")
    
    columnaBusqueda = 8
    columnaBusquedaFecha = 3
    filaBoton = 12
     
    valorBusqueda = Format(fechaBusqueda, "mm/dd/yyyy")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, columnaBusqueda).End(xlUp).Row

   
    datos = wsOrigen.Range("A1:I" & ultimaFila).Value


    ReDim resultado(1 To UBound(datos, 1), 1 To UBound(datos, 2))


    contador = 0


    For i = 1 To UBound(datos, 1)
        If i = filaBoton Then GoTo Siguiente
        
        If datos(i, columnaBusqueda) = nombreBusqueda Then
            If CDate(datos(i, columnaBusquedaFecha)) >= fechaInicio And CDate(datos(i, columnaBusquedaFecha)) <= fechaFin Then
                contador = contador + 1
                For j = 1 To UBound(datos, 2)
                    resultado(contador, j) = datos(i, j)
                Next j
            End If
        End If

Siguiente:
    Next i

    If contador > 0 Then
        Set wbDestino = Workbooks.Add
        Set wsDestino = wbDestino.Sheets(1)
        wsDestino.name = nombreBusqueda
        
        With wsDestino
            .Range("A1:I1").Value = Array("OT", "Cliente", "FECHA", "HORAS", "MNS", "VALOR", "Maq.", "Operario", "LUGAR OPERACIÓN")
            .Rows("1:1").Font.Bold = True
            .Rows("1:1").Interior.Color = RGB(255, 255, 0)
            .Rows("1:1").Font.Color = RGB(165, 42, 42)
        End With
        
        wsDestino.Range("A2").Resize(contador, UBound(datos, 2)).Value = resultado
        
        wsDestino.Columns("A:I").AutoFit
        
        Set chartRangeX = wsDestino.Range("B2:B" & contador + 1)
        Set chartRangeY = wsDestino.Range("F2:F" & contador + 1)

        Set chartObj = wsDestino.ChartObjects.Add(Left:=wsDestino.Columns("K").Left, Width:=700, Top:=wsDestino.Rows(2).Top, Height:=500)

        With chartObj.Chart
            .ChartType = xlColumnClustered
            .SetSourceData Source:=chartRangeY
            .SeriesCollection(1).XValues = chartRangeX
            .HasTitle = True
            .ChartTitle.Text = "Distribución de Valor por Cliente"
            .Legend.Delete
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Cliente"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Valor"
        End With


        MsgBox "Filas copiadas con éxito."
    Else
        MsgBox "No se encontraron filas con el valor de búsqueda.", vbInformation
        Debug.Print "Buscando nombre: " & nombreBusqueda & " y fecha: " & valorBusqueda
    End If
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub mostrarUser()
    UserForm1.Show
End Sub
