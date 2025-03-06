Sub crearexcel()
    Dim nombreArchivo As String
    Dim rutaArchivo As String

    nombreArchivo = ThisWorkbook.Sheets("Hoja1").Range("E24").Value

    If nombreArchivo = "" Then
        MsgBox "Por favor, ingresa un nombre en la celda E24.", vbExclamation
        Exit Sub
    End If
    
    rutaDocumentos = Environ("USERPROFILE") & "\Documents\"
    
    rutaArchivo = rutaDocumentos & nombreArchivo & ".xlsx"

    
    Dim nuevoLibro As Workbook
    Set nuevoLibro = Workbooks.Add
    
    ' Guarda el nuevo libro con el nombre especificado
    nuevoLibro.SaveAs Filename:=rutaArchivo, FileFormat:=xlOpenXMLWorkbook
    
    ' Cierra el nuevo libro
    nuevoLibro.Close SaveChanges:=False
    
    MsgBox "El archivo " & nombreArchivo & ".xlsx ha sido creado exitosamente.", vbInformation

End Sub
