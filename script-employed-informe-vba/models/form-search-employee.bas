Private Sub buscarBoton_Click()
    If nameText.Text = "" Or diaBox.Text = "" Or mesBox.Text = "" Then
        MsgBox "Por favor, rellene el formulario completo para continuar", vbExclamation
        Exit Sub
    End If

    Dim fecha As Date
    Dim fechaFin As Date
    
    fecha = DateSerial(CInt(yearBox.Text), CInt(mesBox.Text), CInt(diaBox.Text))
    fechaFin = DateSerial(CInt(yearBoxfin.Text), CInt(mesBoxfin.Text), CInt(diaBoxfin.Text))
    
    copiar_pegar nameText.Text, fecha, fechaFin
   
    Unload Me
    
End Sub


Private Sub UserForm_Initialize()
   Dim i As Integer
    For i = 1 To 31
        diaBox.AddItem i
        diaBoxfin.AddItem i
    Next i
    
    
    Dim j As Integer
    For j = 1 To 12
        mesBox.AddItem j
        mesBoxfin.AddItem j
    Next j
    
    yearBox.AddItem 2023
    yearBoxfin.AddItem 2023
    yearBoxfin.AddItem 2024
    yearBox.AddItem 2024
    
End Sub

