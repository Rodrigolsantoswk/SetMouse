Dim Today

Private Sub bCancelar_Click()
    UserForm1.Height = 109
    tbCompound = ""
    tbBoxId = ""
    tbScale = ""
    tbMixer = ""
    tbTime = ""
    tbOperator = ""
    tbPeso = ""
    tbVersao = ""
    TextBox1 = ""
    TextBox1.Enabled = True
    TextBox1.SetFocus
    ToggleButton1.Enabled = True
End Sub

Private Sub TextBox1_AfterUpdate()
    On Error GoTo erro
    LabelErro = ""
    Workbooks("Inventario compostos.xlsm").Sheets("aux").Range("A1:Z1").Clear
    If TextBox1 <> "" Then
        If TextBox1 = "" Then
            
        End If
    End If
    
erro:
    MsgBox "Erro de execução: " & Err.Description
    TextBox1 = ""
    TextBox1.SetFocus
    
End Sub


Private Sub ToggleButton1_Click()


    If localizacao <> "" Then
        UserForm2.TextBox3.Clear
        UserForm2.TextBox4.Clear
        
        lin = 1
        
        Do Until Sheets("p1").Cells(lin, 1) = ""
        
        UserForm2.TextBox4.AddItem Sheets("p1").Cells(lin, 1)
        lin = lin + 1
        Loop
        lin = 1
        Do Until Sheets("p1").Cells(lin, 2) = ""
        
        UserForm2.TextBox3.AddItem Sheets("p1").Cells(lin, 2)
        lin = lin + 1
        Loop
        UserForm2.Show
       
    Else
        LabelErro = "Bipe a localização primeiro."
    End If
End Sub

Private Sub ToggleButton3_Click()
    
    
    
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(1, 0) = tbCompound
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 1) = tbVersao
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 2) = tbBoxId
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 3) = TbNSaco
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 4) = tbScale
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 5) = tbMixer
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 6) = tbTime
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 7) = tbOperator
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 8) = tbPeso
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 9) = Now
    Workbooks("Inventario compostos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 10) = localizacao
                
    
    
    LabelAviso = "Última caixa validada: " + Right(TextBox1, 12)
    
    UserForm1.Height = 109
    ToggleButton1.Enabled = True
    tbCompound = ""
    tbBoxId = ""
    tbScale = ""
    tbMixer = ""
    tbTime = ""
    tbOperator = ""
    tbPeso = ""
    tbVersao = ""
    TbNSaco = ""
    
    TextBox1.Enabled = True
    ToggleButton1.Enabled = True
    TextBox1 = ""
    
    TextBox1.SetFocus
    
    'Application.SendKeys ("{TAB}")
End Sub

Private Sub UserForm_Click()

End Sub