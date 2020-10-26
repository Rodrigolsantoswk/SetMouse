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
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1:Z1").Clear
    If TextBox1 <> "" Then
        If localizacao <> "" Then
            If TextBox1 = "Terreo" Or TextBox1 = "Primeiro Piso" Then
                localizacao = TextBox1
                LabelTitulo = "Bipe os pallets"
                LabelLocal.Caption = "Localização atual: " + localizacao
                LabelAviso = "Localização alterada para: " + localizacao
                TextBox1 = ""
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
            ElseIf UCase(Left(TextBox1, 5)) = "9241H" Then
                Call SetMouse
                 
                Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Paste Destination:=Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1048576").End(xlUp).Offset(0, 0)
                Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1048576").End(xlUp).Offset(0, 16) = localizacao
                Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1048576").End(xlUp).Offset(0, 17) = Now
                Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Rows(Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1048576").End(xlUp).Row - 1).EntireRow.Delete
               
                UserForm1.Height = 254
                tbCompound = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 1)
                'Isolando a versão
                aux = Right(tbCompound, 11)
                tbVersao = Left(aux, 4)
                
                a1 = Right(UserForm1.TextBox1, 5)
                a2 = Left(a1, 1)
                UCase (a2)
                
                If (a2 = "A" Or a2 = "B" Or a2 = "C" Or a2 = "D" Or a2 = "E" Or a2 = "F") Then
                    tbBoxId = Right(UserForm1.TextBox1, 13)
                    
                Else
                    tbBoxId = Right(UserForm1.TextBox1, 12)
                End If
                
                tbScale = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 4)
                tbMixer = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 3)
                tbTime = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 13)
                tbOperator = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 10)
                tbPeso = Workbooks("Inventario sala de quimicos.xlsm").Sheets("aux").Range("A1").Offset(0, 6)
                TextBox1.Enabled = False
                ToggleButton1.Enabled = False
                Application.SendKeys ("{TAB}")
                Application.SendKeys ("{TAB}")
                Application.SendKeys ("{TAB}")
                Application.SendKeys ("{TAB}")
                Application.SendKeys ("{TAB}")
            Else
                LabelErro = "Bipe uma etiqueta de pigmentos."
                TextBox1 = ""
            End If
        Else
            If TextBox1 = "Primeiro Piso" Or TextBox1 = "Terreo" Then
                localizacao = TextBox1
                LabelTitulo = "Bipe os pallets"
                LabelLocal.Caption = "Localização atual: " + localizacao
                TextBox1 = ""
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")
                    Application.SendKeys ("{TAB}")

            ElseIf UCase(Left(TextBox1, 5)) = "9241H" Then
                LabelErro = "Bipe a localização primeiro."
                TextBox1 = ""
            Else
                LabelErro = "Bipe uma localização válida."
                TextBox1 = ""
            End If
        End If
    End If
    Exit Sub
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
