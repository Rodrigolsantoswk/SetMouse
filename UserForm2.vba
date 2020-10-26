
Private Sub ToggleButton1_Click()
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(1, 0) = TextBox1
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 1) = TextBox3
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 2) = "Sem Etiqueta"
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 3) = TextBox4
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 4) = "Sem Etiqueta"
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 5) = "Sem Etiqueta"
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 6) = "Sem Etiqueta"
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 7) = "Sem Etiqueta"
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 8) = TextBox2
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 9) = Now
    Workbooks("Inventario sala de quimicos.xlsm").Sheets("base").Range("A1048576").End(xlUp).Offset(0, 10) = localizacao
    UserForm2.Hide
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    UserForm1.TextBox1.SetFocus
End Sub
