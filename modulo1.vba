Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
'declare sleep
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public localizacao As String
Public lote As String
Public bit As Integer
Public tipoinventario As String

Sub iniciar()
    bit = 0
    UserForm1.Show
    tipoinventario = ""
    localizacao = ""
    UserForm1.LabelLocal.Caption = "Localização atual: " + localizacao
    
End Sub

Sub SetMouse()

    'UserForm1.tbVersao.Clear
    UserForm1.TbNSaco.Clear
    
    lin = 1
    
    Do Until Sheets("p1").Cells(lin, 1) = ""

    UserForm1.TbNSaco.AddItem Sheets("p1").Cells(lin, 1)
    lin = lin + 1
    Loop
    'lin = 1
    'Do Until Sheets("p1").Cells(lin, 2) = ""
    '
    'UserForm1.tbVersao.AddItem Sheets("p1").Cells(lin, 2)
    'lin = lin + 1
    'Loop

    'lote2 = Right(UserForm1.TextBox1, 12)
    'lote = Left(lote2, 8)
    
    a1 = Right(UserForm1.TextBox1, 5)
    a2 = Left(a1, 1)
    UCase (a2)
    
    If (a2 = "A" Or a2 = "B" Or a2 = "C" Or a2 = "D" Or a2 = "E" Or a2 = "F" Or a2 = "a" Or a2 = "b" Or a2 = "c" Or a2 = "d" Or a2 = "e" Or a2 = "f") Then
        lote2 = Right(UserForm1.TextBox1, 13)
        lote2 = Left(lote2, 9)
    Else
        lote2 = Right(UserForm1.TextBox1, 12)
        lote2 = Left(lote2, 8)
    End If
    
    'Posicionar no icone do MMS
    SetCursorPos 260, 1041
    Call LeftClick
    'Posicionar na opção Ordens
    Sleep 100
    SetCursorPos 20, 75
    Call LeftClick
    'posicionar em OrderId
    SetCursorPos 75, 290
    Call LeftClick
    Application.SendKeys ("{DEL}")
    ''Retorna para TboxId
    'SetCursorPos 292, 280
    Call LeftClick
    Application.SendKeys lote2
    SetCursorPos 150, 120
    Call LeftClick
    Sleep 400
    'copiar
    SetCursorPos 292, 330
    Call RightClick
    Sleep 400
    SetCursorPos 350, 800
    Call LeftClick
    Sleep 100
    SetCursorPos 310, 1041
    Call LeftClick
    Sleep 150
    
End Sub

Private Sub LeftClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Sleep 50
End Sub

Private Sub RightClick()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 50
End Sub



