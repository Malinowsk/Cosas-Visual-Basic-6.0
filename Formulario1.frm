'Resultado As Integer
Dim Contador As Integer

'ACCION CUANDO RECIBE UN CLICK
Private Sub BUTTON_Click()
    MsgBox (Me.Text1.Text)
    'MsgBox ("Bienvenido a Visual Basic 6.0")
End Sub
'ACCION CUANDO RECIBE EL FOCO
Private Sub Text1_GotFocus()
    Contador = Contador + 1
    MsgBox ("El primer botón ha recibido el enfoque" + Str(Contador))
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
'antes que se carge el formulario
Private Sub Form_Load()
    Contador = 0
    Me.BUTTON.Caption = "verific"
    Me.Text1.Text = "123"
End Sub

'sub se usa para procedimiento - para funcion es function
Public Function suma() As Integer
    suma = 2 + 2
End Function


