Private Sub CommandButtonRegistrar_Click()
Rem declarando variables
Dim ultfila As Integer
Dim fila As Integer

If Me.ComboBox1.Text = "" Then
MsgBox "Ingresar Tipo de Fecha Valido"
Exit Sub
End If

If Me.TextBox1.Text = "" Then
MsgBox "Ingresar Descripción de Fecha Valida"
Exit Sub
End If

If Me.TextBox2.Text = "" Then
MsgBox "Ingresar Descripción de Fecha Valida"
Exit Sub
End If

If Not IsDate(Me.TextBox2.Text) Then
MsgBox "Ingresar Fecha Valida"
Exit Sub
End If

ultfila = Cells(Rows.Count, 1).End(xlUp).Row + 1

Cells(ultfila, 1).Value = Me.ComboBox1.Text
Cells(ultfila, 2).Value = Me.TextBox1.Text
Cells(ultfila, 3).Value = Me.TextBox2.Text

Me.ComboBox1.Value = ""
Me.TextBox1.Value = ""
Me.TextBox2.Value = ""
Me.ComboBox1.SetFocus

MsgBox "¡¡¡Datos registrados con éxito!!!", vbInformation, "Resolución"
End Sub

Private Sub Userform_Activate()
Me.ComboBox1.RowSource = "Tipo"
End Sub
-----------------------------------------------------------------------------------
Private Sub TextBox1_Change()
Rem Insertando Modulo para filtro de busqueda
  Dim filtro As String
  filtro = "*" & Me.TextBox1.Text & "*"

  With Me.Range("B6").CurrentRegion
  .AutoFilter Field:=2, Criteria1:=filtro
  End With
End Sub

<img width="410" height="283" alt="image" src="https://github.com/user-attachments/assets/54dd963c-85eb-48e9-8f41-d0bdd9a3050d" />
