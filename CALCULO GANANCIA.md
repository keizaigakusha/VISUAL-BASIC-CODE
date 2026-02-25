Private Sub BtnCalcular_Click()

Rem Creando mi formulario de multiproducto

Rem Declarando variables
Dim precosto As Double
Dim ganancia As Double
Dim precioventa As Double

Rem pasando a precosto lo que se almacena en TxtPrecioCosto
precosto = Val(Me.TxtPrecioCosto.Text)

Rem realizando los calculos
ganancia = precosto * 0.2
precioventa = precosto + ganancia

Rem pasando los datos de ganancia a precioventa
Me.LblGanancia.Caption = ganancia
Me.LblPrecioVenta.Caption = precioventa

End Sub


<img width="554" height="383" alt="image" src="https://github.com/user-attachments/assets/22fa12d0-495a-42a8-8bfd-84722c34b636" />
