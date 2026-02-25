Option Explicit

Private Sub BtnNuevo_Click()
Me.TxtCodigo = ""
Me.TxtNomApe = ""
Me.CmbCarrera = ""
Me.CmbTurno = ""
Me.LblPension = ""
Me.TxtCodigo.SetFocus
End Sub
---------------------------------------------------------------------
Private Sub BtnRegistrar_Click()
Rem declarando variables
Dim ultfila As Integer
Dim fila As Integer
Dim duplicado As Boolean
Dim x As Integer

Rem Asumiendo el caso que no hay duplicados en la columna 1 "Codigo" no se repite
duplicado = False

Rem Busca la última fila usada en la columna A y dime el numero de la siguiente fila vacía
Rem Rows.Count te da el numero de filas que hay en excel "1,048,576"
Rem Rows.Count, 1 especifica que solo busque en la columna 1 donde estan los códigos
Rem .End(xlup) ordena que suba desde la fila "1,048,576" hasta la fila donde encuentre datos y sume +1 fila
Rem se suma +1 para que la variable  ultfila te de la fila que esta vacia que sigue a la fila que esta con el codigo
ultfila = Cells(Rows.Count, 1).End(xlUp).Row + 1

Rem cuenta las celdas que no estan vacias de la columna codigo de la tabla alumno y te da el numero
fila = Application.WorksheetFunction.CountA(Range("TablaAlumno[Codigo]"))

Rem Repite desde 1 hasta la cantidad de registros
For x = 1 To fila

Rem a la variable x se le suma + 1 porque la celda de la columna tiene encabezado de titulo
Rem aqui se empieza el condicional if
If Cells(x + 1, 1).Value = Me.TxtCodigo.Text Then
MsgBox "Dato duplicado en la FILA" & x
duplicado = True
End If

Next
'Si es verdadero existe duplicado y si es falso no existe duplicado
'entonces hay que guardar el registro
If Not duplicado Then
Cells(ultfila, 1).Value = Me.TxtCodigo.Text
Cells(ultfila, 2).Value = Me.TxtNomApe.Text
Cells(ultfila, 3).Value = Me.CmbCarrera.Text
Cells(ultfila, 4).Value = Me.CmbTurno.Text
Cells(ultfila, 5).Value = Me.LblPension.Caption
End If

End Sub
------------------------------------------------------------------------
Private Sub BtnSalir_Click()
End
End Sub
-----------------------------------------------------------------------
Private Sub CmbCarrera_Change()

Rem declarando variable
Dim Indice As Integer
Dim P As Single

Rem Aqui vamos a asignar el monto de la pensión dependiendo del orden del rango "Carrera"
Indice = Me.CmbCarrera.ListIndex
Select Case Indice
Case 0: P = 750
Case 1: P = 520
Case 2: P = 700
Case 3: P = 630
Case 4: P = 1200
Case 5: P = 900
End Select
Rem el valor de la variable P lo guardamos en la etiqueta "Pension"
Me.LblPension.Caption = P

End Sub
-----------------------------------------------------------------------------------------
Rem Asignando el rango valores al cuadro de lista combinado
Rem previamente ya habia asignado nombre de rango donde estan los valores de turno:"Tarde" y "Noche"
Rem lo mismo con el nombre de rango "Carrera"
Private Sub UserForm_Activate()
Me.CmbCarrera.RowSource = "Carrera"
Me.CmbTurno.RowSource = "Turno"

End Sub
