REM MODULO PARA AGREGAR CONTRASEÑA AL USERFORM

Public Sub AbrirFormulario()
Dim clave As String
Dim ClaveCorrecta As String

ClaveCorrecta = "ABC123"
clave = InputBox("Ingrese la Contraseña para acceder", "Acceso Restringido")
If clave = "" Then
Exit Sub
End If

If clave <> ClaveCorrecta Then
MsgBox "Contraseña Incorrecta"
Exit Sub
End If

UserForm1.Show

End Sub

------------------------------------------------------------------------
REM MODULO HECHO CON GRABADORA DE MACRO PARA BUSCAR EL VALOR "***" EN TODA UNA COLUMNA Y REMPLAZARLO POR OTRO VALOR "***"
Sub Reemplazar()
' Macro1 Macro

    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="E011", Replacement:="E015", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("T_VENTAS_26[[#Headers],[FECHA]]").Select
End Sub
