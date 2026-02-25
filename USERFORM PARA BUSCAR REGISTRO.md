Option Explicit

Private Sub BtnFinalizar_Click()
End
End Sub

Private Sub LstDato_Click()
Me.LblCodigo.Caption = LstDato.List(LstDato.ListIndex, 0)
Me.LblNomApe.Caption = LstDato.List(LstDato.ListIndex, 1)
Me.LblCarrera.Caption = LstDato.List(LstDato.ListIndex, 2)
Me.LblTurno.Caption = LstDato.List(LstDato.ListIndex, 3)
Me.LblPension.Caption = LstDato.List(LstDato.ListIndex, 4)
End Sub

Private Sub TxtNomApe_Change()
Dim Lista1 As Range
Dim Lista2 As Range
Dim Lista3 As Range
Dim Lista4 As Range
Dim Lista5 As Range
Dim NomApe As Range
Dim reg As Long
Set Lista1 = Range("TablaAlumno[CODIGO]")
        Set Lista2 = Range("TablaAlumno[NOMBRES Y APELLIDOS]")
        Set Lista3 = Range("TablaAlumno[CARRERA]")
        Set Lista4 = Range("TablaAlumno[TURNO]")
        Set Lista5 = Range("TablaAlumno[PENSION]")
        Me.LstDato.Clear
        reg = 1
        For Each NomApe In Lista2.Cells
    If TxtNomApe.Text <> "" Then
        If VBA.LCase(NomApe.Value) Like "*" & VBA.LCase(TxtNomApe.Text) & "*" Then
            LstDato.AddItem Lista1(reg)
            LstDato.List(LstDato.ListCount - 1, 1) = Lista2(reg)
            LstDato.List(LstDato.ListCount - 1, 2) = Lista3(reg)
            LstDato.List(LstDato.ListCount - 1, 3) = Lista4(reg)
            LstDato.List(LstDato.ListCount - 1, 4) = Lista5(reg)
        End If
    End If
    reg = reg + 1
Next NomApe
End Sub

Private Sub UserForm_Activate()
Me.LstDato.ColumnWidths = "55 ; 90; 120 ; 40 ; 30"
End Sub


<img width="1154" height="326" alt="image" src="https://github.com/user-attachments/assets/2b9d5f72-3607-4e8f-8278-946c7878b120" />
<img width="763" height="411" alt="image" src="https://github.com/user-attachments/assets/88f1909d-594e-4d8f-a53d-4996cee86b8c" />
