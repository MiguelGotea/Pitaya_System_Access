' ==========================================================
' Modulo  : Form_Costeo Recetas
' Tipo    : 100
' Lineas  : 44
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub

Private Sub ing1_Change()
Me.uni1.Requery

End Sub
Private Sub ing2_Change()
Me.uni2.Requery
End Sub
Private Sub ing3_Change()
Me.uni3.Requery
End Sub
Private Sub ing4_Change()
Me.uni4.Requery
End Sub
Private Sub ing5_Change()
Me.uni5.Requery
End Sub
Private Sub ing6_Change()
Me.uni6.Requery
End Sub
Private Sub ing7_Change()
Me.uni7.Requery
End Sub
Private Sub top1_Change()
Me.unit1.Requery
End Sub
Private Sub top2_Change()
Me.unit2.Requery
End Sub
Private Sub top3_Change()
Me.unit3.Requery
End Sub
Private Sub top4_Change()
Me.unit4.Requery
End Sub
Private Sub top5_Change()
Me.unit5.Requery
End Sub
