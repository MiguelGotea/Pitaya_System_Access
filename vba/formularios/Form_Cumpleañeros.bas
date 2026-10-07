' ==========================================================
' Modulo  : Form_Cumpleañeros
' Tipo    : 100  |  Lineas: 9
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
'Deifnir origen de base de datos base de datos mixed
Me.RecordSource = "SELECT Day([ClientesClub]![Cumpleanos]) AS Dia, Month([ClientesClub]![Cumpleanos]) AS Mes, ClientesClub.CodCliente, " & _
"[ClientesClub]![Nombre] & ' ' & [ClientesClub]![Apellidos] AS Nombre FROM ClientesClub IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb' " & _
"WHERE (((Month([ClientesClub]![Cumpleanos]))=Month(Date()))) ORDER BY Day([ClientesClub]![Cumpleanos])"
End Sub
