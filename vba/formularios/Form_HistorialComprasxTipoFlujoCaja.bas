' ==========================================================
' Modulo  : Form_HistorialComprasxTipoFlujoCaja
' Tipo    : 100  |  Lineas: 38
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Private Sub Comando562_Click()
Dim tip1, tip2 As String
Dim tip0 As String
Dim cloca As Integer
Dim mesi, ano As Integer
tip0 = Me.Tipo
tip1 = Me.TIPO1
tip2 = Me.TIPO2
mesi = Me.ames
ano = Me.aano
cloca = Me.Local

DoCmd.OpenForm "HistorialComprasGobalDetalle", acNormal
Forms("HistorialComprasGobalDetalle").Filter = "[local]= " & cloca & " and [ctipo]= '" & tip0 & "' and [TIPO1]= '" & tip1 & "' and [TIPO2]= '" & tip2 & "' and [ames]= " & mesi & " and [aano]= " & ano
'"[Tipo]= '" & tip0 & "' and [TIPO1]= '" & tip1 & "' and [TIPO2]= '" & tip2 & "' and [ames]= " & mesi & " and [aano]= " & ano

Forms("HistorialComprasGobalDetalle").FilterOn = True

End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.InsideHeight = 8000
'Deifnir origen de base de datos base de datos mixed
Me.RecordSource = "SELECT Month([Compras]![Fecha]) AS ames, Year([Compras]![Fecha]) AS aano," & _
" DBIngredientes.TIPO1, DBIngredientes.TIPO2, DBIngredientes.Tipo," & _
" Sum(Compras.CostoTotal) AS SumaDeCostoTotal, Compras.local" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN Compras ON Cotizaciones.CodCotizacion = Compras.CodCotizacion" & _
" IN 'C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb'" & _
" GROUP BY Month([Compras]![Fecha]), Year([Compras]![Fecha]), DBIngredientes.TIPO1," & _
" DBIngredientes.TIPO2, DBIngredientes.Tipo, Compras.local" & _
" ORDER BY DBIngredientes.TIPO1, DBIngredientes.TIPO2, DBIngredientes.Tipo"

End Sub
