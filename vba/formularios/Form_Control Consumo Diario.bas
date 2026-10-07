' ==========================================================
' Modulo  : Form_Control Consumo Diario
' Tipo    : 100
' Lineas  : 49
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database

Private Sub Comando155_Click()

Dim sin As Double
Dim ING As Double
Dim mer As Double
Dim cre As Double
Dim ingres As String
Dim semis As Integer
Dim nami As String

ingres = Me.aproducto
nami = DLookup("[Nombre]", "[DBIngredientes]", "[CodIngrediente]='" & ingres & "'")
semis = numerosemana(Date)
sin = StockIngrediente(ingres, semis) + StockCotizacion(ingres, semis)
ING = IngresoTotalIngrediente(ingres, semis)
mer = MermaIngrediente(ingres, semis) + MermaCotizacion(ingres, semis)
cre = ConsumoSemanalProducto(ingres, semis)

DoCmd.OpenReport "ControlInventarioDiario", acViewReport

[Reports]![ControlInventarioDiario].Report.Printer.LeftMargin = 0
[Reports]![ControlInventarioDiario].Report.Printer.RightMargin = 0
[Reports]![ControlInventarioDiario].Report.Printer.BottomMargin = 0
[Reports]![ControlInventarioDiario].Report.Printer.TopMargin = 0


[Reports]![ControlInventarioDiario]![CodIngrediente] = ingres
[Reports]![ControlInventarioDiario]![Nombre] = nami
[Reports]![ControlInventarioDiario]![asemana] = semis

[Reports]![ControlInventarioDiario]![SI] = sin
[Reports]![ControlInventarioDiario]![IN] = ING
[Reports]![ControlInventarioDiario]![ME] = mer
[Reports]![ControlInventarioDiario]![CR] = cre
[Reports]![ControlInventarioDiario]![SFR] = sin + ING - mer - cre

DoCmd.PrintOut , 1, 1
DoCmd.Close acReport, "ControlInventarioDiario"
End Sub

Private Sub Comando56_Click()
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
End Sub
