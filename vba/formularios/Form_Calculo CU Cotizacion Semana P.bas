' ==========================================================
' Modulo  : Form_Calculo CU Cotizacion Semana P
' Tipo    : 100
' Lineas  : 97
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub Comando119_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
End Sub

Private Sub Comando170_Click()

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim cot As Integer
Dim ING As String
Dim Cal As Double
Dim valo As Double

miSQL = "SELECT Cotizaciones.CodCotizacion, DBIngredientes.TIPO1, Cotizaciones.Subproducto, Cotizaciones.CodIngrediente" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (DBIngredientes.TIPO1='VARIABLES' AND Cotizaciones.Subproducto<>0) ORDER BY Cotizaciones.CodCotizacion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

For I = 1 To Cant
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        ING = rst("CodIngrediente")
        Cal = FRCostoUnitarioIngrediente(ING, Me.asemana) * FRConversionCotizacion(cot, Me.asemana)
        valo = Cal
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalCUCotiSemanal(CodCotizacion, CU, Valor, Semana, FechaPublicada) values (" & cot & _
        ", " & Cal & ", " & valo & ", " & Me.asemana & ", #" & Now & "#)"
        DoCmd.SetWarnings True
        '/ Barra de carga
        Me.carga = I / Cant
        Me.carga.Left = Me.InsideWidth * I / Cant
        Me.barracarga.width = Me.InsideWidth * I / Cant
    End If
    rst.MoveNext
Next I

rst.Close
'/reseteando barra de carga
Me.barracarga.width = 0
Me.carga.Left = 0
Me.icoti = 1
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call Comando170_Click
End Sub

Private Sub Comando201_Click()
Dim Cal As Double
Dim valo As Double
Cal = FRCostoUnitarioIngrediente(Me.CodIngrediente, Me.asemana) * FRConversionCotizacion(Me.CodCotizacion, Me.asemana)
valo = Cal
DoCmd.RunSQL "INSERT INTO CalCUCotiSemanal(CodCotizacion, CU, Valor, Semana, FechaPublicada) values (" & Me.CodCotizacion & _
", " & Cal & ", " & valo & ", " & Me.asemana & ", #" & Now & "#)"
MsgBox "Valor Recalculado a " & Cal
Me.CU.Requery
End Sub

Private Sub Comando206_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
Me.Filter = "CU = 0"
Me.FilterOn = True
Me.Requery
End Sub

Private Sub Comando230_Click()
Dim sem As Integer
For sem = Me.sdesde To Me.shasta
    Me.asemana = sem
    Call Comando170_Click
Next sem
MsgBox "Autoguardado Completo"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.sactual = numerosemana(Date)
Me.sdesde = numerosemana(Date) - 1
Me.shasta = numerosemana(Date) - 1
End Sub

