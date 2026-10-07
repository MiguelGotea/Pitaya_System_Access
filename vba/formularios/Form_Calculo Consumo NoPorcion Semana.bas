' ==========================================================
' Modulo  : Form_Calculo Consumo NoPorcion Semana
' Tipo    : 100
' Lineas  : 90
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database

Private Sub Comando119_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery

'Me.calculado = DSum("VentasXTipoNoPorcionSemana(" & Me.asemana & ",[CodIngrediente])", "DBIngredientes", "[TIPO1]='VARIABLES'")

End Sub

Public Sub Comando170_Click()

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim ingrex As String
Dim Cal As Double
Dim CalMax As Double

miSQL = "SELECT DBIngredientes.CodIngrediente, DBIngredientes.TIPO1, DBIngredientes.Vigente FROM DBIngredientes" & _
" WHERE (DBIngredientes.TIPO1='VARIABLES' AND DBIngredientes.Vigente<>0) ORDER BY DBIngredientes.CodIngrediente"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

For I = 1 To Cant
    If I >= Me.icoti Then
        ingrex = rst("CodIngrediente")
        Cal = VentasXTipoNoPorcionSemana(Me.asemana, ingrex)
        CalMax = VentasXTipoNoPorcionMaximoSemana(Me.asemana, ingrex, 4)
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalConsumoNoPorcionSemanal(CodIngrediente, Consumo, Semana, FechaPublicada, ConsumoMaximo)" & _
        " values ('" & ingrex & "', " & Cal & ", " & Me.asemana & ", #" & Now & "#, " & CalMax & ")"
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
Dim CalMax As Double
Cal = VentasXTipoNoPorcionSemana(Me.asemana, Me.CodIngrediente)
CalMax = VentasXTipoNoPorcionMaximoSemana(Me.asemana, Me.CodIngrediente, 4)
DoCmd.RunSQL "INSERT INTO CalConsumoNoPorcionSemanal(CodIngrediente, Consumo, Semana, FechaPublicada, ConsumoMaximo)" & _
" values ('" & Me.CodIngrediente & "', " & Cal & ", " & Me.asemana & ", #" & Now & "#, " & CalMax & ")"
MsgBox "Valor Recalculado a " & Cal
Me.Consumo.Requery
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

