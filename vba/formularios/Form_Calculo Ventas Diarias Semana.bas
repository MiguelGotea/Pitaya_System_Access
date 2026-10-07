' ==========================================================
' Modulo  : Form_Calculo Ventas Diarias Semana
' Tipo    : 100
' Lineas  : 101
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Private Sub Comando119_Click()
Me.Filter = "semana = " & Me.asemana
Me.FilterOn = True
Me.Requery
Me.calculado = VentaRealSemanal(Me.asemana)

End Sub

Public Sub Comando170_Click()

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim montpos As Double
Dim montnopos As Double
Dim Cantidad As Double
Dim fech As Date
Dim grup As Long


miSQL = "SELECT Grupos.CodGrupo, FechaSistema.Dates, numerosemana([FechaSistema]![Dates]) AS semana" & _
" FROM Grupos, FechaSistema" & _
" WHERE (((numerosemana([FechaSistema]![Dates]))=" & Me.asemana & ")) ORDER BY Grupos.CodGrupo, FechaSistema.Dates"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst


For I = 1 To Cant
    If I >= Me.icoti Then
        fech = rst("Dates")
        grup = rst("CodGrupo")
        montnopos = VentaxGrupoDia(fech, grup, 2)
        montpos = VentaxGrupoDia(fech, grup, 1)
        Cantidad = VentaxGrupoDia(fech, grup, 0)
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalVentasDiariasxGrupo(CodGrupo, Fecha, Cantidad, MontoNoPOS, MontoPOS, FechaPublicada) values (" & grup & _
        ", #" & fech & "#, " & Cantidad & ", " & montnopos & ", " & montpos & ", #" & Now & "#)"
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
Dim Cant As Double
Dim montpos As Double
Dim montnopos As Double
Cant = VentaxGrupoDia(Me.Dates, Me.CodGrupo, 0)
montpos = VentaxGrupoDia(Me.Dates, Me.CodGrupo, 1)
montnopos = VentaxGrupoDia(Me.Dates, Me.CodGrupo, 2)
DoCmd.RunSQL "INSERT INTO CalVentasDiariasxGrupo(CodGrupo, Fecha, Cantidad, MontoNoPOS, MontoPOS, FechaPublicada) values (" & Me.CodGrupo & _
", #" & Me.Dates & "#, " & Cant & ", " & montnopos & ", " & montpos & ", #" & Now & "#)"
MsgBox "Valor Recalculado a cantidad: " & Cant & " y montoPOS: " & montpos & " y montoNoPOS: " & montnopos
Me.Cantidad.Requery
Me.MontoPOS.Requery
Me.MontoNoPOS.Requery
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

Me.Filter = "Dates = 0" 'No filtra nada
Me.FilterOn = True
Me.Requery

End Sub
