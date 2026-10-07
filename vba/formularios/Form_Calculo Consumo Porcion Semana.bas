' ==========================================================
' Modulo  : Form_Calculo Consumo Porcion Semana
' Tipo    : 100  |  Lineas: 101
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:08
' ==========================================================

Option Compare Database

Private Sub Comando119_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery

End Sub

Public Sub Comando170_Click()

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim cotiz, cotiporcion As Integer
Dim Cal As Double
Dim CalMax As Double

miSQL = "SELECT Cotizaciones.CodCotizacion, Cotizaciones.Marca" & _
" FROM Cotizaciones" & _
" WHERE Cotizaciones.Marca = 'Almacen Global' ORDER BY Cotizaciones.CodCotizacion"

Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

For I = 1 To Cant
    If I >= Me.icoti Then
        cotiz = rst("CodCotizacion")
        cotiporcion = PorcionDeCotizacion(rst("CodCotizacion"))
        Cal = consumoporciones(cotiporcion, Me.asemana)
        CalMax = consumomaximoporciones(cotiporcion, Me.asemana, 4)
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalConsumoPorcionSemanal(CodCotizacion, Consumo, Semana, FechaPublicada, ConsumoMaximo)" & _
        " values (" & cotiz & ", " & Cal & ", " & Me.asemana & ", #" & Now & "#, " & CalMax & ")"
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
Dim cotiz, cotiporcion As Integer
Dim Cal As Double
Dim CalMax As Double
cotiz = Me.CodCotizacion
cotiporcion = PorcionDeCotizacion(Me.CodCotizacion)
Cal = consumoporciones(cotiporcion, Me.asemana)
CalMax = consumomaximoporciones(cotiporcion, Me.asemana, 4)
DoCmd.RunSQL "INSERT INTO CalConsumoPorcionSemanal(CodCotizacion, Consumo, Semana, FechaPublicada, ConsumoMaximo)" & _
" values (" & cotiz & ", " & Cal & ", " & Me.asemana & ", #" & Now & "#, " & CalMax & ")"
MsgBox "Valor Recalculado a " & Cal
Me.Consumo.Requery
End Sub

Private Sub Comando206_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
Me.Filter = "Consumo=0"
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

