' ==========================================================
' Modulo  : Form_Calculo Conversion Cotizacion Semana
' Tipo    : 100
' Lineas  : 128
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Public Sub Comando119_Click()
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
Dim cot As Integer
Dim Cal As Double

miSQL = "SELECT Cotizaciones.CodCotizacion, DBIngredientes.TIPO1" & _
" FROM DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente" & _
" WHERE (DBIngredientes.TIPO1='VARIABLES') ORDER BY Cotizaciones.CodCotizacion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

For I = 1 To Cant
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        Cal = conversioncalculado(cot, Me.asemana)
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalConversionSemanal(CodCotizacion, Conversion, Semana, FechaPublicada) values (" & cot & _
        ", " & Cal & ", " & Me.asemana & ", #" & Now & "#)"
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
Cal = conversioncalculado(Me.CodCotizacion, Me.asemana)
DoCmd.RunSQL "INSERT INTO CalConversionSemanal(CodCotizacion, Conversion, Semana, FechaPublicada) values (" & Me.CodCotizacion & _
", " & Cal & ", " & Me.asemana & ", #" & Now & "#)"
MsgBox "Valor Recalculado a " & Cal
Me.Conversion.Requery
End Sub

Private Sub Comando206_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
Me.Filter = "Conversion = 0"
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

Private Sub Comando245_Click()
Me.Filter = "(Tipo = 'Frutas'  or Tipo = 'Verduras') And (Marca <> 'Almacen Global' or IsNull([Marca])) And Conversion = 0"
Me.FilterOn = True
Me.Requery
End Sub

Public Sub Comando437_Click()
Me.Filter = "Natural = 0" '"iif([Conversionanterior]=0,0,([Conversion]-[Conversionanterior])/[Conversionanterior]) <> 0 And Natural = 0"
Me.FilterOn = True
Me.Requery

Dim Direccion, Direccion2 As String
Dim archivo As String

Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.asemana
Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & Me.asemana & " - 5.Rendimientos Insumos.pdf"
DoCmd.OutputTo acOutputForm, "Calculo Conversion Cotizacion Semana", acFormatPDF, archivo, False
End Sub

Private Sub Form_Close()
Me.Filter = ""
Me.FilterOn = True
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.sactual = numerosemana(Date)
Me.sdesde = numerosemana(Date) - 1
Me.shasta = numerosemana(Date) - 1

End Sub

