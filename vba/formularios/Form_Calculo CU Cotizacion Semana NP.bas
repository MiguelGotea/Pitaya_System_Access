' ==========================================================
' Modulo  : Form_Calculo CU Cotizacion Semana NP
' Tipo    : 100
' Lineas  : 183
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:27
' ==========================================================
Option Compare Database

Public Sub Comando119_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
End Sub

Private Sub DescargarTablaCoti()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim cot As Integer
Dim Cal As Double
Dim sema As Integer
Dim val As Double

sema = Me.asemana
miSQL = "SELECT CodCotizacion, formula, valores FROM TempCUCotizaciones ORDER BY CodCotizacion"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

Me.proceso = "Copiando Datos"
For I = 1 To Cant
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        Cal = rst("formula")
        val = rst("valores")
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalCUCotiSemanal(CodCotizacion, CU, Valor, Semana, FechaPublicada) values (" & cot & _
        ", " & Cal & ", " & val & ", " & sema & ", #" & Now & "#)"
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
Call DescargarTablaCoti
End Sub

Private Sub CopiarTablaPrincipalCoti()

On Error GoTo AgainAgain2

If DCount("*", "MSysObjects", "Type=1 AND Name='TempCUCotizaciones'") > 0 Then
DoCmd.RunSQL "DROP TABLE TempCUCotizaciones"
End If

'Crear carpeta temporal
DoCmd.SetWarnings False
DoCmd.RunSQL "SELECT x.* INTO TempCUCotizaciones" & _
" FROM (SELECT DBIngredientes.TIPO1, Cotizaciones.Subproducto, TiposVariables.Control, DBIngredientes.Vigente," & _
" Cotizaciones.Descontinuado, Cotizaciones.CodCotizacion, CostoUnitarioCotizacionxCompra([Cotizaciones]![CodCotizacion]," & Me.asemana & ") AS formula," & _
" 0 AS valores" & _
" FROM (DBIngredientes INNER JOIN Cotizaciones ON DBIngredientes.CodIngrediente = Cotizaciones.CodIngrediente)" & _
" INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" WHERE (((DBIngredientes.TIPO1)='VARIABLES') AND ((Cotizaciones.Subproducto)=0)" & _
" AND ((TiposVariables.Control)<>0) AND ((DBIngredientes.Vigente)<>0) AND ((Cotizaciones.Descontinuado)=0)))x"
DoCmd.SetWarnings True

'" ValorCotizacionGlobal([Cotizaciones]![CodCotizacion]," & Me.asemana & ") AS valores"
'MsgBox "Datos Guaradados en tabla temporal"
Exit Sub

AgainAgain2:
Call CopiarTablaPrincipalCoti
End Sub
Private Sub Comando170_Click()
Me.icoti = 1
Me.proceso = "Guardando Tabla"
Call CopiarTablaPrincipalCoti
Me.proceso = "Tabla temporal de semana " & Me.asemana & " creada"
Call DescargarTablaCoti
Me.icoti = Null
Me.proceso = Null
MsgBox "Autoguardado Completo"
End Sub

Private Sub Comando201_Click()
Dim Cal As Double
Dim val As Double

Cal = CostoUnitarioUnidad(Me.CodCotizacion, Me.asemana)
val = ValorCotizacionGlobal(Me.CodCotizacion, Me.asemana)
DoCmd.RunSQL "INSERT INTO CalCUCotiSemanal(CodCotizacion, CU, Valor, Semana, FechaPublicada) values (" & Me.CodCotizacion & _
", " & Cal & ", " & val & ", " & Me.asemana & ", #" & Now & "#)"
MsgBox "Valor Recalculado a " & Cal & " y " & val
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
    Me.icoti = 1
    Me.proceso = "Guardando Tabla"
    Call CopiarTablaPrincipalCoti
    Me.proceso = "Tabla temporal de semana " & Me.asemana & " creada"
    Call DescargarTablaCoti
    Me.proceso = "Datos Correctamente guardados"
Next sem
Me.icoti = Null
Me.proceso = Null
MsgBox "Autoguardado Completo"
End Sub

Private Sub Comando245_Click()
Me.Filter = "(Tipo = 'Frutas' or Tipo = 'Verduras') And (Marca <> 'Almacen Global' or IsNull([Marca])) And CU = 0"
Me.FilterOn = True
Me.Requery
End Sub

Public Sub Comando437_Click()
If codigoLocal() = 0 Then
    Me.Filter = "(Tipo = 'Frutas' or Tipo = 'Verduras') And (Marca <> 'Almacen Global' or IsNull([Marca]))"
    Me.FilterOn = True
    Me.Requery
Else
    Me.Filter = "iif([CUA]=0,0,([CU]-[CUA])/[CUA]) <> 0"
    Me.FilterOn = True
    Me.Requery
End If

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

archivo = Direccion & "\" & Me.asemana & " - 4.Costo Unitario Cotizaciones.pdf"
DoCmd.OutputTo acOutputForm, "Calculo CU Cotizacion Semana NP", acFormatPDF, archivo, False

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

