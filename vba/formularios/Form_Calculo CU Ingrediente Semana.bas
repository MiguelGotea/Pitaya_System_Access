' ==========================================================
' Modulo  : Form_Calculo CU Ingrediente Semana
' Tipo    : 100  |  Lineas: 156
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:17
' ==========================================================

Option Compare Database

Private Sub DescargarTabla()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Cant As Integer
Dim ingre As String
Dim Cal As Double
Dim sema As Integer

sema = Me.asemana
miSQL = "SELECT CodIngrediente, formula FROM TempCUIngredientes ORDER BY CodIngrediente"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

Me.proceso = "Copiando Datos"
For I = 1 To Cant
    If I >= Me.icoti Then
        ingre = rst("CodIngrediente")
        Cal = rst("formula")
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO CalCUIngSemanal(CodIngrediente, CU, Semana, FechaPublicada) values ('" & ingre & _
        "', " & Cal & ", " & sema & ", #" & Now & "#)"
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
Call DescargarTabla
End Sub

Private Sub CopiarTablaPrincipal()

On Error GoTo AgainAgain2

If DCount("*", "MSysObjects", "Type=1 AND Name='TempCUIngredientes'") > 0 Then
DoCmd.RunSQL "DROP TABLE TempCUIngredientes"
End If

'Crear carpeta temporal
DoCmd.SetWarnings False
DoCmd.RunSQL "SELECT x.* INTO TempCUIngredientes" & _
" FROM (SELECT DBIngredientes.CodIngrediente, DBIngredientes.TIPO1, DBIngredientes.Vigente, TiposVariables.Control," & _
" CostoUnitarioIngredientexCompra([DBIngredientes]![CodIngrediente]," & Me.asemana & ") AS formula" & _
" FROM DBIngredientes INNER JOIN TiposVariables ON DBIngredientes.Tipo = TiposVariables.Tipo" & _
" WHERE (((DBIngredientes.TIPO1)='VARIABLES') AND ((DBIngredientes.Vigente)<>0) AND ((TiposVariables.Control)<>0)))x"
DoCmd.SetWarnings True

'MsgBox "Datos Guaradados en tabla temporal"
Exit Sub

AgainAgain2:
Call CopiarTablaPrincipal
End Sub

Private Sub Comando119_Click()
Me.Requery
Me.Filter = ""
Me.FilterOn = True
Me.Requery
End Sub

Private Sub Comando170_Click()
Me.icoti = 1
Me.proceso = "Guardando Tabla"
Call CopiarTablaPrincipal
Me.proceso = "Tabla temporal de semana " & Me.asemana & " creada"
Call DescargarTabla
Me.icoti = Null
Me.proceso = Null
MsgBox "Autoguardado Completo"
End Sub


Private Sub Comando201_Click()
Dim Cal As Double
Cal = CostoUnitarioIngredientexCompra(Me.CodIngrediente, Me.asemana)
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO CalCUIngSemanal(CodIngrediente, CU, Semana, FechaPublicada) values ('" & Me.CodIngrediente & _
"', " & Cal & ", " & Me.asemana & ", #" & Now & "#)"
DoCmd.SetWarnings True
MsgBox "Valor Recalculado a " & Cal
Me.CU.Requery
End Sub

Private Sub Comando240_Click()
Dim Cal As Double
Cal = InputBox("Ingresar Valor Unitario de " & Me.Nombre)
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO CalCUIngSemanal(CodIngrediente, CU, Semana, FechaPublicada) values ('" & Me.CodIngrediente & _
"', " & Cal & ", " & Me.asemana & ", #" & Now & "#)"
DoCmd.SetWarnings True
MsgBox "Valor Recalculado a " & Cal
Me.CU.Requery
End Sub



Private Sub Comando230_Click()
Dim sem As Integer
For sem = Me.sdesde To Me.shasta
    Me.asemana = sem
    Me.icoti = 1
    Me.proceso = "Guardando Tabla"
    Call CopiarTablaPrincipal
    Me.proceso = "Tabla temporal de semana " & Me.asemana & " creada"
    Call DescargarTabla
    Me.proceso = "Datos Correctamente guardados"
Next sem
Me.icoti = Null
Me.proceso = Null
MsgBox "Autoguardado Completo"
End Sub

Private Sub anombre_Enter()
Me.reinencontrar = 0
End Sub


Private Sub Comando255_Click()
If Me.reinencontrar = 0 Then ' nueva busqueda
    Me.Requery
End If
Me.Nombre.SetFocus
DoCmd.FindRecord Me.anombre, acAnywhere, False, acCurrent, , , False
Me.reinencontrar = 1
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.sactual = numerosemana(Date)
Me.sdesde = numerosemana(Date) - 1
Me.shasta = numerosemana(Date) - 1

End Sub


