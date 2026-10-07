' ==========================================================
' Modulo  : Form_Calculo Pareto Semana
' Tipo    : 100
' Lineas  : 93
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:17
' ==========================================================
Option Compare Database

Public Sub Comando119_Click()
Me.Requery
Me.OrderBy = "total DESC"
Me.OrderByOn = True
End Sub

Private Sub Comando182_Click()

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim ING As String

miSQL = "SELECT DBIngredientes.CodIngrediente, DBIngredientes.TIPO1, ProductoCompraVenta([DBIngredientes]![CodIngrediente]) AS CYV, FRConsumoIngrediente([DBIngredientes]![CodIngrediente]," & Me.asemana & ")*FRCostoUnitarioIngrediente([DBIngredientes]![CodIngrediente]," & Me.asemana & ") AS Total" & _
" FROM DBIngredientes" & _
" WHERE (((DBIngredientes.TIPO1) = 'VARIABLES') And (Not (ProductoCompraVenta([DBIngredientes]![CodIngrediente])) = 1))" & _
" ORDER BY FRConsumoIngrediente([DBIngredientes]![CodIngrediente]," & Me.asemana & ")*FRCostoUnitarioIngrediente([DBIngredientes]![CodIngrediente]," & Me.asemana & ") DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

For I = 1 To 25
    ING = rst("CodIngrediente")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO CalParetoSemanal(CodIngrediente, Posicion, Semana, FechaPublicada) values ('" & ING & _
    "', " & I & ", " & Me.asemana & ", #" & Now & "#)"
    DoCmd.SetWarnings True
    rst.MoveNext
    
    '/ Barra de carga
    Me.carga = I / 25
    Me.carga.Left = Me.InsideWidth * I / 25
    Me.barracarga.width = Me.InsideWidth * I / 25
Next I

rst.Close
'MsgBox "Se almacenaron " & I - 1 & " datos correctamente"

'/reseteando barra de carga
Me.barracarga.width = 0
Me.carga.Left = 0
Exit Sub

AgainAgain:
Call Comando182_Click
End Sub

Public Sub Comando187_Click()
If codigoLocal() = 0 Then
    Me.Filter = "Tipo = 'Frutas' or Tipo = 'Verduras'"
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

archivo = Direccion & "\" & Me.asemana & " - 6.Pareto Semanal.pdf"
DoCmd.OutputTo acOutputForm, "Calculo Pareto Semana", acFormatPDF, archivo, False
End Sub

Private Sub Comando230_Click()
Dim sem As Integer
For sem = Me.sdesde To Me.shasta
    Me.asemana = sem
    Call Comando182_Click
Next sem
MsgBox "Autoguardado Completo"
End Sub

Private Sub Form_Close()
Me.OrderBy = ""
Me.OrderByOn = False
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.sactual = numerosemana(Date)
Me.sdesde = numerosemana(Date) - 1
Me.shasta = numerosemana(Date) - 1
End Sub

