' ==========================================================
' Modulo  : Form_Menu Endulzantes
' Tipo    : 100
' Lineas  : 451
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:21
' ==========================================================



Private Sub alto_Click() ' AZUCAR ALTA
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer

Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")

tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))
promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 0, 'AZUCAR ALTO', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    Dim tieneazucar As Integer
    tieneazucar = DLookup("[Endulzante]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
    
    If tieneazucar > 0 Then
        'Se agrega azucar de la receta original +0.5
        DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
        " VALUES ('E02', " & 1.5 * tieneazucar & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"
    Else
        'Se agrega azucar de la receta original +1
        DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
        " VALUES ('E02', 1, " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"
    End If
    
DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"

Exit Sub

Again:
Call alto_Click
End Sub
    
Private Sub Comando1968_Click() ' AZUCAR NORMAL
'On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))
promox = IIf(tipovinculo = 20, 104, 5) ' grupo combos hotel livano

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 0, 'AZUCAR NORMAL', " & Me.vinculos & ")"

    Dim subpe As Long
    subpe = UltimoSubPedido()
    Dim tieneazucar As Integer
    tieneazucar = DLookup("[Endulzante]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
    
    'Se agrega azucar de la receta original +0.5
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E02', " & tieneazucar * 1 & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

    
DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"

Exit Sub

Again:
Call Comando1968_Click
End Sub

Private Sub Comando1970_Click() ' AZUCAR BAJO
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 0, 'AZUCAR BAJO', " & Me.vinculos & ")"

    Dim subpe As Long
    subpe = UltimoSubPedido()
    Dim tieneazucar As Integer
    tieneazucar = DLookup("[Endulzante]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")

    'Se resta azucar de la receta original
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E02', " & tieneazucar * 0.5 & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"
    
DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call Comando1970_Click
End Sub

Private Sub miel_Click() ' MIEL
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'MIEL', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la miel, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E01', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"

Exit Sub

Again:
Call miel_Click
End Sub



Private Sub Comando1976_Click() ' STEVIA X 1
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False
    
    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'STEVIA x 1', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la STEVIA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E03', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call Comando1976_Click
End Sub

Private Sub stevia_Click() ' STEVIA X 2
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, Codpromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'STEVIA x 2', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la STEVIA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E03', " & 2 * Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call stevia_Click
End Sub

Private Sub Comando1975_Click() ' STEVIA X 3
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, Codpromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'STEVIA x 3', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la STEVIA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E03', " & 3 * Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call Comando1975_Click
End Sub

Private Sub splenda_Click() ' SPLENDA X 1
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'SPLENDA x 1', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la SPLENDA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, Codpromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E04', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call splenda_Click
End Sub


Private Sub Comando1981_Click() ' SPLENDA X 2
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False
    
    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'SPLENDA x 2', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la SPLENDA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E04', " & 2 * Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"


DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call Comando1981_Click
End Sub


Private Sub Comando1982_Click() ' SPLENDA X 3
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'SPLENDA x 3', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    
    'se agrega la SPLENDA, anexado al principal
    DoCmd.RunSQL "INSERT INTO SubPedido(CodBatido, Cantidad, CodPedido, CodPromocion, Observaciones, Azucar, Vinculo)" & _
    " VALUES ('E04', " & 3 * Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", 'Para el ' & '" & Me.nbatido & "', '-', " & subpe & ")"
    
DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"


Exit Sub

Again:
Call Comando1982_Click
End Sub

Private Sub Comando1962_Click() ' SIN ENDULZANTE
On Error GoTo Again

Dim tipovinculo As Integer
Dim promox As Integer
Dim grupi As Integer
grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cbatido & "'")
tipovinculo = IIf(Me.vinculos = 0, 0, DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & DLookup("[CodBatido]", "[SubPedido]", "[CodSubPedido]=" & Me.vinculos & "") & "'"))

promox = IIf(tipovinculo = 20, 104, 5)

DoCmd.SetWarnings False

    ' se agrega el producto principal
    DoCmd.RunSQL "INSERT INTO SubPedido (CodBatido, Cantidad, CodPedido, CodPromocion, SinAzucar, Azucar, Vinculo)" & _
    " VALUES ('" & Me.cbatido & "', " & Me.ccantidad & ", " & Me.cpedido & ", " & promox & ", -1, 'SIN ENDULZANTE', " & Me.vinculos & ")"
    
    Dim subpe As Long
    subpe = UltimoSubPedido()
    

DoCmd.SetWarnings True

'cerrar ventana de producto especifico
If tipovinculo = 20 Then
    DoCmd.Close acForm, menucombo
End If
DoCmd.Close acForm, "Menu Endulzantes"

Exit Sub

Again:
Call Comando1962_Click
End Sub

Private Sub Comando2051_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
