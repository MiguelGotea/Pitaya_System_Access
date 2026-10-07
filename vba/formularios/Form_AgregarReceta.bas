' ==========================================================
' Modulo  : Form_AgregarReceta
' Tipo    : 100
' Lineas  : 374
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:26
' ==========================================================
Option Compare Database



Private Sub Comando398_Click()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String


Dim canti, Cant As Double
Dim inge, bati, tip As String
Dim codpo As Long
Dim insucla As Integer
Dim Orden As Integer

miSQL = "SELECT CodBatido, CodIngrediente, Cantidad, Tipo, codporcion, InsumoClave, ordenreceta" & _
" FROM SubReceta WHERE CodBatido='" & Me.cod & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
Cant = rst.RecordCount
rst.MoveFirst

For I = 1 To Cant
    If I >= Me.icoti Then
        bati = Me.acodbatido
        inge = rst("CodIngrediente")
        canti = rst("Cantidad")
        tip = rst("Tipo")
        insucla = rst("InsumoClave")
        Orden = Nz(rst("ordenreceta"), 0)
        
        If IsNull(rst("codporcion")) = True Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubReceta(CodIngrediente, CodBatido, Cantidad, Tipo, InsumoClave, ordenreceta) values" & _
            " ('" & inge & "', '" & bati & "', " & canti & ", '" & tip & "', " & insucla & ", " & Orden & ")"
            DoCmd.SetWarnings True
        Else
            codpo = rst("codporcion")
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO SubReceta(CodIngrediente, CodBatido, Cantidad, Tipo, codporcion, InsumoClave, ordenreceta) values" & _
            " ('" & inge & "', '" & bati & "', " & canti & ", '" & tip & "', " & codpo & ", " & insucla & ", " & Orden & ")"
            DoCmd.SetWarnings True
        End If
        
    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
'MsgBox "Receta Copiada"
Me.Requery
Me.cod.Requery
Exit Sub

AgainAgain:
MsgBox "Receta No Copiada"
Me.icoti = I
rst.Close
Call Comando398_Click
End Sub

Private Sub Comando401_Click()
On Error GoTo NoExiste
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO DBBatidos(CodBatido, Nombre, CodGrupo, Medida, Precio, Vigencia, Marca, CodSubGrupo, Endulzante) values" & _
" ('" & Me.acodbatido & "', '" & Me.anombre & "', " & Me.acodgrupo & ", '" & Me.amedida & _
"', " & Me.aprecio & ", " & Me.avigencia & ", '" & Me.amarca & "', " & Me.acodsubgrupo & ", " & Me.aendulzante & ")"
DoCmd.SetWarnings True
'MsgBox "Producto Agregado"
Exit Sub
NoExiste:
MsgBox "No se agregado producto"
End Sub

Private Sub Comando443_Click()
On Error GoTo NoExiste
If yaexistecodigobatido(Me.acodbatido) = 1 Then
    MsgBox "Codigo Ya Existe"
Else
    'MsgBox "Codigo Disponible"
End If

Exit Sub
NoExiste:
MsgBox "Error al buscar codigo"
End Sub

Private Sub Comando511_Click()
On Error GoTo NoExiste
If yaexistecodigobatido(Me.acodbatido) = 1 Then
    MsgBox "Codigo Ya Existe, no se gnera receta"
Else
    'MsgBox "Codigo Disponible"
    Me.resultado = Me.cod & " a " & Me.acodbatido & " " & Me.areceta & " " & Me.aprecio
    Call Comando401_Click
    Me.Requery
    Call Comando398_Click
    Me.Requery
    Me.cod = Me.acodbatido
    Call Comando57_Click
End If

Exit Sub
NoExiste:
MsgBox "Error al buscar codigo"

End Sub

Private Sub Comando513_Click()

DoCmd.OpenForm "Menu PITAYA Global"
End Sub

Private Sub Comando530_Click()
If MsgBox("Desea desactivar producto vigente?", vbYesNo) = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE DBBatidos SET Vigencia = 0 WHERE CodBatido = '" & cod & "'"
    DoCmd.SetWarnings True
    MsgBox "Producto Desactivado"
    Me.avigencia = 0
    Me.cod.Requery
End If
End Sub

Private Sub Comando57_Click()
Me.acodbatido = Me.cod
Me.anombre = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.amedida = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.aprecio = DLookup("[Precio]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.acodgrupo = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.avigencia = DLookup("[Vigencia]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.aendulzante = DLookup("[Endulzante]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.amarca = DLookup("[Marca]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.acodsubgrupo = DLookup("[CodSubGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.cod & "'")
Me.Requery
End Sub


Private Sub Comando599_Click()
Dim cambio As Integer
cambio = InputBox("Ingresar el factor de endulzante", "Factor de endulzante", 1)

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE DBBatidos SET Endulzante = " & cambio & " WHERE CodBatido = '" & cod & "'"
DoCmd.SetWarnings True

Me.aendulzante = cambio


End Sub

Function GenerarNuevaVersion(bati As String) As String
    Dim texto As String
    Dim posV As Integer
    Dim posD As Integer
    Dim numeroStr As String
    Dim Numero As Integer
    Dim prefijo As String
    Dim sufijo As String
    
    texto = Trim(bati)
    
    ' Encontrar la posición de "v" o "V"
    posV = InStr(1, texto, "v", vbTextCompare)
    If posV = 0 Then
        ' Si no encuentra "v", devolver el texto original
        GenerarNuevaVersion = texto
        Exit Function
    End If
    
    ' Buscar si hay "d" o "D" después de la "v"
    posD = InStr(posV, texto, "d", vbTextCompare)
    
    If posD > 0 Then
        ' Extraer el número entre "v" y "d"
        numeroStr = Mid(texto, posV + 1, posD - posV - 1)
        prefijo = Left(texto, posV)
        sufijo = Mid(texto, posD)
    Else
        ' Extraer todo después de "v" hasta el final
        numeroStr = Mid(texto, posV + 1)
        prefijo = Left(texto, posV)
        sufijo = ""
    End If
    
    ' Verificar si el número es válido y aumentarlo
    If IsNumeric(numeroStr) Then
        Numero = CInt(numeroStr) + 1
        GenerarNuevaVersion = prefijo & CStr(Numero) & sufijo
        
        ' Loop hasta encontrar un código que no exista
        Do While yaexistecodigobatido(GenerarNuevaVersion) = 1
            Numero = Numero + 1
            GenerarNuevaVersion = prefijo & CStr(Numero) & sufijo
        Loop
    Else
        ' Si no es un número, devolver el original
        GenerarNuevaVersion = texto
    End If
End Function

Private Sub Comando673_Click()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String

Dim cantx As Integer
Dim grupi As Integer
Dim Nombrex As String
Dim bati As String
Dim nuevobati As String
Dim codigoantiguo As String
Dim ValidarTipoBati As Integer

miSQL = "SELECT DBBatidos.CodBatido, DBBatidos.Vigencia, DBBatidos.CodGrupo, DBBatidos.Nombre" & _
" FROM DBBatidos" & _
" WHERE (((DBBatidos.Vigencia)<>0) AND ((DBBatidos.CodGrupo)=" & Me.grupomasivo & "))" & _
" ORDER BY DBBatidos.Nombre, DBBatidos.CodBatido"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantx = rst.RecordCount
rst.MoveFirst

For I = 1 To cantx
    If I >= Me.icoti Then
        grupi = rst("CodGrupo")
        Nombrex = rst("Nombre")
        bati = rst("CodBatido")
    
        Me.hgrupo = grupi
        Me.areceta = Nombrex
        Me.cod = bati
        codigoantiguo = bati
        Me.cod.Requery
        Call Comando57_Click
        'MsgBox "Producto Original"
        
        If DLookup("[Tipo]", "[Grupos]", "[CodGrupo]=" & grupi) = "Batido" Or DLookup("[Tipo]", "[Grupos]", "[CodGrupo]=" & grupi) = "Limonada" Then
            If Right(bati, 1) = "d" Then
                ValidarTipoBati = 3
            ElseIf InStr(1, bati, "Gv") > 0 Then
                ValidarTipoBati = 2
            ElseIf InStr(1, bati, "Mv") > 0 Then
                ValidarTipoBati = 1
            Else
                ValidarTipoBati = 0
            End If
        ElseIf DLookup("[Tipo]", "[Grupos]", "[CodGrupo]=" & grupi) = "Waffles" Then
            If Right(bati, 1) = "d" Then
                ValidarTipoBati = 5
            Else
                ValidarTipoBati = 4
            End If
        Else ' solo se admite batido, limonada y waffle , los demas no se cambia precios por grupos
            ValidarTipoBati = 0
        End If
        
        nuevobati = GenerarNuevaVersion(bati)
        Me.acodbatido = nuevobati
        
        Select Case ValidarTipoBati
        Case 1
            Me.aprecio = Me.masivo16oz
        Case 2
            Me.aprecio = Me.masivo20oz
        Case 3
            Me.aprecio = Me.masivo16ozpy
        Case 4
            Me.aprecio = Me.masivowaffle
        Case 5
            Me.aprecio = Me.masivowafflePY
        Case Else
            Me.aprecio = 0
        End Select
        
        If yaexistecodigobatido(Me.acodbatido) = 1 Then
            MsgBox "Codigo Ya Existe, no se hiso cambio"
            Me.logmasivo = Me.logmasivo & Me.cod & " NO SE CREO " & Me.acodbatido & " POR CODIGO EXISTENTE " & Me.areceta & Chr(13) & Chr(10)
        ElseIf nuevobati = bati Then 'Cuando nuevo batido es igual al anterior no se pudo aumetnara la version
            MsgBox "No se detecto la version"
            Me.logmasivo = Me.logmasivo & Me.cod & " NO SE CREO " & Me.acodbatido & " POR AUMENTO DE VERSION " & Me.areceta & Chr(13) & Chr(10)
        ElseIf Me.aprecio = 0 Then 'Nos e ddetecto tamano o si es d epeddosya
            MsgBox "No se detecto tamano o si es pedidosya"
            Me.logmasivo = Me.logmasivo & Me.cod & " NO SE CREO " & Me.acodbatido & " DETECCION DE TAMANO O PEDIDOSYA " & Me.areceta & Chr(13) & Chr(10)
        Else
            'MsgBox "Codigo Disponible"
            If Me.crearreceta = True Then ' check de crear receta activa, si no solo es reporte de como quedaria
                Call Comando511_Click
            End If
            Me.logmasivo = Me.logmasivo & codigoantiguo & " a " & Me.acodbatido & " " & Me.areceta & " " & Me.aprecio & Chr(13) & Chr(10)
        End If

    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
MsgBox "Reetas creadas con nuevo precio"
Me.Requery
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call Comando673_Click
End Sub


Private Sub btnFiltroVigencia_Click()

    Dim sqlVigentes As String
    Dim sqlTodos As String
    
    sqlVigentes = "SELECT [DBBatidos]![Nombre] & "" "" & IIf(InStr(1,[DBBatidos]![CodBatido],""d"")>0,""Delivery"",""EnTienda"") & "" "" & [DBBatidos]![Medida] & "" "" & [DBBatidos]![CodBatido] AS Name, " & _
                  "DBBatidos.CodBatido, DBBatidos.Vigencia, DBBatidos.CodGrupo, DBBatidos.Nombre " & _
                  "FROM DBBatidos " & _
                  "WHERE (((DBBatidos.Vigencia)<>0) AND ((DBBatidos.CodGrupo)=[Formularios]![AgregarReceta]![hgrupo]) AND ((DBBatidos.Nombre)=[Formularios]![AgregarReceta]![areceta])) " & _
                  "ORDER BY [DBBatidos]![Nombre] & "" "" & IIf(InStr(1,[DBBatidos]![CodBatido],""d"")>0,""Delivery"",""EnTienda"") & "" "" & [DBBatidos]![Medida] & "" "" & [DBBatidos]![CodBatido];"
    
    sqlTodos = "SELECT [DBBatidos]![Nombre] & "" "" & IIf(InStr(1,[DBBatidos]![CodBatido],""d"")>0,""Delivery"",""EnTienda"") & "" "" & [DBBatidos]![Medida] & "" "" & [DBBatidos]![CodBatido] AS Name, " & _
               "DBBatidos.CodBatido, DBBatidos.Vigencia, DBBatidos.CodGrupo, DBBatidos.Nombre " & _
               "FROM DBBatidos " & _
               "WHERE (((DBBatidos.CodGrupo)=[Formularios]![AgregarReceta]![hgrupo]) AND ((DBBatidos.Nombre)=[Formularios]![AgregarReceta]![areceta])) " & _
               "ORDER BY [DBBatidos]![Nombre] & "" "" & IIf(InStr(1,[DBBatidos]![CodBatido],""d"")>0,""Delivery"",""EnTienda"") & "" "" & [DBBatidos]![Medida] & "" "" & [DBBatidos]![CodBatido];"
    
    ' Usar el Caption del botón como estado
    If Me.btnFiltroVigencia.Caption = "Vigentes" Then
        ' Cambiar a TODOS
        Me.btnFiltroVigencia.Caption = "Todos"
        Me.btnFiltroVigencia.BackColor = RGB(100, 100, 100)  ' gris
        Me.cod.RowSource = sqlTodos
    Else
        ' Cambiar a VIGENTES
        Me.btnFiltroVigencia.Caption = "Vigentes"
        Me.btnFiltroVigencia.BackColor = RGB(0, 120, 0)      ' verde
        Me.cod.RowSource = sqlVigentes
    End If
    
    Me.cod.Value = Null
    Me.cod.Requery

End Sub


Private Sub Comando880_Click()
DoCmd.OpenForm "Control Consumo Diario"
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.btnFiltroVigencia.Caption = "Vigentes"
Me.btnFiltroVigencia.BackColor = RGB(0, 120, 0)
End Sub

Private Sub hgrupo_Exit(Cancel As Integer)
Me.areceta = ""
Me.areceta.Requery
Me.cod = ""
Me.cod.Requery
Me.resultado = ""
End Sub
Private Sub areceta_Exit(Cancel As Integer)
Me.cod = ""
Me.cod.Requery
Me.resultado = ""
End Sub
Private Sub cod_Exit(Cancel As Integer)

Me.resultado = ""
End Sub
