' ==========================================================
' Modulo  : Form_RegistroCambiosPreIngresosPitaya
' Tipo    : 100
' Lineas  : 204
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:25
' ==========================================================
Option Compare Database
Private Sub ingresoprelista()

If MsgBox("Desea darle ingresos a todos los productos de la lista #" & Me.apreing & "?", vbYesNo, "Confirmacion") = vbNo Then
    Exit Sub
End If

On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim cantli As Integer
Dim cot As Long
Dim afech As Date
Dim canti As Double
Dim cotglob As Integer
Dim codsubpor As Long

miSQL = "SELECT CambiosPreIngresosPitaya.CodPreIngresoPitaya, CambiosPreIngresosPitaya.CodCotizacion," & _
" CambiosPreIngresosPitaya.Cantidad, PreIngresoPitaya.Fecha, CambiosPreIngresosPitaya.CodCambiosPreIngresoPitaya" & _
" FROM CambiosPreIngresosPitaya" & _
" INNER JOIN PreIngresoPitaya ON CambiosPreIngresosPitaya.CodPreIngresoPitaya = PreIngresoPitaya.CodPreIngresoPitaya" & _
" WHERE (((CambiosPreIngresosPitaya.CodPreIngresoPitaya)=" & Me.apreing & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst
afech = rst("Fecha")

For I = 1 To cantli
    If I >= Me.icoti Then
        cot = rst("CodCotizacion")
        cotglob = PorcionGlobalDePorcion(cot)
        canti = rst("Cantidad")
        
        If cotglob <> 0 Then ' es porcion
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cotglob & ", " & canti & ", #" & afech & "#)"
            
            'crear subporcionamiento primero
            DoCmd.RunSQL "INSERT INTO SubPorcionamiento(Procedencia, CodProcesamiento, Cantidad, Fecha) values" & _
            " (" & cotglob & ", 0, " & canti & ", #" & afech & "#)"
            Me.Requery
            codsubpor = UltimoCodSubporcionamiento()
    
            'lugo crear porcionamiento normal
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia, CodSubPorcionamiento) values" & _
            " (" & cot & ", 0, " & canti & ", #" & afech & "#, " & OperarioAzar() & ", " & cotglob & ", " & codsubpor & ")"
            DoCmd.SetWarnings True
        Else 'cotiacion general
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO [IngresosPitaya](CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cot & ", " & canti & ", #" & afech & "#)"
            DoCmd.SetWarnings True
        End If

    End If
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO [StatusCambiosPreingreso](CodCambiosPreIngresoPitaya, Status, CodOperario, Fecha)" & _
    " values (" & rst("CodCambiosPreIngresoPitaya") & ", -1, " & Me.codigoope & " ,#" & Now & "#)"
    DoCmd.SetWarnings True
    
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
MsgBox "Ingresos Completos"
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call ingresoprelista
End Sub


Private Sub Comando139_Click()
DoCmd.OpenForm "IngresoAutomaticoProductos", acNormal
[Forms]![IngresoAutomaticoProductos]![afechapro] = Me.fechaac
[Forms]![IngresoAutomaticoProductos]![adestino] = "[CambiosPreIngresosPitaya]"
[Forms]![IngresoAutomaticoProductos]![apreingreso] = Me.acodigo
[Forms]![IngresoAutomaticoProductos]![adesdeform] = "[RegistroCambiosPreIngresosPitaya]"
End Sub

Private Sub Comando148_Click()

'On Error GoTo ErrorHandler
Dim cantip As Integer
llave = Me.CodCambiosPreIngresoPitaya

' Confirmar antes de eliminar
If MsgBox("¿Estás seguro de querer eliminar el registro?", vbYesNo + vbInformation, "Confirmar") = vbYes Then

    ' Comprobar si el producto es una mezcla

    If ExisteMezcla(Me.CodCotizacion) = 1 Then
        ' Si es una mezcla, eliminar la mezcla y sus subproductos
    
        ' Obtener la cantidad de subproductos
        cantip = cantidadProductosMezcla(Me.CodCotizacion)
    
        ' Eliminar subproductos relacionados de SubPreIngresosPitaya
        For sp = 1 To cantip
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM CambiosPreIngresosPitaya WHERE CodCambiosPreIngresoPitaya = " & llave + sp
            DoCmd.SetWarnings True
        Next sp
    End If
    
    ' Eliminar la mezcla original de SubPreIngresosPitaya
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM CambiosPreIngresosPitaya WHERE CodCambiosPreIngresoPitaya = " & llave
    DoCmd.SetWarnings True

Else
    'no hace nada
End If

' Actualizar el formulario o la lista para reflejar los cambios
Me.Requery

Exit Sub

ErrorHandler:
MsgBox "Error al eliminar el registro"

End Sub

Private Sub Comando217_Click()
DoCmd.OpenForm "RegistrarProductosPorciones"
[Forms]![RegistrarProductosPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosPorciones]![adesdetabla] = "[CambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosPorciones]![adesdeform] = "[RegistroCambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosPorciones]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosPorciones]![guardado].width = 0
End Sub

Private Sub Comando241_Click()
DoCmd.OpenForm "RegistrarProductosMostrador"
[Forms]![RegistrarProductosMostrador]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosMostrador]![adesdetabla] = "[CambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosMostrador]![adesdeform] = "[RegistroCambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosMostrador]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosMostrador]![guardado].width = 0
End Sub

Private Sub Comando257_Click()
DoCmd.OpenForm "RegistrarProductosNoPorciones"
[Forms]![RegistrarProductosNoPorciones]![fechaprocedencia] = Me.fechaac
[Forms]![RegistrarProductosNoPorciones]![adesdetabla] = "[CambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosNoPorciones]![adesdeform] = "[RegistroCambiosPreIngresosPitaya]"
[Forms]![RegistrarProductosNoPorciones]![apreingre] = Me.acodigo
[Forms]![RegistrarProductosNoPorciones]![guardado].width = 0
End Sub

Private Sub Comando270_Click()

On Error GoTo Nulo
Dim cantic As Double
Dim cantip As Integer
Dim llavep As Integer
cantic = InputBox("Ingresar nueva cantidad", "Corregir Cantidad", 0)
llave = Me.CodCambiosPreIngresoPitaya

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE CambiosPreIngresosPitaya SET Cantidad = " & cantic & " WHERE CodCotizacion = " & Me.CodCotizacion & " AND CodPreIngresoPitaya = " & Me.acodigo
DoCmd.SetWarnings True

' Comprobar si el producto es una mezcla

If ExisteMezcla(Me.CodCotizacion) = 1 Then
    ' Si es una mezcla, modificar la cantidad en la mezcla y sus subproductos
    cantip = cantidadProductosMezcla(Me.CodCotizacion)
    ' Modificar la cantidad en subproductos relacionados de SubPreIngresosPitaya
        
    For sp = 1 To cantip
    
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE CambiosPreIngresosPitaya SET Cantidad = " & cantic & " WHERE CodCambiosPreIngresoPitaya = " & llave + sp
        DoCmd.SetWarnings True

    Next sp
    
End If

Me.Requery
Exit Sub

Nulo:
MsgBox "Volver a intentar"

End Sub

Private Sub Comando282_Click()
Me.apreing = Me.CodPreIngresoPitaya
Call ingresoprelista
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.Requery
End Sub
