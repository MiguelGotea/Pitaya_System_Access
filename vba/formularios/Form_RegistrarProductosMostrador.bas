' ==========================================================
' Modulo  : Form_RegistrarProductosMostrador
' Tipo    : 100  |  Lineas: 113
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:15
' ==========================================================

Option Compare Database

Private Sub Comando131_Click()
On Error GoTo AgainAgain
Dim rst As DAO.Recordset
Dim miSQL As String
Dim rsPorciones As DAO.Recordset

Dim cantli As Integer
Dim cot As Integer

Dim canti As Double
Dim tota As Double
Dim cantip As Integer

tota = InputBox("Ingrese la cantidad de productos a registrar", "Cantidad", 0)

miSQL = "SELECT SubReceta.CodBatido, SubReceta.CodIngrediente, SubReceta.Cantidad, SubReceta.codporcion" & _
" FROM SubReceta WHERE (((SubReceta.CodBatido)='" & Me.CodBatido & "') AND ((SubReceta.InsumoClave)<>0))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
cantli = rst.RecordCount
rst.MoveFirst

For I = 1 To cantli
    If I >= Me.icoti Then
        If IsNull(rst("codporcion")) Then 'tiene codigo cotizacion directo a ingresar
            cot = CotiDirectoDeIngrediente(rst("CodIngrediente"))
        Else 'tiene codigo de porcion entonces se agrega la porcion
            cot = rst("codporcion")
        End If
        
        canti = Round(rst("Cantidad") / DLookup("[Conversion]", "[Cotizaciones]", "[CodCotizacion]=" & cot), 1)
    
        
        If Me.adesdetabla = "[SubPreIngresosPitaya]" Or Me.adesdetabla = "[CambiosPreIngresosPitaya]" Then
            
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & cot & ", " & tota * canti & ", " & Me.apreingre & ")"
            DoCmd.SetWarnings True
            
            ' Verificar si existe relación con CodCotizacionPorcion
        
            If ExisteMezcla(cot) = 1 Then
                ' Agregar la mezcla original
                ' Obtener CodCotizacionPorcion relacionados
                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & cot)
                
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & tota * canti & ", " & Me.apreingre & ")"
                    DoCmd.SetWarnings True
                    rsPorciones.MoveNext
                Loop
                rsPorciones.Close
            End If
            
        Else
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & cot & ", " & tota * canti & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
            
            If ExisteMezcla(cot) = 1 Then
                ' Agregar la mezcla original
                Set rsPorciones = CurrentDb.OpenRecordset("SELECT CodCotizacionPorcion FROM MezclaPorciones WHERE CodCotizacionMezcla = " & cot)
                
                ' Agregar CodCotizacionPorcion relacionados con la misma cantidad
                Do While Not rsPorciones.EOF
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO " & Me.adesdetabla & "(CodCotizacion, Cantidad, Fecha)" & _
                        " values (" & rsPorciones!CodCotizacionPorcion & ", " & tota * canti & ", #" & Me.fechaprocedencia & "#)"
                    DoCmd.SetWarnings True
                    rsPorciones.MoveNext
                Loop
            
                rsPorciones.Close
            End If
            
        End If
        
    End If
    rst.MoveNext
Next I

rst.Close
Me.icoti = 1
Me.guardado.Requery
Exit Sub

AgainAgain:
Me.icoti = I
rst.Close
Call Comando131_Click
End Sub

Private Sub Form_Close()
If Me.adesdeform = "[Ingreso Inventario Pitaya]" Then
    Forms(Me.adesdeform).Form.Subformulario_Inventario_Cotizacion.Requery
Else
    Forms(Me.adesdeform).Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False

End Sub

