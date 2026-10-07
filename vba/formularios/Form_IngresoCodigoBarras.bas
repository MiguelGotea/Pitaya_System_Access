' ==========================================================
' Modulo  : Form_IngresoCodigoBarras
' Tipo    : 100
' Lineas  : 487
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:19
' ==========================================================
Option Compare Database

Private Sub cantidad_GotFocus()
Me.codigo.SetFocus
End Sub

Private Sub codigo_Exit(Cancel As Integer)
On Error GoTo Salir
Dim val As Long
Dim codin As Long
Dim cantin As Long
Dim codinglob As Long

If Me.codigo = "C66" Then
    
    cantin = Me.Cantidad
    
    Select Case Me.adesde
    
        Case "[IngresosPitaya]"

            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & 437 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'envase
            DoCmd.SetWarnings True
            'Registro Ingredientes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & 820 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" ' galleta de avena global
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
            " (" & 821 & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & 820 & ")"
            DoCmd.SetWarnings True

    
        Case "[Inventario Cotizacion]"
            
            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & 437 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'envase
            DoCmd.SetWarnings True
            'Registrar ingredientes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & 821 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" ' galleta de avena global
            DoCmd.SetWarnings True
            
            
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
            
            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
            " (" & 437 & ", " & cantin & ", " & Me.apreingreso & ")" 'envase
            DoCmd.SetWarnings True
            'Registrar ingredientes
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
            " (" & 821 & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
            

        Case Else
        
            MsgBox "No hay ninguna ventana abierta para registrar producto"
    End Select
    
ElseIf Me.codigo Like "PIT*" Then   'Producto marca pitaya
   
    Dim aetiq1 As Integer
    Dim aetiq2 As Integer
    
    cantin = Me.Cantidad
    
    'datos especificos
    Select Case Me.codigo
        Case "PIT002" 'semilla cacao 151gr
            aetiq1 = 700 'ziploc
            aetiq2 = 795 'sticker marca pitaya
            codin = 701 'codigo porcion no global de conntenido
            codinglob = 822 'porcion almacen global
        Case "PIT014" 'cocoa 151gr
            aetiq1 = 700 'ziploc
            aetiq2 = 795 'sticker marca pitaya
            codin = 818 'codigo porcion no global de conntenido
            codinglob = 819 'porcion almacen global
        Case "PIT008" 'Granola 230gr
            aetiq1 = 700 'ziploc
            aetiq2 = 795 'sticker marca pitaya
            codin = 711 'codigo porcion no global de conntenido
            codinglob = 747 'porcion almacen global
        Case "PIT001" 'Harina Avena 200gr
            aetiq1 = 700 'ziploc
            aetiq2 = 795 'sticker marca pitaya
            codin = 703 'codigo porcion no global de conntenido
            codinglob = 737 'porcion almacen global
        Case "PIT004" 'Pecanas
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 560 'codigo porcion no global de conntenido
            codinglob = 823 'porcion almacen global
        Case "PIT005" 'Maranon
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 705 'codigo porcion no global de conntenido
            codinglob = 824 'porcion almacen global
        Case "PIT006" 'Almendra
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 706 'codigo porcion no global de conntenido
            codinglob = 735 'porcion almacen global
        Case "PIT007" 'Pistachos
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 707 'codigo porcion no global de conntenido
            codinglob = 828 'porcion almacen global
        Case "PIT010" 'Semillas mixtas, solo mani
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 760 'codigo porcion no global de conntenido
            codinglob = 826 'porcion almacen global
        Case "PIT012" 'Miel Grande
            aetiq1 = 756 'envase grane
            aetiq2 = 795 'sticker marca pitaya
            codin = 762 'codigo porcion no global de conntenido
            codinglob = 755 'porcion almacen global
        Case "PIT013" 'Miel pequeña
            aetiq1 = 796 'envase  pequeno
            aetiq2 = 795 'sticker marca pitaya
            codin = 798 'codigo porcion no global de conntenido
            codinglob = 797 'porcion almacen global
        Case "PIT016" 'Frutos Deshidratadis
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 898 'codigo porcion no global de conntenido
            codinglob = 899 'porcion almacen global
        Case "PIT015" 'Frutos Deshidratadis
            aetiq1 = 438 'bolsa celofan
            aetiq2 = 436 'etiqueta kraft
            codin = 900 'codigo porcion no global de conntenido
            codinglob = 901 'porcion almacen global
        Case Else
            MsgBox "No se encontro codigo de barra de Marca Pitaya"
            Exit Sub
    End Select
    

    Select Case Me.adesde
    
        Case "[IngresosPitaya]"

            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & aetiq1 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'envase
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & aetiq2 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'etqieyta
            
            'Registro Ingredientes
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & codinglob & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
            " (" & codin & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob & ")"
            DoCmd.SetWarnings True
            
            'adicional de semillas mixtas
            If Me.codigo = "PIT010" Then
            
                codin = 759 'codigo porcion no global de conntenido
                codinglob = 825 'porcion almacen global
                
                DoCmd.SetWarnings False
                'Almencdra
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
                " (" & codinglob & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
                DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
                " (" & codin & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob & ")"
                
                codin = 761 'codigo porcion no global de conntenido
                codinglob = 827 'porcion almacen global
                'Pasas
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
                " (" & codinglob & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
                DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia) values" & _
                " (" & codin & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob & ")"
                DoCmd.SetWarnings True
                
            End If
    
        Case "[Inventario Cotizacion]"
            
            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & aetiq1 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'envase
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & aetiq2 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)" 'etqieyta
            
            'Registrar ingredientes
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
            " (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
            
            'adicional de semillas mixtas
            If Me.codigo = "PIT010" Then
            
                codin = 759 'codigo porcion no global de conntenido
                codinglob = 825 'porcion almacen global
                
                DoCmd.SetWarnings False
                'Almencdra
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
                " (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
                
                codin = 761 'codigo porcion no global de conntenido
                codinglob = 827 'porcion almacen global
                'Pasas
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha) values" & _
                " (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
                DoCmd.SetWarnings True

            End If
            
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
            
            'Registrar etiquetas
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
            " (" & aetiq1 & ", " & cantin & ", " & Me.apreingreso & ")" 'envase
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
            " (" & aetiq2 & ", " & cantin & ", " & Me.apreingreso & ")" 'etqieyta
              
            'Registrar ingredientes
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
            " (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
            
            'adicional de semillas mixtas
            If Me.codigo = "PIT010" Then
            
                codin = 759 'codigo porcion no global de conntenido
                codinglob = 825 'porcion almacen global
                
                DoCmd.SetWarnings False
                'Almencdra
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
                " (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
                
                codin = 761 'codigo porcion no global de conntenido
                codinglob = 827 'porcion almacen global
                'Pasas
                DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya) values" & _
                " (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
                DoCmd.SetWarnings True

            End If
        Case Else
        
            MsgBox "No hay ninguna ventana abierta para registrar producto"
    End Select
    
ElseIf Me.codigo Like "P*" Then   'Porcion
    
    val = Right(Me.codigo, Len(Me.codigo) - 1)
    
    codin = Int(val / 10000) 'codigo de porcion, no de almacen global
    codinglob = PorcionGlobalDePorcion(codin)
    cantin = (val - Int(val / 10000) * 10000) * Me.Cantidad

    Select Case Me.adesde
    
        Case "[IngresosPitaya]"
        
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codinglob & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            'Porcionado
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia)" & _
            " values (" & codin & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob & ")"
            DoCmd.SetWarnings True
        
        Case "[Inventario Cotizacion]"
        
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
            
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
    
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
        
        Case Else
        
            MsgBox "No hay ninguna ventana abierta para registrar producto"
        
    End Select
    
ElseIf Me.codigo Like "ESPECIAL*" Then   'Porcion combinada
    
    Dim codin1 As Integer
    Dim codin2 As Integer
    Dim codinglob1 As Integer
    Dim codinglob2 As Integer
    
    Select Case Me.codigo
        Case "ESPECIAL30XC10C10"
            cantin = 30 * Me.Cantidad
            codin1 = 326 'cocoa
            codinglob1 = 718
            codin2 = 371 'cacao
            codinglob2 = 723
        Case "ESPECIAL30XC8C8"
            cantin = 30 * Me.Cantidad
            codin1 = 325 'cocoa
            codinglob1 = 717
            codin2 = 370 'cacao
            codinglob2 = 722
        Case "ESPECIAL10XW60C2"
            cantin = 10 * Me.Cantidad
            codin1 = 693 'cocoa
            codinglob1 = 721
            codin2 = 692 'waffle
            codinglob2 = 727
        Case Else
            MsgBox "No existe porcion especial"
            Exit Sub
    End Select
    
    Select Case Me.adesde
    
        Case "[IngresosPitaya]"
        
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codinglob1 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            'Porcionado
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia)" & _
            " values (" & codin1 & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob1 & ")"
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codinglob2 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            'Porcionado
            DoCmd.RunSQL "INSERT INTO Porcionamiento(CodCotizacion, CodProcesamiento, Cantidad, Fecha, CodOperario, Procedencia)" & _
            " values (" & codin2 & ", 0, " & cantin & ", #" & Me.fechaprocedencia & "#, " & OperarioAzar() & ", " & codinglob2 & ")"
            DoCmd.SetWarnings True
        
        Case "[Inventario Cotizacion]"
        
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codin1 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codin2 & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
            
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
    
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & codin1 & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & codin2 & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
        
        Case Else
        
            MsgBox "No hay ninguna ventana abierta para registrar producto"
        
    End Select
    
ElseIf Me.codigo Like "AUTO*" Then   'autogenerado porque no tiene codigo de barra
    
    val = Right(Me.codigo, Len(Me.codigo) - 4)
    codin = val
    cantin = Me.Cantidad
    
    Select Case Me.adesde
    
        Case "[IngresosPitaya]", "[Inventario Cotizacion]"
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
            
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
            DoCmd.SetWarnings False
            'Ingreso
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
            
        Case Else
            MsgBox "No hay ninguna ventana abierta para registrar producto"
            
    End Select
    
Else ' resto de productos cotiacion con codigo de barra normal
    
    codin = DLookup("[CodCotizacion]", "[CodigoBarraCotizacion]", "[CodigoBarra]='" & Me.codigo & "'")
    cantin = Me.Cantidad
    
    Select Case Me.adesde
    
        Case "[IngresosPitaya]", "[Inventario Cotizacion]"
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, Fecha)" & _
            " values (" & codin & ", " & cantin & ", #" & Me.fechaprocedencia & "#)"
            DoCmd.SetWarnings True
    
        Case "[SubPreIngresosPitaya]", "[CambiosPreIngresosPitaya]"
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO " & Me.adesde & "(CodCotizacion, Cantidad, CodPreIngresoPitaya)" & _
            " values (" & codin & ", " & cantin & ", " & Me.apreingreso & ")"
            DoCmd.SetWarnings True
            
        Case Else
            MsgBox "No hay ninguna ventana abierta para registrar producto"
            
    End Select
End If


Me.acodigo = codin
Me.acantidad = cantin
Me.aprodu = nombreproductocoti(Me.acodigo)
    
Me.codigo = ""
Me.codigo.SetFocus
Me.Cantidad = 1

Exit Sub

Salir:
If IsNull(Me.codigo) = True Or Me.codigo = "" Then
Else
    MsgBox "Codigo de Barras no existe"
    Me.Cantidad = 1
End If

Me.codigo = ""
Me.codigo.SetFocus

'Me.Codigo.SelStart = 1
'SendKeys "{TAB}"
End Sub

Private Sub Comando669_Click()
On Error GoTo Salir
Dim ingcant As Long
ingcant = InputBox("Ingresar Nueva Cantidad:", "")
Me.Cantidad = ingcant

Me.codigo = ""
Me.codigo.SetFocus

Exit Sub
Salir:
MsgBox "Ingresar la cantidad correctamente"
End Sub

Private Sub Form_Close()
If CurrentProject.AllForms("Ingresos a Pitaya").IsLoaded Then
    [Forms]![Ingresos a Pitaya].Form.Requery
End If
If CurrentProject.AllForms("Ingreso Inventario Pitaya").IsLoaded Then
    [Forms]![Ingreso Inventario Pitaya].Form.Subformulario_Inventario_Cotizacion.Requery
End If
If CurrentProject.AllForms("RegistroPreIngresosPitaya").IsLoaded Then
    [Forms]![RegistroPreingresosPitaya].Form.Requery
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
End Sub
