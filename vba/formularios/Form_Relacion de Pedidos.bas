' ==========================================================
' Modulo  : Form_Relacion de Pedidos
' Tipo    : 100  |  Lineas: 1389
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:10
' ==========================================================

Option Compare Database



Private Sub resetearpromocion(subpedi As Long)
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 5" & _
" WHERE SubPedido.CodSubPedido = " & subpedi
DoCmd.SetWarnings True
End Sub



Private Sub CodPromocion_Change()
Dim codi As Long
Dim prom As Integer
Dim batix As String
Dim pedid As Long
Dim cpunt As Double
Dim cantiped As Double
Dim grupy As Integer
Dim grantgrup As String
Dim subgrup As Integer
Dim nombrep As String

Dim ratio1 As Double
Dim ratio2 As Double
Dim ratio3 As Double
Dim ratio4 As Double
Dim ratio5 As Double
Dim ratio6 As Double
Dim ratio20 As Double
Dim ratio200 As Double
        
Dim ocluby As Long

        
codi = Me.CodSubPedido
prom = Me.CodPromocion
batix = Me.CodBatido
pedid = Me.CodPedido
cantiped = Me.Cantidad
'Me.CodPromocion.Requery

If DLookup("[usointerno]", "[DBPromociones]", "[CodPromocion]=" & prom) <> 0 Then 'promciones de Uso interno no se puede elegir manualmente
    Me.Requery
    resetearpromocion (codi)
    [Forms]![Nota de Pedido].Form.Secundario57.Requery
    Exit Sub
End If

[Forms]![Nota de Pedido].Form.Requery

cpunt = [Forms]![Nota de Pedido].pacumulado
grupy = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
grantgrup = DLookup("[Tipo]", "[Grupos]", "[CodGrupo]=" & grupy)
subgrup = DLookup("[CodSubGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
nombrep = DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
ocluby = DLookup("[CodCliente]", "[NotaDePedido]", "[CodPedido]=" & pedid)

'[Forms]![Nota de Pedido].Form.Secundario57.Requery
'MsgBox Me.CodPromocion

Select Case prom
    Case 93 '1. Lunes Waffle 10%off x compra batido

        If DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = 14 Then 'promocion ubicado en waffle
            If pedidocontienegrupo("Batido", pedid) = 0 Then 'pedido contiene no contienen ningun batido
                MsgBox "No hay batido facturado que acompañe a waffle, no puede aplicar esta promocion"
                resetearpromocion (codi)
                DoCmd.SetWarnings True
            End If
        Else
            MsgBox "Promocion tiene que ser aplicado a Waffle"
            resetearpromocion (codi)
        End If

    Case 94 '2. Martes 10%off x2 Dinamita
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            MsgBox "Promocion solo aplica a a batidos medianos"
            resetearpromocion (codi)
        Else
            If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Dinamita" Then ' restringir promo a dinamitas

                ratio1 = pedidocontieneproducto("Dinamita", pedid)
                ratio2 = pedidocontienepromocion(94, pedid)
    
                If ratio1 / 2 - Int(ratio1 / 2) = 0 Then 'hay cantidad de dinamitas par
                    If ratio2 > ratio1 Then 'No se admite + proms que cantidad de batidos
                        MsgBox "No puede facturar mas promociones ya que no tiene suficientes batidos"
                        resetearpromocion (codi)
                    End If
                Else ' hay cantidad de dinamitas impar
                    If ratio2 >= ratio1 Then 'prmociones igual o mayor que productos
                        MsgBox "No puede facturar mas promociones ya que no tiene suficientes batidos"
                        resetearpromocion (codi)
                    End If
                End If
    
            Else
                MsgBox "Promocion aplicado solo a dinamitas"
                resetearpromocion (codi)
            End If
        End If
    Case 55 '3. Miercoles 3x2 Batidos
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            MsgBox "Promocion solo aplica a a batidos medianos"
            resetearpromocion (codi)
        Else
            Dim grupi As Integer
            grupi = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
            If grupi = 1 Or grupi = 3 Or grupi = 8 Then ' batidos saludables, clasicos y especiales
                ' cumple tipo de batido de promocion

                   ratio3 = pedidocontienegrupo("Batido", pedid)
                   ratio4 = pedidocontienepromocion(55, pedid)
                   
                   If ratio4 > Int(ratio3 / 3) Then
                       MsgBox "No hay suficientes batidos para aplicar promocion"
                       resetearpromocion (codi)
                   End If
            Else
                MsgBox "Promocion aplicado solo a saludables, clasicos y especiales"
                resetearpromocion (codi)
            End If
        End If
    Case 95 '4. Jueves 10%off x2 Saludables
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            MsgBox "Promocion solo aplica a a batidos medianos"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a pequeños
            If DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = 8 Then ' restringir promo a saludables

                ratio5 = pedidocontienesubgrupo(8, pedid)
                ratio6 = pedidocontienepromocion(95, pedid)
    
                If ratio5 / 2 - Int(ratio5 / 2) = 0 Then 'hay cantidad de saludables par
                    If ratio6 > ratio5 Then 'No se admite + proms que cantidad de batidos
                        MsgBox "No puede facturar mas promociones ya que no tiene suficientes batidos saludables para aplicar"
                        resetearpromocion (codi)
                    End If
                Else ' hay cantidad de dinamitas impar
                    If ratio6 >= ratio5 Then 'prmociones igual o mayor que productos
                        MsgBox "No puede facturar mas promociones ya que no tiene suficientes batidos"
                        resetearpromocion (codi)
                    End If
                End If
    
            Else
                MsgBox "Promocion aplicado solo a Saludables"
                resetearpromocion (codi)
            End If
        End If
    Case 105 '2 clasicos x 95
    If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
        MsgBox "Promocion solo aplica a a batidos medianos"
        resetearpromocion (codi)
    Else ' promocion aplicado solo a pequeños
        If DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = 3 Then ' restringir promo a clasicos
            ' si es clasico
        Else
            MsgBox "Promocion aplicado solo a Saludables"
            resetearpromocion (codi)
        End If
    End If
    Case 77, 120, 131 '5. Viernes Agranda Batido Club
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
            MsgBox "Promocion solo aplica a  batidos grandes"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a grandes
            Dim grupix As Integer
            grupix = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
            If grupix = 1 Or grupix = 3 Or grupix = 8 Then ' batidos saludables, clasicos y especiales
                'aplic normal
            Else
                MsgBox "Promocion solo aplica saludables, clasicos y especiales"
                resetearpromocion (codi)
            End If
        End If
        
    Case 141 '2. 2 premium mediano x 200
        If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            ' aplica correctamente
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                Dim grupix2 As Integer
                grupix2 = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                If grupix2 = 2 Then ' batidos saludables, clasicos y especiales
                    'aplic normal
                Else
                    MsgBox "Promocion solo aplica batidos Premium"
                    resetearpromocion (codi)
                End If
                
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion solo aplica a  batidos Medianos"
                resetearpromocion (codi)
            End If
        Else
            'no deja cambiar promocion
            resetearpromocion (codi)
            MsgBox "Promocion aplica a numero par de productos"
        End If
    Case 194, 217 '2. 2 premium mediano x 215
        'If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            ' aplica correctamente
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                If grupy = 2 Then ' batidos saludables, clasicos y especiales
                    'aplic normal
                Else
                    MsgBox "Promocion solo aplica batidos Premium"
                    resetearpromocion (codi)
                End If
                
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion solo aplica a  batidos Medianos"
                resetearpromocion (codi)
            End If
        'Else
        '    'no deja cambiar promocion
        '    resetearpromocion (codi)
        '    MsgBox "Promocion aplica a numero par de productos"
        'End If
    Case 190, 215, 237 '2. 2 premium mediano x 179
        'If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            ' aplica correctamente
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                If grupy = 1 Then ' batidos   especiales
                    'aplic normal
                Else
                    MsgBox "Promocion solo aplica batidos Premium"
                    resetearpromocion (codi)
                End If
                
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion solo aplica a  batidos Medianos"
                resetearpromocion (codi)
            End If
        'Else
        '    'no deja cambiar promocion
        '    resetearpromocion (codi)
        '    MsgBox "Promocion aplica a numero par de productos"
        'End If
    Case 142 '3. 2 BATIDOS CON PROTEINA mediano x 179
        If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                Select Case grupy
                    Case 24
                        ' APlica normal
                    Case Else ' promocion aplicado solo a grandes
                        MsgBox "Promocion solo aplica a  batidos con Proteina"
                        resetearpromocion (codi)
                End Select
    
            Else
                MsgBox "Promocion solo aplica a  batidos Mediano"
                resetearpromocion (codi)
            End If

        Else
            'no deja cambiar promocion
            resetearpromocion (codi)
            MsgBox "Promocion aplica a numero par de productos"
        End If
    Case 193, 216, 238 '3. 2 BATIDOS CON PROTEINA mediano x 189
        'If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                Select Case grupy
                    Case 24
                        ' APlica normal
                    Case Else ' promocion aplicado solo a grandes
                        MsgBox "Promocion solo aplica a  batidos con Proteina"
                        resetearpromocion (codi)
                End Select
    
            Else
                MsgBox "Promocion solo aplica a  batidos Mediano"
                resetearpromocion (codi)
            End If

        'Else
        '    'no deja cambiar promocion
        '    resetearpromocion (codi)
        '    MsgBox "Promocion aplica a numero par de productos"
        'End If
        
    Case 202 '2do  batido a mitad de precio
        If grantgrup <> "Batido" And grantgrup <> "Limonada" Then 'No se esta aplicando a un batido
            MsgBox "La promocion solo aplica a batidos o limonadas de 16 o 20 onzas"
            resetearpromocion (codi)
            
        Else ' promocion aplicado solo a grandes
            ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
            ratio4 = pedidocontienepromocion(202, pedid)
            
            If ratio4 > Int(ratio3 / 2) Then
                MsgBox "No hay suficientes batidos para aplicar promocion"
                resetearpromocion (codi)
            End If
        End If

    Case 208 'Llevá 2 y el 3° al 50%
        If grantgrup <> "Batido" And grantgrup <> "Limonada" Then 'No se esta aplicando a un batido
            MsgBox "La promocion solo aplica a batidos o limonadas de 16 o 20 onzas"
            resetearpromocion (codi)
            
        Else ' promocion aplicado solo a grandes
            ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
            ratio4 = pedidocontienepromocion(208, pedid)
            
            If ratio4 > Int(ratio3 / 3) Then
                MsgBox "No hay suficientes batidos para aplicar promocion"
                resetearpromocion (codi)
            End If
        End If
        
    Case 204 '2do  batido a mitad de precio de 8am a 12pm
        
        If Hour(Time()) < 8 Or Hour(Time()) > 12 Then
            'No aplica fuera de las 11 y la 1pm
            MsgBox "Promocion solo valida entre las 8am y la 12pm"
            resetearpromocion (codi)
            
        Else
            If grantgrup <> "Batido" And grantgrup <> "Limonada" Then 'No se esta aplicando a un batido
                MsgBox "La promocion solo aplica a batidos o limonadas de 16 o 20 onzas"
                resetearpromocion (codi)
                
            Else ' promocion aplicado solo a grandes
                
                If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                    ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
                    ratio4 = pedidocontienepromocion(204, pedid)
                    
                    If ratio4 > Int(ratio3 / 2) Then
                        MsgBox "No hay suficientes batidos para aplicar promocion"
                        resetearpromocion (codi)
                    End If
                    
                Else ' promocion aplicado solo a grandes
                    MsgBox "Promocion solo aplica a  batidos Medianos"
                    resetearpromocion (codi)
                End If
                
            End If
        End If
        

        
    Case 144, 162 '4. 2 Bowl Dragon   x C$250
        If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
       
            If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Dragon" Then
                ' APlica normal
            Else
                MsgBox "Promocion solo aplica a Bowl Dragon"
                resetearpromocion (codi)
            End If
         Else
            'no deja cambiar promocion
            resetearpromocion (codi)
            MsgBox "Promocion aplica a numero par de productos"
        End If
    Case 145 '2. 2 claisoc mediano x 149
        If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                Dim grupix3 As Integer
                grupix3 = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                If grupix3 = 3 Then ' batidos clasico
                    'aplic normal
                Else
                    MsgBox "Promocion solo aplica batidos Clasicos"
                    resetearpromocion (codi)
                End If
                
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion solo aplica a  batidos Medianos"
                resetearpromocion (codi)
            End If
       Else
            'no deja cambiar promocion
            resetearpromocion (codi)
            MsgBox "Promocion aplica a numero par de productos"
        End If
    Case 192, 214, 236 '2. 2 claisoc mediano x 165
        'If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
                If grupy = 3 Then ' batidos clasico
                    'aplic normal
                Else
                    MsgBox "Promocion solo aplica batidos Clasicos"
                    resetearpromocion (codi)
                End If
                
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion solo aplica a  batidos Medianos"
                resetearpromocion (codi)
            End If
       'Else
       '     'no deja cambiar promocion
       '     resetearpromocion (codi)
       '     MsgBox "Promocion aplica a numero par de productos"
       ' End If
    Case 146 '6. 2 Waffle Especial a C$219
        If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") Like "Waffle Especial*" Then
                ' APlica normal
            Else
                MsgBox "Promocion solo aplica a Waffle Especial"
                resetearpromocion (codi)
            End If
       Else
            'no deja cambiar promocion
            resetearpromocion (codi)
            MsgBox "Promocion aplica a numero par de productos"
        End If
    Case 195, 218 '6. 2 Waffle Especial a C$219
        'If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
            If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") Like "Waffle Especial*" Then
                ' APlica normal
            Else
                MsgBox "Promocion solo aplica a Waffle Especial"
                resetearpromocion (codi)
            End If
       'Else
       '     'no deja cambiar promocion
       '     resetearpromocion (codi)
       '     MsgBox "Promocion aplica a numero par de productos"
       ' End If
    Case 140 'Agranda Batido Club nueva version todo permitido
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            ' APLica Normal
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion solo aplica a  batidos grandes que se preparara"
            resetearpromocion (codi)
        End If
    Case 191, 220, 234 'Agranda Batido Club nueva version todo permitido
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            ' APLica Normal
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion solo aplica a  batidos grandes que se preparara"
            resetearpromocion (codi)
        End If
        
    Case 147, 221, 235 'Agranda tu batido x Puntos
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            ' APLica Normal
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion solo aplica a  batidos grandes"
            resetearpromocion (codi)
        End If
    Case 148 ' Aplica cuando hay combo de batido + masrca pitaua
        Dim facsemi As Integer
        Dim facbati As Integer
        Dim canjesemi As Integer
        Dim canjebati As Integer
        Dim disponible As Integer
        
        
        Select Case grupy
        Case 7 ' mostrador
            Dim subgrupf As Integer
            subgrupf = DLookup("[CodSubGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
            
            If subgrupf = 1 Or subgrupf = 2 Then ' aplica smeillas pequenas y grandes
                
                facsemi = cantidadsemillapedido(pedid)
                facbati = cantidadbatidospedido(pedid)
                If facsemi < facbati Then
                    disponible = facsemi
                Else
                    disponible = facbati
                End If
                
                canjesemi = cantidadpromocionescanjeadasensemilla(pedid, 148)
                
                If canjesemi > disponible Then
                    MsgBox "Excedio la cantidad de descuentos a las semillas por combos comprados"
                    resetearpromocion (codi)
                Else
                    'cumple hay disponibel todavia
                End If
            Else
                MsgBox "Promocion solo aplica a  Semillas pequenas y grandes"
                resetearpromocion (codi)
            End If
        Case 1, 2, 3, 4, 8, 16, 24 'Batidos
            facsemi = cantidadsemillapedido(pedid)
            facbati = cantidadbatidospedido(pedid)
            If facsemi < facbati Then
                disponible = facsemi
            Else
                disponible = facbati
            End If

            canjebati = cantidadpromocionescanjeadasenbatido(pedid, 148)
            
            If canjebati > disponible Then
                MsgBox "Excedio la cantidad de descuentos a los batidos por combos comprados"
                resetearpromocion (codi)
            Else
                'cumple hay disponibel todavia
            End If
                
        Case Else ' Intentan colocar aplicacion a no batido
            MsgBox "Promocion solo aplica a  batidos y semillas"
            resetearpromocion (codi)
        End Select
        
    Case 156 '2x1 cuando es cons ervicio de motorizado propio
        
        If grantgrup <> "Batido" Then 'No se esta aplicando a un batido
            MsgBox "La promocion solo aplica a batidos de 16 o 20 onzas"
            resetearpromocion (codi)
        End If
        
        If pedidocontieneproducto("Servicio Delivery", pedid) = 0 Then ' no tiene servicio delivery facturado
            MsgBox "Se tiene que facturar el servicio de delivery para aplicar a la promocion, Promocion solo disponible para servicio de delivery pitaya"
            resetearpromocion (codi)
        End If
                
    Case 137 '2x1 Inauguracion Masaya
        If grupy = 1 Or grupy = 2 Or grupy = 3 Or grupy = 8 Or grupy = 16 Or grupy = 24 Then 'Batidos

            If cantiped = 1 Then  'SOlo se aplica a 1 producto el segundo
        
            Else ' La promocion solo se aplica a 1 batido
                MsgBox "La promocion solo se hace el descuento a 1 producto , el segundo batido"
                resetearpromocion (codi)
        
            End If
        Else ' Intentan colocar aplicacion a no batido
                MsgBox "La prmoocion solo aplica a batidos"
                resetearpromocion (codi)
        End If
        
    Case 261 '10% coincidencia de apellido
        If grupy = 1 Or grupy = 2 Or grupy = 3 Or grupy = 8 Or grupy = 16 Or grupy = 24 Or grupy = 6 Or grupy = 14 Then  'Batidos Bowl y Waffle
            Dim nombreclico As String
            nombreclico = DatosClubHost(ocluby, codigoLocal())(5) 'Datos de clinte con apellido
            
            ' --- Verificacion de apellido por dia ---
            Dim fechaHoy As Integer
            Dim apellidosDia As String
            Dim nombreNorm As String
            Dim apellidosArr() As String
            Dim I As Integer
            Dim coincide As Boolean
            
            fechaHoy = Day(Date)
            
            Select Case fechaHoy
                Case 23: apellidosDia = "LOPEZ,MARTINEZ,RODRIGUEZ,GARCIA,PEREZ"
                Case 24: apellidosDia = "HERNANDEZ,SANCHEZ,MORALES,RUIZ,GOMEZ"
                Case 25: apellidosDia = "MENDOZA,GONZALEZ,TORREZ,DIAZ,FLORES"
                Case 26: apellidosDia = "ESPINOZA,GUTIERREZ,CASTILLO,REYES,URBINA"
                Case 27: apellidosDia = "CRUZ,MORAGA,LACAYO,MOLINA,MEJIA,VARGAS"
                Case 28: apellidosDia = "PALACIOS,ESTRADA,SILVA,SOLIS,VASQUEZ"
                Case Else: apellidosDia = ""
            End Select
            
            If apellidosDia = "" Then
                MsgBox "La promocion no aplica hoy"
                resetearpromocion (codi)
            Else
                ' Normalizar nombre del cliente (quitar tildes, mayusculas)
                nombreNorm = UCase(nombreclico)
                nombreNorm = Replace(nombreNorm, "A", "A") ' placeholder para estructura
                
                ' Reemplazar vocales con tilde
                Dim acentos(9, 1) As String
                acentos(0, 0) = Chr(193): acentos(0, 1) = "A"  ' Á
                acentos(1, 0) = Chr(201): acentos(1, 1) = "E"  ' É
                acentos(2, 0) = Chr(205): acentos(2, 1) = "I"  ' Í
                acentos(3, 0) = Chr(211): acentos(3, 1) = "O"  ' Ó
                acentos(4, 0) = Chr(218): acentos(4, 1) = "U"  ' Ú
                acentos(5, 0) = Chr(225): acentos(5, 1) = "A"  ' á
                acentos(6, 0) = Chr(233): acentos(6, 1) = "E"  ' é
                acentos(7, 0) = Chr(237): acentos(7, 1) = "I"  ' í
                acentos(8, 0) = Chr(243): acentos(8, 1) = "O"  ' ó
                acentos(9, 0) = Chr(250): acentos(9, 1) = "U"  ' ú
                
                Dim j As Integer
                For j = 0 To 9
                    nombreNorm = Replace(nombreNorm, acentos(j, 0), acentos(j, 1))
                Next j
                
                ' Revisar si alguna palabra del nombre coincide con los apellidos del dia
                apellidosArr = Split(apellidosDia, ",")
                coincide = False
                
                Dim palabrasNombre() As String
                palabrasNombre = Split(nombreNorm, " ")
                
                For I = 0 To UBound(palabrasNombre)
                    Dim palabra As String
                    palabra = Trim(palabrasNombre(I))
                    If palabra <> "" Then
                        For j = 0 To UBound(apellidosArr)
                            If palabra = Trim(apellidosArr(j)) Then
                                coincide = True
                                Exit For
                            End If
                        Next j
                    End If
                    If coincide Then Exit For
                Next I
                
                If Not coincide Then
                    MsgBox "La promocion solo aplica a los apellidos seleccionados"
                    resetearpromocion (codi)
                End If
            End If
            ' --- Fin verificacion de apellido ---
            
        Else ' Intentan colocar aplicacion a no batido
            MsgBox "La prmoocion solo aplica a batidos, waffles y bowls"
            resetearpromocion (codi)
        End If
        
    Case 138 '2x1 Promocion Granada
        If grupy = 1 Or grupy = 2 Or grupy = 3 Or grupy = 8 Or grupy = 16 Or grupy = 24 Then 'Batidos

            If cantiped = 1 Then  'SOlo se aplica a 1 producto el segundo
        
            Else ' La promocion solo se aplica a 1 batido
                MsgBox "La promocion solo se hace el descuento a 1 producto , el segundo batido"
                resetearpromocion (codi)
        
            End If
        Else ' Intentan colocar aplicacion a no batido
                MsgBox "La prmoocion solo aplica a batidos"
                resetearpromocion (codi)
        End If
        
    Case 172 'canjeo de membresia por puntos
        
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Membresia" Then
            ' permitido facturar solo para membresia
        Else
            MsgBox "Promocion aplica solo a la membresia"
            resetearpromocion (codi)
        End If
        
        If cantiped = 1 Then  'SOlo se aplica a 1 membresi
        
        Else ' La promocion solo se aplica a 1 batido
            MsgBox "La promocion solo se canjea una membresia a la vez"
            resetearpromocion (codi)
    
        End If
        
        'If Date - ultimacompracliente(ocluby) > 90 Then
        '    MsgBox "la renovacion solo aplica si el cliente ha hecho alguna compra en los ultimos 3 meses"
        '    resetearpromocion (codi)
        'End If
    Case 196 'canjeo de membresia por 3 puntos
        
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Membresia" Then
            ' permitido facturar solo para membresia
        Else
            MsgBox "Promocion aplica solo a la membresia"
            resetearpromocion (codi)
        End If
        
        If cantiped = 1 Then  'SOlo se aplica a 1 membresi
        
        Else ' La promocion solo se aplica a 1 batido
            MsgBox "La promocion solo se canjea una membresia a la vez"
            resetearpromocion (codi)
    
        End If
        
        'If Date - ultimacompracliente(ocluby) > 90 Then
        '    MsgBox "la renovacion solo aplica si el cliente ha hecho alguna compra en los ultimos 3 meses"
        '    resetearpromocion (codi)
        'End If

    Case 169, 170, 108 'Membresia gratis a contretic y webholp
        If grupy = 5 Then 'membresia

            If cantiped = 1 Then  'SOlo se aplica a 1 membresi
        
            Else ' La promocion solo se aplica a 1 batido
                MsgBox "La promocion solo se canjea una membresia a la vez"
                resetearpromocion (codi)
        
            End If
        Else ' Intentan colocar aplicacion a no batido
                MsgBox "La prmoocion solo aplica a membresias"
                resetearpromocion (codi)
        End If
        
    Case 8 ' CUmpleanos
        If cantiped = 1 Then
            'aplica producto 1 cantidad solamente
            Dim mescum As Integer
            mescum = Month(fechacumpleclub(ocluby))
        
            'If mescum = Month(DLookup("[Fecha]", "[NotaDePedido]", "[CodPedido]=" & pedid)) Then
        
                If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then 'no aplica a grande
                    MsgBox "Promocion no aplica a tamaño grande"
                    resetearpromocion (codi)
                Else
                    Dim grupr As Integer
                    grupr = DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                    If grupr = 1 Or grupr = 3 Or grupr = 8 Or grupr = 2 Or grupr = 16 Or grupr = 24 Then ' batidos saludables, clasicos y especiales aplica y limonada
                        'Aplica directamente
                        'Call EnviarMensajeTelegram("Batido de cumpleanos canjeado por cliente " & ocluby & ": " & nombreCliente(ocluby), grupotoperaciones())
                        'Dim cano As Integer
                        'Dim canje As Integer
    
                        'cano = Year(DLookup("[Fecha]", "[NotaDePedido]", "[CodPedido]=" & pedid))
                        'canje = totalbatidogratiscumple(ocluby, cano)
                        'If canje > 0 Then 'ya canjeo batido de cumple este año
                        '    MsgBox "Ya tiene un batido aplicado este año"
                        '    DoCmd.SetWarnings False
                        '    DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 5" & _
                        '    " WHERE SubPedido.CodSubPedido = " & codi
                        '    DoCmd.SetWarnings True
                        'End If
                    Else
                        MsgBox "Promocion solo aplica a batidos saludables, clasicos, premium y especiales"
                        resetearpromocion (codi)
                    End If
                End If
            'Else
            '    MsgBox "Promocion solo aplica en mes de cumpleaños"
            '    DoCmd.SetWarnings False
            '    DoCmd.RunSQL "UPDATE SubPedido SET SubPedido.CodPromocion = 5" & _
            '    " WHERE SubPedido.CodSubPedido = " & codi
            '    DoCmd.SetWarnings True
            'End If
        Else
            MsgBox "Promocion aplica a cantidad de un solo batido"
            resetearpromocion (codi)
        End If
    Case 130 ' Agranda batido si tiene semilla en factura
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
            MsgBox "Promocion solo aplica a  batidos Normales"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a grandes
            
            ratio20 = pedidocontienesubgrupo(7, pedid)

            If ratio20 > 0 Then ' existe amrca pitaya en pedido
                'aplic normal
            Else
                MsgBox "Se debe de facturar una semilla para poder aplicar promocion"
                resetearpromocion (codi)
            End If
        End If
    Case 149 ' Agranda batido si tiene semilla en factura
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
            MsgBox "Promocion solo aplica a  batidos o limonadas Normales"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a grandes
            
            ratio200 = cantidadsubgrupodegrupopedido(pedid, 7, 1) + cantidadsubgrupodegrupopedido(pedid, 7, 2)

            If ratio200 > 0 Then ' existe amrca pitaya en pedido
                'aplic normal
            Else
                MsgBox "Se debe de facturar un snack marca pitaya para poder aplicar promocion"
                resetearpromocion (codi)
            End If
        End If
    Case 163 ' Agranda batido si hay productos de mosgtrador no semilla
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
            MsgBox "Promocion solo aplica a  batidos o limonadas Normales"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a grandes
            
            ratio200 = cantidadsubgrupodegrupopedido(pedid, 7, 0)

            If ratio200 > 1 Then ' existe ptro producto de mostrador 2 a mas
                'aplic normal
            Else
                MsgBox "Se debe de facturar 2 productos de mostrador que no sean marca pitaya para poder aplciar la promocion"
                resetearpromocion (codi)
            End If
        End If
    Case 164 ' Agranda batido x compra de 3 a mas batidos y liminadas
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then
            MsgBox "Promocion solo aplica a  batidos o limonadas Normales"
            resetearpromocion (codi)
        Else ' promocion aplicado solo a grandes
            
            ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
            ratio4 = pedidocontienepromocion(164, pedid)
            
            If ratio4 > Int(ratio3 / 3) Then
                MsgBox "No hay suficientes batidos para aplicar promocion"
                resetearpromocion (codi)
            End If
        End If
    Case 166 ' pro compra de 4 snakc uno es gratis
        
        If grupy = 7 And subgrup = 1 Then
            ratio3 = cantidadsubgrupodegrupopedido(pedid, 7, 1) + cantidadsubgrupodegrupopedido(pedid, 7, 2)
            ratio4 = pedidocontienepromocion(166, pedid)
            
            If ratio4 > Int(ratio3 / 4) Then
                MsgBox "No hay suficientes productos para aplicar promocion"
                resetearpromocion (codi)
            End If
        Else ' promocion aplicado solo a snack marca pitaya
            MsgBox "Promocion solo aplica a semillas marca pitaya de menor precio"
            resetearpromocion (codi)
        End If
    Case 165 '20% descuento a una granola
    
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") Like "*Granola Grande Pitaya*" Or DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") Like "*Cocoa en Polvo Grande Pitaya*" Then
            ratio3 = pedidocontienegrupo("Bowl", pedid) + pedidocontienegrupo("Waffles", pedid)
            
            If ratio3 < 1 Then
                MsgBox "No hay suficientes waffles o bowls  para aplicar promocion"
                resetearpromocion (codi)
            End If
           
        Else
            ' APlica solo granola
            MsgBox "No aplica a otro producto, solo aplica a Granola Grande Marca Pitaya o Cocoa Grande Marca PItaya"
            resetearpromocion (codi)
        End If
        
    Case 132 ' Dos Galletas de avena a 30 por compra de batido grande
        If pedidocontienegrupotamano(1, "Gigantona", pedid) + pedidocontienegrupotamano(16, "Gigantona", pedid) + pedidocontienegrupotamano(2, "Gigantona", pedid) + pedidocontienegrupotamano(3, "Gigantona", pedid) + pedidocontienegrupotamano(8, "Gigantona", pedid) > 0 Then
            
            If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Galleta de Avena" Then
                If cantiped = 2 Then ' solo se fatura 2
                    'cumple con condicion
                Else
                    MsgBox "Debe ingresar solamente 2 galletas que aplicaran la promocion"
                    resetearpromocion (codi)
                End If
            Else ' promocion aplicado solo a grandes
                MsgBox "Promocion no aplica a productos que no sean galleta"
                resetearpromocion (codi)
            End If
        Else ' promocion aplicado solo si hay grandes
            MsgBox "Promocion solo aplica si se ha agregado un batido de 20oz en la factura"
            resetearpromocion (codi)
        End If
    Case 129, 223 ' Descuento x compra de galletas
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Galleta de Avena" Then
            If cantiped / 3 - Round(cantiped / 3, 1) = 0 Then ' multipli de 3
                'cumple con condicion
            Else
                MsgBox "Debe ingresar una cantidad de galletas multiplo de 3"
                resetearpromocion (codi)
            End If
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion no aplica a productos que no sean galleta"
            resetearpromocion (codi)
        End If
    Case 197, 219 ' Descuento x compra de galletas 4X3 CLUB
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Galleta de Avena" Then
            If cantiped / 4 - Round(cantiped / 4, 1) = 0 Then ' multipli de 4
                'cumple con condicion
            Else
                MsgBox "Debe ingresar una cantidad de galletas multiplo de 4"
                resetearpromocion (codi)
            End If
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion no aplica a productos que no sean galleta"
            resetearpromocion (codi)
        End If
    Case 128, 222 ' Descuento x compra de galletas
        If DLookup("[Nombre]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Galleta de Avena" Then
            If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then ' multipli de 2
                'cumple con condicion
            Else
                MsgBox "Debe ingresar una cantidad de galletas multiplo de 2"
                resetearpromocion (codi)
            End If
        Else ' promocion aplicado solo a grandes
            MsgBox "Promocion no aplica a productos que no sean galleta"
            resetearpromocion (codi)
        End If

    Case 184 ' waffle batido clasio al 50%
        If grupy = 3 Then  ' se aplica a un batido clasico
            ' se aplica a un batyido de 16 oz
        Else
            MsgBox "El decuento solamente se aplica a el batido clasico"
            resetearpromocion (codi)
        End If
        ratio3 = pedidocontienesubgrupo(14, pedid)
        ratio4 = pedidocontienepromocion(184, pedid)
        
        If ratio3 < ratio4 Then
            MsgBox "No hay suficientes Waffles para aplicar promocion"
            resetearpromocion (codi)
        End If
        
        
    Case 159 'Combo waffle + batido, se aplica a abtido , busca si hay waffle claisco
        If grupy = 3 And DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then ' se aplica a un batido clasico de 16 oz procede
            ' se aplica a un batyido de 16 oz
        Else
            MsgBox "El decuento solamente se aplica a el batido de 16 onzas"
            resetearpromocion (codi)
        End If
    
        If pedidocontieneproducto("Waffle Clasico c/ Miel", pedid) <> 0 Or pedidocontieneproducto("Waffle Clasico c/ L Condensada", pedid) <> 0 Or pedidocontieneproducto("Waffle Clasico c/ Chocolate", pedid) <> 0 Then
            'hay un waffle clasico de cualquiera de los 3 tipos en la factura
        Else
            MsgBox "La promocion aplica si hay un waffle clasico facturado"
            resetearpromocion (codi)
        End If
        
    Case 160
        If grantgrup = "Batido" Or grantgrup = "Limonada" Then  ' se aplica a un batido
            ' se aplica a un batyido
        Else
            MsgBox "El decuento solamente se aplica a el batido de 16 o de 20 onzas"
            resetearpromocion (codi)
        End If
        
        ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
        ratio4 = pedidocontienepromocion(160, pedid)
        
        If ratio4 > Int(ratio3 / 4) Then
            MsgBox "No hay suficientes batidos para aplicar promocion"
            resetearpromocion (codi)
        End If
        
    Case 203
        If Hour(Time()) < 12 Or Hour(Time()) > 13 Then
            'No aplica fuera de las 11 y la 1pm
            MsgBox "Promocion solo valida entre las 12pm y la 1pm"
            resetearpromocion (codi)
        Else
            If grantgrup = "Batido" Or grantgrup = "Limonada" Then  ' se aplica a un batido
                ' se aplica a un batyido
            Else
                MsgBox "El decuento solamente se aplica a el batido de 16 o de 20 onzas"
                resetearpromocion (codi)
            End If
            
            ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
            ratio4 = pedidocontienepromocion(203, pedid)
            
            If ratio4 > Int(ratio3 / 4) Then
                MsgBox "No hay suficientes batidos para aplicar promocion"
                resetearpromocion (codi)
            End If
        End If

    Case 161  ' 2x1 happy hour granada
        
        If Hour(Time()) > 12 Or Hour(Time()) < 11 Then
            'No aplica fuera de las 11 y la 1pm
            MsgBox "Promocion solo valida entre las 11am y la 1pm"
            resetearpromocion (codi)
        End If
        
        If grantgrup = "Batido" Or grantgrup = "Limonada" Then  ' se aplica a un batido
            ' se aplica a un batyido
        Else
            MsgBox "El decuento solamente se aplica a el batido y limonadas de 16 o de 20 onzas"
            resetearpromocion (codi)
        End If
        
        ratio3 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("Limonada", pedid)
        ratio4 = pedidocontienepromocion(161, pedid)
        
        If ratio4 > Int(ratio3 / 2) Then
            MsgBox "No hay suficientes batidos para aplicar promocion"
            resetearpromocion (codi)
        End If
        
    Case 265  ' Inauguracion Villa fotnana 2x1
        
        If grantgrup = "Batido" Or grantgrup = "Limonada" Then  ' se aplica a un batido
            ' se aplica a un batyido
        Else
            MsgBox "El decuento solamente se aplica a el batido y limonadas de 20 onzas"
            resetearpromocion (codi)
        End If
        
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            'Si aplica a grande
        Else
            MsgBox "Promocion solo aplica a a batidos o limonadas normales"
            resetearpromocion (codi)
        End If
        
    Case 267  ' 2 Batidos Premium x c$250
        
        If grupy = 2 Then  ' se aplica a un batido premium
            ' se aplica a un batyido
        Else
            MsgBox "El decuento solamente se aplica a batidos premium"
            resetearpromocion (codi)
        End If
        
        If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Gigantona" Then
            'Si aplica a grande
        Else
            MsgBox "Promocion solo aplica a batidos 20 onzas"
            resetearpromocion (codi)
        End If
        
    Case 268  ' 2 Waffles Especiales x c$270

        If grupy = 14 Then ' se aplica a un waffle
            ' se aplica a un batyido
        Else
            MsgBox "El decuento solamente se aplica a waffles"
            resetearpromocion (codi)
        End If

        If nombrep Like "Waffle Especial*" Then
            ' APlica solo waffle especial
           
        Else
            MsgBox "El decuento solamente se aplica a waffles Especiales"
            resetearpromocion (codi)
            
        End If
        
    Case 173 '14% descuento proutos mismo grupo
        Dim ratio1731, ratio1732, ratio1733, ratio1734 As Integer
        Dim ratio1735, ratio1736, ratio1737, ratio1738 As Double
                                
        Select Case grantgrup
            Case "Pitaya Store"
                If subgrup = 1 Or subgrup = 2 Then
                    ratio1731 = cantidadsubgrupodegrupopedido(pedid, 7, 1) + cantidadsubgrupodegrupopedido(pedid, 7, 2)
                    ratio1735 = pedidocontienepromocionxsubgrupo(173, pedid, 7, 1) + pedidocontienepromocionxsubgrupo(173, pedid, 7, 2)
                    If ratio1735 > 2 * Int(ratio1731 / 2) Then
                        MsgBox "No hay suficientes productos para aplicar promocion"
                        resetearpromocion (codi)
                    End If
                Else
                    MsgBox "Promocion solo aplica a semillas marca pitaya, batidos, limonadas, waffles o bowls"
                    resetearpromocion (codi)
                End If
                
            Case "Batido", "Limonada"
                ratio1732 = pedidocontienegrupo("Batido", pedid) + pedidocontienegrupo("limonada", pedid)
                ratio1736 = pedidocontienepromocionxtipo(173, pedid, "Batido") + pedidocontienepromocionxtipo(173, pedid, "limonada")
                If ratio1736 > 2 * Int(ratio1732 / 2) Then
                    MsgBox "No hay suficientes productos para aplicar promocion"
                    resetearpromocion (codi)
                End If
                
            Case "Bowl"
                ratio1733 = pedidocontienegrupo("Bowl", pedid)
                ratio1737 = pedidocontienepromocionxtipo(173, pedid, "Bowl")
                If ratio1737 > 2 * Int(ratio1733 / 2) Then
                    MsgBox "No hay suficientes productos para aplicar promocion"
                    resetearpromocion (codi)
                End If
                
            Case "Waffles"
                ratio1734 = pedidocontienegrupo("Waffles", pedid)
                ratio1738 = pedidocontienepromocionxtipo(173, pedid, "Waffles")
                If ratio1738 > 2 * Int(ratio1734 / 2) Then
                    MsgBox "No hay suficientes productos para aplicar promocion"
                    resetearpromocion (codi)
                End If
                
            Case Else
                MsgBox "Promocion solo aplica a semillas marca pitaya, batidos, limonadas, waffles o bowls"
                resetearpromocion (codi)
        End Select
    Case 174 ' 2 membresias al 50%
        Dim ratio1741, ratio1742 As Integer
        ratio1741 = pedidocontienegrupo("Membresia", pedid)
        ratio1742 = pedidocontienepromocionxtipo(174, pedid, "Membresia")
        If ratio1742 > 2 * Int(ratio1741 / 2) Then ' no es mltiplo de 2
            MsgBox "No hay suficientes productos para aplicar promocion"
            resetearpromocion (codi)
        End If
        
        If nombrep <> "Membresia" Then ' no es membresia
            MsgBox "Promocion solo aplica a membresias"
            resetearpromocion (codi)
        End If
        
    Case 175
        'aplica solo TripleBerry, Leon , Bowl Dragon
        If nombrep = "Triple Berry" Or nombrep = "Leon" Or nombrep = "Dragon" Then
            ' APlica solo 3 productos
        Else
            MsgBox "Promocion solo aplica a Triple Berry, Leon o Dragon"
            resetearpromocion (codi)
        End If
        
    Case 186
        'aplica solo Granada y Waffle Clasico con miel
        If nombrep = "Granada" Then
            If DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "Mediano" Then 'no aplica a grande
                ' cumple con ser batido mediano
            Else
                MsgBox "No aplica tamaño grande"
                resetearpromocion (codi)
            End If
        ElseIf nombrep = "Waffle Clasico c/ Miel" Then
            ' APlica solo wafle clasico
        Else
            MsgBox "Promocion solo aplica a Batido granada de 16oz y waffle clasico con miel"
            resetearpromocion (codi)
        End If
        
    Case 187
        'aplica solo Bowl Dragon y Batido Leon
        If nombrep = "Dragon" Then

        ElseIf nombrep = "Leon" Then
            ' APlica solo wafle clasico
        Else
            MsgBox "Promocion solo aplica a Batido Leon  y Bowl Dragon"
            resetearpromocion (codi)
        End If
    Case 92 ' cambio de ingrediente no se puede aplicar a anda
        resetearpromocion (codi)
    Case 22
        If cpunt < 0 Then 'despues de descontar queda negativo
            MsgBox "No cuenta con suficientes puntos para canjear"
            resetearpromocion (codi)
        Else
            
            Select Case grupy
                
                Case 1, 2, 3, 4, 8, 24 ' Batidos - ambos tamanos aplican desde Oct 2026
                    Dim medBat As String
                    medBat = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                    If medBat = "Mediano" Or medBat = "Gigantona" Then
                        ' cumple - Mediano (16oz) y Gigantona (20oz) aplican
                    Else
                        MsgBox "No aplica a este tamaño de batido"
                        resetearpromocion (codi)
                    End If

                Case 16 ' Limonadas - nuevo en canjeo desde Oct 2026
                    Dim medLim As String
                    medLim = DLookup("[Medida]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                    If medLim = "Mediano" Or medLim = "Gigantona" Then
                        ' cumple - ambos tamanos aplican
                    Else
                        MsgBox "No aplica a este tamaño de limonada"
                        resetearpromocion (codi)
                    End If

                Case 6 ' Energy Bowl - Acai ahora SI aplica desde Oct 2026
                    ' Todos los bowls aplican. El costo diferenciado lo calcula PuntosClubProducto

                Case 14 ' Waffles - Clasico, Especial y Proteina
                    ' Todos los Waffles aplican. El costo diferenciado lo calcula PuntosClubProducto
                    
                Case 7 ' Pitaya Store
                    If DLookup("[Marca]", "[DBBatidos]", "[CodBatido]='" & batix & "'") = "PitayaStore" Then 'no aplica a no marca pitaya
                        ' APlica a Pitaya Store
                        Dim codsubg As Integer
                        codsubg = DLookup("[CodSubGrupo]", "[DBBatidos]", "[CodBatido]='" & batix & "'")
                        
                        Select Case codsubg
                            Case 1 ' Semillas pequenas (Almendras/Pistachos = 8pts) y mixes (Maranon/Mix* = 12pts)
                                ' APlica - el costo diferenciado lo calcula PuntosClubProducto

                            Case 2 ' Frascos grandes (Granola/Cocoa = 21pts, Semilla de Cacao = 29pts)
                                ' APlica - el costo diferenciado lo calcula PuntosClubProducto

                            Case 3 ' Galletas de Avena - requiere cantidad par (cada 2 por 5 puntos)
                                If cantiped / 2 - Round(cantiped / 2, 0) = 0 Then '' cantidad par
                                    ' aplica promocion porque es cantidad par
                                Else
                                    MsgBox "La Promocion aplica a un numero par de galletas de avena, Cada 2 por 5 puntos"
                                    resetearpromocion (codi)
                                End If
                            Case Else
                                MsgBox "No aplica productos de Marca Pitaya que no sean semillas, mixes, frascos o galletas"
                                resetearpromocion (codi)
                        End Select
                        
                    Else
                        MsgBox "No aplica productos de mostrador que no sean Marca Pitaya"
                        resetearpromocion (codi)
                    End If
                Case Else
                    MsgBox "No aplica para este producto"
                    resetearpromocion (codi)

            End Select
            
            If [Forms]![Nota de Pedido].VALIDACIONCEDULA = False Then
                'despues de maquinar todo lo de los puntos si no presenta cedula no se aplica
                If APIDisponible() Then
                    DoCmd.OpenForm "ValidacionCedulaClub"
                    [Forms]![ValidacionCedulaClub].codigoclub = [Forms]![Nota de Pedido].CodCliente
                    [Forms]![ValidacionCedulaClub].codigopedido = [Forms]![Nota de Pedido].CodPedido
                End If
            End If
            
        End If
                  
End Select
[Forms]![Nota de Pedido].Form.Requery
End Sub




Private Sub CodPromocion_Enter()
If Me.Vinculo <> 0 Then  ' es un producto vinculado, no permite cambio de promocion
    Me.nombreorden.SetFocus
End If

If Me.CodBatido Like "combo*" Then  ' es un producto combo, se aplicara desde 3/5 que lleve diferencia de precio por romocion
    Me.nombreorden.SetFocus
End If
End Sub

Private Sub CodPromocion_Exit(Cancel As Integer)
[Forms]![Nota de Pedido].Form.ppedido.Requery
[Forms]![Nota de Pedido].Form.Texto193.Requery
End Sub



Private Sub CodPromocion_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Permite solo teclas de navegación y eliminación

    KeyCode = 0

    'MsgBox "No se puede ingresar promociones por texto, se tiene que elegir en la lista desplegable"
    'Me.CodPromocion = 5
End Sub

Private Sub Comando177_Click()
If DLookup("[Editable]", "[Grupos]", "[CodGrupo]=" & DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.CodBatido & "'")) <> -1 Then
    'No abre nada porque no es editable
    Exit Sub
End If

'If Me.Vinculo = 0 Or IsNull(Me.Vinculo) Then
    
    DoCmd.OpenForm "Menu Adicionales"
    [Forms]![Menu Adicionales]![acodbatido] = Me.CodBatido
    [Forms]![Menu Adicionales]![apedido] = Me.CodPedido
    [Forms]![Menu Adicionales]![asubpedido] = Me.CodSubPedido
    [Forms]![Menu Adicionales].Form.Requery
    
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Menu Adicionales].Form.SetFocus
    
'End If
End Sub

Private Sub Comando202_Click()

If Me.Vinculo = 0 Or IsNull(Me.Vinculo) Then
    
    DoCmd.OpenForm "Cambio Tamaño", , , "CodSubPedido=" & Me.CodSubPedido
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![Cambio Tamaño].SetFocus
End If
End Sub

Private Sub Comando2050_Click()
If DLookup("[Editable]", "[Grupos]", "[CodGrupo]=" & DLookup("[CodGrupo]", "[DBBatidos]", "[CodBatido]='" & Me.CodBatido & "'")) <> -1 Then
    'No abre nada porque no es editable
    Exit Sub
End If
'If Me.Vinculo = 0 Or IsNull(Me.Vinculo) Then
    
    DoCmd.OpenForm "CambiarIngredientesFactura"
    [Forms]![CambiarIngredientesFactura]![acodbatido] = Me.CodBatido
    [Forms]![CambiarIngredientesFactura]![anombre] = Me.Nombre & " " & Me.Medida
    [Forms]![CambiarIngredientesFactura]![apedido] = Me.CodPedido
    [Forms]![CambiarIngredientesFactura]![asubpedido] = Me.CodSubPedido
    [Forms]![CambiarIngredientesFactura]![cantidadproductos] = Me.Cantidad
    [Forms]![CambiarIngredientesFactura].Form.Requery
    
    [Forms]![Nota de Pedido].Form.Requery
    [Forms]![CambiarIngredientesFactura].SetFocus
'End If

End Sub



Private Sub eliminar_Click()
Dim raiz As Long
Dim pedi As Long
Dim mensajeg As String
Dim itemeliminado As String
Dim operariodeitemeliminado As Integer

pedi = Me.CodPedido
itemeliminado = Me.CodBatido
If CurrentProject.AllForms("Main Pitaya").IsLoaded Then
    operariodeitemeliminado = [Forms]![Main Pitaya].codigologin
Else
    operariodeitemeliminado = 5
End If
        
If IsNull(Me.Vinculo) Or Me.Vinculo = 0 Then
    If MsgBox("Desea eliminar el producto principal y todos sus productos anexados", vbYesNo, "Eliminar Producto") = vbYes Then
        raiz = Me.CodSubPedido
        
        Call resetearpromocionespedidocompleto(pedi)
        [Forms]![Nota de Pedido].Form.ppedido.Requery
        [Forms]![Nota de Pedido].Form.Texto193.Requery
        
        On Error GoTo ErrBorrado
        Call EliminarSubPedidoRecursivo(raiz, pedi)
        On Error GoTo 0
        'DoCmd.SetWarnings False
        'DoCmd.RunSQL "DELETE * FROM SubPedido WHERE Vinculo =" & raiz & " AND CodPedido =" & pedi
        'DoCmd.RunSQL "DELETE * FROM SubPedido WHERE CodSubpedido =" & raiz & " AND CodPedido =" & pedi
        'DoCmd.SetWarnings True
        
        Me.Requery
        [Forms]![Nota de Pedido]![ProductosEliminados] = -1
        
        'mensajeg = "Producto eliminado: " & nombreproductoventa(itemeliminado) & vbCrLf & _
        "Realizado por: " & NombreOperario(operariodeitemeliminado) & vbCrLf & _
        "Pedido: " & pedi
        'Call EnviarMensajeTelegram(mensajeg, grupotanulaciones())
    Else
        ' no se elmina nada
    End If

ElseIf Me.Vinculo <> 0 Then 'producto vinculado no aplica
    MsgBox "No se puede elimnar un subproducto, elimina el producto principal para que se eliminen todos sus elementos"
    'If MsgBox("No se puede eliminar un producto anexado, se eliminara el producto principal y todos sus productos anexados", vbYesNo, "Eliminar Producto") = vbYes Then
    '    raiz = Me.Vinculo
    '
    '    Call resetearpromocionespedidocompleto(pedi)
    '    [Forms]![Nota de Pedido].Form.ppedido.Requery
    '    [Forms]![Nota de Pedido].Form.Texto193.Requery
    '
    '    DoCmd.SetWarnings False
    '    DoCmd.RunSQL "DELETE * FROM SubPedido WHERE Vinculo =" & raiz & " AND CodPedido =" & pedi
    '    DoCmd.RunSQL "DELETE * FROM SubPedido WHERE CodSubpedido =" & raiz & " AND CodPedido =" & pedi
    '    DoCmd.SetWarnings True
    '    Me.Requery
    '    [Forms]![Nota de Pedido]![ProductosEliminados] = -1
    '
    '    'mensajeg = "Producto eliminado: " & nombreproductoventa(itemeliminado) & vbCrLf & vbCrLf & _
    '    "Realizado por: " & NombreOperario(operariodeitemeliminado) & vbCrLf & vbCrLf & _
    '    "Pedido: " & pedi
    '    'Call EnviarMensajeTelegram(mensajeg, grupotanulaciones())
    '
    'Else
    '    'no s elmino nada
    'End If
    
Else
    MsgBox "No se puede eliminar producto, reportar a lider"
End If

[Forms]![Nota de Pedido].Form.Secundario57.Requery
Exit Sub

ErrBorrado:
    MsgBox "Error al eliminar: " & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub empaquetexto_Click()
'If Me.Vinculo <> 0 Then ' es un producto anexo, no genera cambios
'    Exit Sub
'End If

If Me.Empaque <> 0 Then
    Me.Empaque = 0
    Me.empaquetexto.Requery
    Me.nombreorden.SetFocus
Else
    Me.Empaque = -1
    Me.empaquetexto.Requery
    Me.nombreorden.SetFocus
End If
End Sub


Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.eliminar.Picture = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Sys Resources\Icons\Eliminar.png"
End Sub


