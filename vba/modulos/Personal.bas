' ==========================================================
' Modulo  : Personal
' Tipo    : 1  |  Lineas: 1211
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Function TarifaDia(Ope As Integer, fech As Date) As Double
'Tarifa de un dia especifico de un operario especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Halla la tarifa de dicho dia
miSQL = "SELECT TarifaPersonal.CodOperario, TarifaPersonal.FechaInicial, TarifaPersonal.FechaFinal, TarifaPersonal.Tarifa FROM TarifaPersonal WHERE (((TarifaPersonal.CodOperario)=" & Ope & ") AND ((TarifaPersonal.FechaInicial)<=#" & fech & "#) AND ((TarifaPersonal.FechaFinal)>=#" & fech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
TarifaDia = rst("Tarifa")
rst.Close

Exit Function

Nulo:
TarifaDia = 0

End Function


Function NombreOperario(cod As Integer) As String
'Nombre y Apellido del operario

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Nombre y Apellido del operario segun codigo
miSQL = "SELECT [Operarios]![Nombre] & ' ' & [Operarios]![Apellido] AS Name," & _
" Operarios.CodOperario FROM Operarios WHERE (((Operarios.CodOperario)=" & cod & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
NombreOperario = rst("Name")
rst.Close

Exit Function

Nulo:
NombreOperario = " "

End Function

Function cargooperariooperativo(Cargo As Integer, loc As Integer, des As Date) As Integer
'codigo del operario

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Nombre y Apellido del operario operativo segun cargo
miSQL = "SELECT [AsignacionNivelesCargos]![Fecha]<=#" & des & "# And" & _
" IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date(),[AsignacionNivelesCargos]![Fin])>=#" & des & "# AS Condi," & _
" AsignacionNivelesCargos.CodNivelesCargos, AsignacionNivelesCargos.Sucursal, AsignacionNivelesCargos.CodOperario" & _
" FROM AsignacionNivelesCargos" & _
" WHERE ((([AsignacionNivelesCargos]![Fecha]<=#" & des & "# And" & _
" IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date(),[AsignacionNivelesCargos]![Fin])>=#" & des & "#)<>0)" & _
" AND ((AsignacionNivelesCargos.CodNivelesCargos)=" & Cargo & ") AND ((AsignacionNivelesCargos.Sucursal)=" & loc & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
cargooperariooperativo = rst("CodOperario")
rst.Close

Exit Function

Nulo:
cargooperariooperativo = 0

End Function

Function codigocargosegunfechaoperario(opex As Integer, fechi As Date) As Integer
'codigo del cargo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'codigo de cargo segun fecha y co operario
miSQL = "SELECT AsignacionNivelesCargos.Fecha, AsignacionNivelesCargos.CodOperario, " & _
" [AsignacionNivelesCargos]![Fecha]<=#" & fechi & "# And IIf(IsNull([AsignacionNivelesCargos]![Fin]),Date()," & _
" [AsignacionNivelesCargos]![Fin])>=#" & fechi & "# AS condi, AsignacionNivelesCargos.CodNivelesCargos" & _
" FROM Operarios INNER JOIN AsignacionNivelesCargos ON Operarios.CodOperario = AsignacionNivelesCargos.CodOperario" & _
" WHERE (((AsignacionNivelesCargos.CodOperario) = " & opex & ") And (([AsignacionNivelesCargos]![Fecha] <= #" & fechi & "# And" & _
" IIf(IsNull([AsignacionNivelesCargos]![Fin]), Date, [AsignacionNivelesCargos]![Fin]) >= #" & fechi & "#) = True))" & _
" ORDER BY AsignacionNivelesCargos.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
codigocargosegunfechaoperario = rst("CodNivelesCargos")
rst.Close

Exit Function

Nulo:
codigocargosegunfechaoperario = 2 ' Operario estandar

End Function


Function CorroborarClave(cod As Integer) As String
'Corroborar contrasena de operario

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'busca clave de operario
miSQL = "SELECT Operarios.CodOperario, Operarios.clave FROM Operarios WHERE (((Operarios.CodOperario)=" & cod & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CorroborarClave = rst("clave")
rst.Close

Exit Function

Nulo:
CorroborarClave = "NO"


End Function

Function RedondearHora(Hora As Date, C As Integer) As Double
'redondear la hora de entrada o salida acorde a creiterio de redondeo
'c = 0, entrada - c= 1 , salida

On Error GoTo Nulo
Dim minutos As Integer

'redondear hora
minutos = Minute(Hora) 'sacar minutos de hora marcada

''''''''''''''''''OTRA OPCION''''''''
Select Case C
    Case 0   'Redondear entrada
        Select Case minutos
            Case 0 To 5 'TOlerancia de un minuto
                RedondearHora = Round(CDbl(Hora) * 24, 0)
            Case 6 To 29
                RedondearHora = Round(CDbl(Hora) * 24, 0) + 0.5
            Case 30 To 35 'Tolerancia 1 minuto para turnno medio
                RedondearHora = Round(CDbl(Hora) * 24, 0) - 0.5
            Case 36 To 59
                RedondearHora = Round(CDbl(Hora) * 24, 0)
        End Select
    Case 1   'Redondear salida
        Select Case minutos
            Case 0 To 14 'si no completa la media hora no se considera
                RedondearHora = Round(CDbl(Hora) * 24, 0)
            Case 15 To 29 'despues de los 15 min ya se considera media hora mas
                RedondearHora = Round(CDbl(Hora) * 24, 0) + 0.5
            Case 30 To 44 'despues de los 45 completa la hora, antes no
                RedondearHora = Round(CDbl(Hora) * 24, 0) - 0.5
            Case 45 To 59 'si no completa l ahora no se considera
                RedondearHora = Round(CDbl(Hora) * 24, 0)
        End Select
End Select


'If minutos <= 29 Then

'    If C = 0 Then 'entrada
'        If minutos <= 1 Then
'            RedondearHora = Round(CDbl(Hora) * 24, 0)
'        Else
'            RedondearHora = Round(CDbl(Hora) * 24, 0) + 0.5
'        End If
'    Else ' salida
'        RedondearHora = Round(CDbl(Hora) * 24, 0)
'    End If
    
'ElseIf minutos < 30 Then

'RedondearHora = Round(CDbl(Hora) * 24, 0) + 0.5

'ElseIf minutos = 30 Then

'RedondearHora = Round(CDbl(Hora) * 24, 0) - 0.5

'ElseIf minutos <= 59 Then

'    If C = 0 Then ' entrada
'        If minutos < 40 Then
'           RedondearHora = Round(CDbl(Hora) * 24, 0) - 0.5
'        Else
'           RedondearHora = Round(CDbl(Hora) * 24, 0)
'        End If
'    Else    ' Salida
'        RedondearHora = Round(CDbl(Hora) * 24, 0) - 0.5
'    End If

'Else

'RedondearHora = Round(CDbl(Hora) * 24, 0)

'End If

Exit Function
Nulo:
RedondearHora = 0


End Function

Function EsFeriado(Fecha As Date) As Long
'Busca si una fecha es feriado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'busca si es feriado = 1 o si es normal =0
miSQL = "SELECT EstadoInicial.Fecha, IIf([EstadoInicial]![Feriado]<>0,1,0) AS Libre FROM EstadoInicial WHERE (((EstadoInicial.Fecha)=#" & Fecha & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
EsFeriado = rst("Libre")
rst.Close

Exit Function

Nulo:
EsFeriado = 0


End Function

Function HorasSemanalesAsignadas(sema As Integer, oper As Integer) As Long
'horas totales de la semana de una semana especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TarifaPersonal.CodOperario, numerosemana([TarifaPersonal]![FechaInicial]) AS desde," & _
" numerosemana([TarifaPersonal]![FechaFinal]) AS hasta," & _
" TarifaPersonal.HL, TarifaPersonal.HMA, TarifaPersonal.HMI, TarifaPersonal.HJ, TarifaPersonal.HV," & _
" TarifaPersonal.HS, TarifaPersonal.HD" & _
" FROM Operarios INNER JOIN TarifaPersonal ON Operarios.CodOperario = TarifaPersonal.CodOperario" & _
" WHERE (((TarifaPersonal.CodOperario)=" & oper & ") AND ((numerosemana([TarifaPersonal]![FechaInicial]))<=" & sema & ")" & _
" AND ((numerosemana([TarifaPersonal]![FechaFinal]))>=" & sema & "));"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
HorasSemanalesAsignadas = rst("HL") + rst("HMA") + rst("HMI") + rst("HJ") + rst("HV") + rst("HS") + rst("HD")
rst.Close

Exit Function

Nulo:
HorasSemanalesAsignadas = 0


End Function

Function HorasSemanalesCumplidas(sema As Integer, oper As Integer) As Long
'horas totales de la semana de una semana especifica que realizo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([RegistroHorario]![Fecha]) AS semana, RegistroHorario.CodOperario," & _
" Sum(IIf(IsNull([RegistroHorario]![Salida]) Or IsNull([RegistroHorario]![Ingreso]),0,RedondearHora([RegistroHorario]![Salida],1)-RedondearHora([RegistroHorario]![Ingreso],0))) AS Horas" & _
" FROM RegistroHorario" & _
" GROUP BY numerosemana([RegistroHorario]![Fecha]), RegistroHorario.CodOperario" & _
" HAVING (((numerosemana([RegistroHorario]![Fecha]))=" & sema & ") AND ((RegistroHorario.CodOperario)=" & oper & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
HorasSemanalesCumplidas = rst("Horas")
rst.Close

Exit Function

Nulo:
HorasSemanalesCumplidas = 0


End Function

Function HorasSemanalesCumplidasSinRedondeo(sema As Integer, oper As Integer) As Double
'horas totales de la semana de una semana especifica que realizo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([RegistroHorario]![Fecha]) AS semana, RegistroHorario.CodOperario," & _
" Sum(IIf(IsNull([RegistroHorario]![Salida]) Or IsNull([RegistroHorario]![Ingreso]),0,[RegistroHorario]![Salida]-[RegistroHorario]![Ingreso])*24) AS Horas" & _
" FROM RegistroHorario" & _
" GROUP BY numerosemana([RegistroHorario]![Fecha]), RegistroHorario.CodOperario" & _
" HAVING (((numerosemana([RegistroHorario]![Fecha]))=" & sema & ") AND ((RegistroHorario.CodOperario)=" & oper & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
HorasSemanalesCumplidasSinRedondeo = rst("Horas")
rst.Close

Exit Function

Nulo:
HorasSemanalesCumplidasSinRedondeo = 0
rst.Close

End Function

Function HorasSemanalesCumplidasSinRedondeomixed(sema As Integer, oper As Integer) As Double
'horas totales de la semana de una semana especifica que realizo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT numerosemana([RegistroHorarioMixed]![Fecha]) AS semana, RegistroHorarioMixed.CodOperario," & _
" Sum(IIf(IsNull([RegistroHorarioMixed]![Salida]) Or IsNull([RegistroHorarioMixed]![Ingreso]),0,[RegistroHorarioMixed]![Salida]-[RegistroHorarioMixed]![Ingreso])*24) AS Horas" & _
" FROM RegistroHorarioMixed" & _
" GROUP BY numerosemana([RegistroHorarioMixed]![Fecha]), RegistroHorarioMixed.CodOperario" & _
" HAVING (((numerosemana([RegistroHorarioMixed]![Fecha]))=" & sema & ") AND ((RegistroHorarioMixed.CodOperario)=" & oper & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
HorasSemanalesCumplidasSinRedondeomixed = rst("Horas")
rst.Close

Exit Function

Nulo:
HorasSemanalesCumplidasSinRedondeomixed = 0
rst.Close

End Function

Function HorasTeoricas(Ope As Integer, fech As Date) As Double
'Horas teoricas de dia especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'Halla las horas teoricas del dia especifico
miSQL = "SELECT TarifaPersonal.FechaInicial, TarifaPersonal.CodOperario," & _
" TarifaPersonal.HL, TarifaPersonal.HMA, TarifaPersonal.HMI, TarifaPersonal.HJ, TarifaPersonal.HV, TarifaPersonal.HS, TarifaPersonal.HD" & _
" FROM TarifaPersonal WHERE (((TarifaPersonal.FechaInicial) <= #" & fech & "#)" & _
" And ((TarifaPersonal.CodOperario) = " & Ope & ")) ORDER BY TarifaPersonal.FechaInicial DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)

If Weekday(fech, 2) = 1 Then
    HorasTeoricas = rst("HL")
    rst.Close
End If
If Weekday(fech, 2) = 2 Then
    HorasTeoricas = rst("HMA")
    rst.Close
End If
If Weekday(fech, 2) = 3 Then
    HorasTeoricas = rst("HMI")
    rst.Close
End If
If Weekday(fech, 2) = 4 Then
    HorasTeoricas = rst("HJ")
    rst.Close
End If
If Weekday(fech, 2) = 5 Then
    HorasTeoricas = rst("HV")
    rst.Close
End If
If Weekday(fech, 2) = 6 Then
    HorasTeoricas = rst("HS")
    rst.Close
End If
If Weekday(fech, 2) = 7 Then
    HorasTeoricas = rst("HD")
    rst.Close
End If

Exit Function

Nulo:
HorasTeoricas = 0
rst.Close

End Function

Function SalarioQuincenaFijo(Fecha As Date, Ope As Integer) As Long
'Busca sel salario de quincena de un operario en una fecha especifica

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'busca el codigo de operario y la fecha y el salario fisico asignado para esa fecha
miSQL = "SELECT TarifaPersonal.FechaInicial, TarifaPersonal.CodOperario, TarifaPersonal.SalarioFijo FROM TarifaPersonal WHERE (((TarifaPersonal.FechaInicial) <= #" & Fecha & "#) And ((TarifaPersonal.CodOperario) = " & Ope & ")) ORDER BY TarifaPersonal.FechaInicial DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SalarioQuincenaFijo = rst("SalarioFijo") / 2
rst.Close

Exit Function

Nulo:
SalarioQuincenaFijo = 0


End Function

Function HorasRegulares(desde As Date, hasta As Date, Ope As Integer) As Double
'Salario regular de una quincena de un operario acorde a fecha inicial

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim Fin, fech As Date

HorasRegulares = 0

For fech = desde To hasta
HorasRegulares = HorasRegulares + HorasTeoricas(Ope, fech)
Next fech


Exit Function

Nulo:
HorasRegulares = 0


End Function

Function VacacionesFecha(Fecha As Date, Ope As Integer) As Long
'Dias de vaciaciones ne fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'DIferencia de dias feriados en un fecha
miSQL = "SELECT Vacaciones.Inicio, Vacaciones.Final, Vacaciones.CodOperario, 1 AS Cantidad FROM Vacaciones WHERE (((Vacaciones.Inicio)<=#" & Fecha & "#) AND ((Vacaciones.Final)>=#" & Fecha & "#) AND ((Vacaciones.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
VacacionesFecha = rst("Cantidad")
rst.Close

Exit Function

Nulo:
VacacionesFecha = 0


End Function

Function SubsidiosFecha(Fecha As Date, Ope As Integer) As Long
'Dias de subsidios ne fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'DIferencia de dias SubsidioPersonal en un fecha
miSQL = "SELECT SubsidioPersonal.Inicio, SubsidioPersonal.Final, SubsidioPersonal.CodOperario, 1 AS Cantidad" & _
" FROM SubsidioPersonal WHERE (((SubsidioPersonal.Inicio)<=#" & Fecha & "#) AND ((SubsidioPersonal.Final)>=#" & Fecha & "#)" & _
" AND ((SubsidioPersonal.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
SubsidiosFecha = rst("Cantidad")
rst.Close

Exit Function

Nulo:
SubsidiosFecha = 0


End Function

Function CompensacionFeriadosFecha(fechi As Date, Ope As Integer) As Long
'Dias de compensacion de feriados de operario fecha

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'DIferencia de dias feriados en un fecha
miSQL = "SELECT CompensacionFeriados.Inicio, CompensacionFeriados.Final, CompensacionFeriados.CodOperario, 1 AS Cantidad" & _
" FROM CompensacionFeriados WHERE (((CompensacionFeriados.Inicio)<=#" & fechi & "#)" & _
" AND ((CompensacionFeriados.Final)>=#" & fechi & "#) AND ((CompensacionFeriados.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
CompensacionFeriadosFecha = rst("Cantidad")
rst.Close

Exit Function

Nulo:
CompensacionFeriadosFecha = 0


End Function

Function VacacionesPeriodo(desde As Date, hasta As Date, Ope As Integer) As Double
'Dias de vaciaciones en periodo
Dim fech As Date

VacacionesPeriodo = 0

For fech = desde To hasta
VacacionesPeriodo = VacacionesPeriodo + VacacionesFecha(fech, Ope)
Next fech

Exit Function

Nulo:
VacacionesPeriodo = 0

End Function

Function SubsidiosPeriodo(desde As Date, hasta As Date, Ope As Integer) As Double
'Dias de subsidio en periodo
Dim fech As Date

SubsidiosPeriodo = 0

For fech = desde To hasta
SubsidiosPeriodo = SubsidiosPeriodo + SubsidiosFecha(fech, Ope)
Next fech

Exit Function

Nulo:
SubsidiosPeriodo = 0

End Function

Function HorasSemanaVacaciones(Fecha As Date, Ope As Integer) As Double
'Buscar Horas de semana promedio de vacaciones

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'DIferencia de dias feriados en un fecha
miSQL = "SELECT Vacaciones.Inicio, Vacaciones.Final, Vacaciones.CodOperario, Vacaciones.HorasPromedio FROM Vacaciones WHERE (((Vacaciones.Inicio)<=#" & Fecha & "#) AND ((Vacaciones.Final)>=#" & Fecha & "#) AND ((Vacaciones.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
HorasSemanaVacaciones = rst("HorasPromedio")
rst.Close

Exit Function

Nulo:
HorasSemanaVacaciones = 0


End Function

Function HorasSemanaVacacionesTotal(desde As Date, hasta As Date, Ope As Integer) As Double
'Horas de semana promedio de vacaciones
Dim fech As Date

For fech = desde To hasta
    If HorasSemanaVacaciones(fech, Ope) = 0 Then
        HorasSemanaVacacionesTotal = 0
    Else
        HorasSemanaVacacionesTotal = HorasSemanaVacaciones(fech, Ope)
        Exit Function
    End If

Next fech

Exit Function

Nulo:
HorasSemanaVacacionesTotal = 0

End Function
Function DiferenciaHoras(fech As Date, Ope As Integer) As Double
'sumatoria de hora inicial menos hora final de un dia completo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de horas realizadas
miSQL = "SELECT RegistroHorario.Fecha, Sum(RedondearHora([Salida],1)-RedondearHora([Ingreso],0)) AS Horas, RegistroHorario.CodOperario FROM RegistroHorario GROUP BY RegistroHorario.Fecha, RegistroHorario.CodOperario HAVING (((RegistroHorario.Fecha)=#" & fech & "#) AND ((RegistroHorario.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
DiferenciaHoras = rst("Horas")
rst.Close

Exit Function

Nulo:
DiferenciaHoras = 0

End Function
Function DiferenciaHorassinredondeo(fech As Date, Ope As Integer) As Double
'sumatoria de hora inicial menos hora final de un dia completo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

'SUma de horas realizadas
miSQL = "SELECT RegistroHorario.Fecha, Sum([Salida]-[Ingreso]) AS Horas, RegistroHorario.CodOperario FROM RegistroHorario GROUP BY RegistroHorario.Fecha, RegistroHorario.CodOperario HAVING (((RegistroHorario.Fecha)=#" & fech & "#) AND ((RegistroHorario.CodOperario)=" & Ope & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
DiferenciaHorassinredondeo = rst("Horas") * 24
rst.Close

Exit Function

Nulo:
DiferenciaHorassinredondeo = 0

End Function
Function sitrabajo(fech As Date, Ope As Integer, hor As Integer) As Boolean
'si trabajo o no en una hora especifica en una fecha especifica un operario especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.CodOperario, RegistroHorario.Fecha, Sum(IIf(RedondearHora([ingreso],0)<=" & hor & ", " & _
"IIf(IsNull([Salida])<>0,-1,IIf(RedondearHora([Salida],1)>=" & hor & ",-1,0)),0)) AS Cond " & _
"FROM RegistroHorario GROUP BY RegistroHorario.CodOperario, RegistroHorario.Fecha " & _
"HAVING (((RegistroHorario.CodOperario)=" & Ope & ") AND ((RegistroHorario.Fecha)=#" & fech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
sitrabajo = CBool(rst("Cond"))
'MsgBox rst("Ingreso").Count

'If RedondearHora(rst("Ingreso"), 0) <= hor Then
'    If IsNull(rst("Salida")) Then
'        sitrabajo = True
'    ElseIf RedondearHora(rst("Salida"), 1) >= hor Then
'        sitrabajo = True
'    Else
'        sitrabajo = False
'    End If
'Else
'    sitrabajo = False
'End If
    
rst.Close

Exit Function

Nulo:
sitrabajo = False

End Function

Function horastrabajadasdia(fech As Date) As Double
'horas trabajadas en total de todos los operarios

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.Fecha," & _
" Sum(RedondearHora([RegistroHorario]![Salida],1)-RedondearHora([RegistroHorario]![Ingreso],0)) AS Horas" & _
" FROM RegistroHorario" & _
" GROUP BY RegistroHorario.Fecha" & _
" HAVING (((RegistroHorario.Fecha)=#" & fech & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
horastrabajadasdia = rst("Horas")
    
rst.Close

Exit Function

Nulo:
horastrabajadasdia = 0

End Function

Function ratiohorasventasinicio(fech As Date) As Double
'ventas del dia 12 del mes entre horas trabajads del dia 1 del mes para que se alineen en grafica
Dim diaf As Date
diaf = DateSerial(Year(fech), Month(fech), 1)

ratiohorasventasinicio = AcumuladoDiaGuardado(diaf) / horastrabajadasdia(diaf)

Exit Function

Nulo:
ratiohorasventasinicio = 0

End Function

Function ultimodiamarcadopersonal(codi As Long) As Date
'ultimo dia que amrco un eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.CodOperario, RegistroHorario.Fecha" & _
" FROM RegistroHorario WHERE (((RegistroHorario.CodOperario) = " & codi & "))" & _
" ORDER BY RegistroHorario.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimodiamarcadopersonal = rst("Fecha")
    
rst.Close

Exit Function

Nulo:
ultimodiamarcadopersonal = 0

End Function
Function ultimodiamarcadopersonalmixed(codi As Long) As Date
'ultimo dia que amrco un eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorarioMixed.CodOperario, RegistroHorarioMixed.Fecha" & _
" FROM RegistroHorarioMixed WHERE (((RegistroHorarioMixed.CodOperario) = " & codi & "))" & _
" ORDER BY RegistroHorarioMixed.Fecha DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimodiamarcadopersonalmixed = rst("Fecha")
    
rst.Close

Exit Function

Nulo:
ultimodiamarcadopersonalmixed = 0

End Function

Function primerdiamarcadopersonal(codi As Long) As Date
'primerdia que amrco un eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.CodOperario, RegistroHorario.Fecha" & _
" FROM RegistroHorario WHERE (((RegistroHorario.CodOperario) = " & codi & "))" & _
" ORDER BY RegistroHorario.Fecha ASC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
primerdiamarcadopersonal = rst("Fecha")
    
rst.Close

Exit Function

Nulo:
primerdiamarcadopersonal = 0

End Function

Function primerdiamarcadopersonalmixed(codi As Long) As Date
'primerdia que amrco un eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorarioMixed.CodOperario, RegistroHorarioMixed.Fecha" & _
" FROM RegistroHorarioMixed WHERE (((RegistroHorarioMixed.CodOperario) = " & codi & "))" & _
" ORDER BY RegistroHorarioMixed.Fecha ASC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
primerdiamarcadopersonalmixed = rst("Fecha")
    
rst.Close

Exit Function

Nulo:
primerdiamarcadopersonalmixed = 0

End Function

Function ultimosalariofijopersonal(codi As Long) As Date
'ultimo salario fijo en funcion a horas asignadas y tarifa de eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TarifaPersonal.CodOperario, TarifaPersonal.FechaInicial, [Tarifa]/(1/(30/7)/48)/48*([HL]+[HMA]+[HMI]+[HJ]+[HV]+[HS]+[HD]) AS Salario" & _
" FROM TarifaPersonal WHERE (((TarifaPersonal.CodOperario) = " & codi & "))" & _
" ORDER BY TarifaPersonal.FechaInicial DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimosalariofijopersonal = rst("Salario")
rst.Close

Exit Function

Nulo:
ultimosalariofijopersonal = 0

End Function

Function ultimohorasasignadassemanapersonal(codi As Long) As Date
'ultimo horas asignadas a al semana de  de eprsonal

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT TarifaPersonal.CodOperario, TarifaPersonal.FechaInicial, [HL]+[HMA]+[HMI]+[HJ]+[HV]+[HS]+[HD] AS horario" & _
" FROM TarifaPersonal WHERE (((TarifaPersonal.CodOperario) = " & codi & "))" & _
" ORDER BY TarifaPersonal.FechaInicial DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimohorasasignadassemanapersonal = rst("horario")
rst.Close

Exit Function

Nulo:
ultimohorasasignadassemanapersonal = 0

End Function

Function AbrirPagoPersonal()
DoCmd.OpenForm "LogueoAutorizacion"
[Forms]![LogueoAutorizacion]![CodigoBusqueda] = cargooperariooperativo(8, 18, Date)

[Forms]![LogueoAutorizacion]![origenlogueo] = "Pago Personal"

End Function

Function existeturnodiaoperario(codi As Long, fech As Date) As Integer
'corrobora si ha hecho turno un dia especifico

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.Fecha, RegistroHorario.CodOperario, 1 AS cont" & _
" FROM RegistroHorario " & _
" WHERE (((RegistroHorario.Fecha)=#" & fech & "#) AND ((RegistroHorario.CodOperario)=" & codi & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
existeturnodiaoperario = rst("cont")
rst.Close

Exit Function

Nulo:
existeturnodiaoperario = 0

End Function

Sub rellenarturnosvaciosoperarios(Ope As Long, ini As Date, fini As Date)
'rellenar con turnos vacios de operarios en un periodo de tiempo

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contador As Integer

miSQL = "SELECT FechaSistema.Dates FROM FechaSistema" & _
" WHERE (((FechaSistema.Dates) Between #" & ini & "# And #" & fini & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
contador = 0

Do While Not rst.EOF

    If existeturnodiaoperario(Ope, rst("Dates")) = 0 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [RegistroHorario](Ingreso, Salida, Fecha, CodOperario) values" & _
        " (#12:00 AM#, #12:00 AM#, #" & rst("Dates") & "#, " & Ope & ")"
        DoCmd.SetWarnings True
        contador = contador + 1
    End If
    rst.MoveNext
Loop
rst.Close
'MsgBox "Se agregaron " & contador & " turnos vacios"

Exit Sub

Nulo:
'MsgBox "Error al crear turnos vacios"

End Sub

Function horasacumuladasvacacionesoperario(Ope As Integer, ini As Date, fini As Date)
'horas totales asignadas de dias de vacaciones de un tervalo de fechas

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT FechaSistema.Dates FROM FechaSistema" & _
" WHERE (((FechaSistema.Dates) Between #" & ini & "# And #" & fini & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
horasacumuladasvacacionesoperario = 0

Do While Not rst.EOF
    If VacacionesFecha(rst("Dates"), Ope) = 1 Then
        horasacumuladasvacacionesoperario = horasacumuladasvacacionesoperario + HorasTeoricas(Ope, rst("Dates"))
    End If
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
horasacumuladasvacacionesoperario = 0

End Function

Function ultimocodigooperario() As Long
'ultimo codigo creado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Operarios.CodOperario FROM Operarios ORDER BY Operarios.CodOperario DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimocodigooperario = rst("CodOperario")
    
rst.Close

Exit Function

Nulo:
ultimocodigooperario = 0

End Function

Function ultimocodigoasignacionoperario() As Long
'ultimo codigo creado

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT AsignacionNivelesCargos.CodAsignacionNivelesCargos FROM AsignacionNivelesCargos ORDER BY AsignacionNivelesCargos.CodAsignacionNivelesCargos DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ultimocodigoasignacionoperario = rst("CodAsignacionNivelesCargos")
    
rst.Close

Exit Function

Nulo:
ultimocodigoasignacionoperario = 0

End Function

Function horasmarcadastextototal(opec As Integer, fec As Date)
'horas totales asignadas de dias de vacaciones de un tervalo de fechas

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contax As Integer

miSQL = "SELECT RegistroHorario.CodOperario, RegistroHorario.Fecha, RegistroHorario.Ingreso, RegistroHorario.Salida" & _
" FROM RegistroHorario" & _
" WHERE (((RegistroHorario.CodOperario)=" & opec & ") AND ((RegistroHorario.Fecha)=#" & fec & "#))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
horasmarcadastextototal = ""
contax = 0

Do While Not rst.EOF
    If contax <> 0 Then
        horasmarcadastextototal = horasmarcadastextototal & " / "
    End If
    horasmarcadastextototal = horasmarcadastextototal & Format(rst("Ingreso"), "hh:nn AM/PM") & " a " & Format(rst("Salida"), "hh:nn AM/PM")
    contax = contax + 1
    rst.MoveNext
Loop
rst.Close

Exit Function

Nulo:
horasmarcadastextototal = "Libre"

End Function

Function horasplanificadastextototal(opec As Integer, fec As Date, sucu As Integer)
'horas totales asignadas de dias de vacaciones de un tervalo de fechas

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim contax As Integer
Dim diax As Integer
Dim diaelegidoentrada As String
Dim diaelegidosalida As String
Dim diaelegidotipo As String

miSQL = "SELECT HorariosSucursalesAprobados.CodOperario, HorariosSucursalesAprobados.semana, HorariosSucursalesAprobados.Sucursal," & _
" HorariosSucursalesAprobados.CodTipoLunes, HorariosSucursalesAprobados.TE1Lunes, HorariosSucursalesAprobados.TS1Lunes," & _
" HorariosSucursalesAprobados.CodTipoMartes, HorariosSucursalesAprobados.TE1Martes, HorariosSucursalesAprobados.TS1Martes," & _
" HorariosSucursalesAprobados.CodTipoMiercoles, HorariosSucursalesAprobados.TE1Miercoles, HorariosSucursalesAprobados.TS1Miercoles," & _
" HorariosSucursalesAprobados.CodTipoJueves, HorariosSucursalesAprobados.TE1Jueves, HorariosSucursalesAprobados.TS1Jueves," & _
" HorariosSucursalesAprobados.CodTipoViernes, HorariosSucursalesAprobados.TE1Viernes, HorariosSucursalesAprobados.TS1Viernes," & _
" HorariosSucursalesAprobados.CodTipoSabado, HorariosSucursalesAprobados.TE1Sabado, HorariosSucursalesAprobados.TS1Sabado," & _
" HorariosSucursalesAprobados.CodTipoDomingo, HorariosSucursalesAprobados.TE1Domingo, HorariosSucursalesAprobados.TS1Domingo" & _
" FROM HorariosSucursalesAprobados" & _
" WHERE (((HorariosSucursalesAprobados.CodOperario)=" & opec & ")" & _
" AND ((HorariosSucursalesAprobados.semana)=" & numerosemana(fec) & ")" & _
" AND ((HorariosSucursalesAprobados.Sucursal)=" & sucu & "))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveFirst
horasplanificadastextototal = ""
contax = 0
diax = Weekday(fec, vbMonday)

Select Case diax
    Case 1
        diaelegidoentrada = "TE1Lunes"
        diaelegidosalida = "TS1Lunes"
        diaelegidotipo = "CodTipoLunes"
    Case 2
        diaelegidoentrada = "TE1Martes"
        diaelegidosalida = "TS1Martes"
        diaelegidotipo = "CodTipoMartes"
    Case 3
        diaelegidoentrada = "TE1Miercoles"
        diaelegidosalida = "TS1Miercoles"
        diaelegidotipo = "CodTipoMiercoles"
    Case 4
        diaelegidoentrada = "TE1Jueves"
        diaelegidosalida = "TS1Jueves"
        diaelegidotipo = "CodTipoJueves"
    Case 5
        diaelegidoentrada = "TE1Viernes"
        diaelegidosalida = "TS1Viernes"
        diaelegidotipo = "CodTipoViernes"
    Case 6
        diaelegidoentrada = "TE1Sabado"
        diaelegidosalida = "TS1Sabado"
        diaelegidotipo = "CodTipoSabado"
    Case 7
        diaelegidoentrada = "TE1Domingo"
        diaelegidosalida = "TS1Domingo"
        diaelegidotipo = "CodTipoDomingo"
    Case Else
        diaelegidoentrada = ""
        diaelegidosalida = ""
        diaelegidotipo = 0
End Select

Do While Not rst.EOF
    If contax <> 0 Then
        horasplanificadastextototal = horasplanificadastextototal & " / "
    End If
    horasplanificadastextototal = horasplanificadastextototal & Format(rst(diaelegidoentrada), "hh:nn AM/PM") & " a " & Format(rst(diaelegidosalida), "hh:nn AM/PM")
    contax = contax + 1
    rst.MoveNext
Loop

If horasplanificadastextototal = " a " Then
    horasplanificadastextototal = DLookup("[Nombre]", "[TiposTurnoHorariosSucursales]", "[CodTiposTurnoHorariosSucursales]=" & rst(diaelegidotipo))
End If
rst.Close

Exit Function

Nulo:
horasplanificadastextototal = "Libre"
rst.Close
End Function

Function operariosegunclave(clave As String) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT Operarios.clave, Operarios.CodOperario FROM Operarios WHERE (((Operarios.clave)='" & clave & "'))"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
operariosegunclave = rst("CodOperario")
    
rst.Close

Exit Function

Nulo:
operariosegunclave = 0

End Function
Function horaingresopendienteoperariodia(opex As Integer, diax As Date) As Date

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.CodHorario, RegistroHorario.Salida, RegistroHorario.CodOperario, RegistroHorario.Fecha, RegistroHorario.Ingreso" & _
" FROM RegistroHorario" & _
" WHERE (((RegistroHorario.Salida) Is Null) And ((RegistroHorario.CodOperario) = " & opex & ") And ((RegistroHorario.Fecha) = #" & diax & "#))" & _
" ORDER BY RegistroHorario.CodHorario DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
horaingresopendienteoperariodia = rst("Ingreso")
    
rst.Close

Exit Function

Nulo:
horaingresopendienteoperariodia = 0

End Function
Function registroencursooperariodia(opex As Integer, diax As Date) As Long

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT RegistroHorario.CodHorario, RegistroHorario.Salida, RegistroHorario.CodOperario, RegistroHorario.Fecha" & _
" FROM RegistroHorario" & _
" WHERE (((RegistroHorario.Salida) Is Null) And ((RegistroHorario.CodOperario) = " & opex & ") And ((RegistroHorario.Fecha) = #" & diax & "#))" & _
" ORDER BY RegistroHorario.CodHorario DESC"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
registroencursooperariodia = rst("CodHorario")
rst.Close

Exit Function

Nulo:
registroencursooperariodia = 0

End Function


Public Function ObtenerCodOperario(ByVal claveOperario As String) As Long
    ' Función que devuelve el CodOperario según la clave
    ' Si no existe, devuelve 0
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim CodOperario As Long
    
    ' Inicializar con 0 (valor por defecto si no encuentra)
    CodOperario = 0
    
    ' Crear la consulta SQL
    sql = "SELECT Operarios.clave, Operarios.CodOperario " & _
          "FROM marcacioneshoy INNER JOIN Operarios " & _
          "ON marcacioneshoy.CodOperario = Operarios.CodOperario " & _
          "WHERE Operarios.clave = '" & Replace(claveOperario, "'", "''") & "'"
    
    ' Otra versión alternativa si quieres solo buscar en Operarios
    ' sql = "SELECT CodOperario FROM Operarios WHERE clave = '" & Replace(claveOperario, "'", "''") & "'"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Verificar si hay registros
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        If Not IsNull(rs!CodOperario) Then
            CodOperario = rs!CodOperario
        End If
    End If
    
    ' Cerrar objetos
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ObtenerCodOperario = CodOperario
    Exit Function
    
ErrorHandler:
    ' En caso de error, devolver 0
    ObtenerCodOperario = 0
    
    ' Cerrar objetos si están abiertos
    If Not rs Is Nothing Then
        'If Not (rs.State And adStateClosed) Then rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function

Public Function usuarioconpermisosclave(ByVal claveOperario As String) As Long
    ' Versión optimizada de la consulta
    ' Devuelve el CodOperario si es líder activo, 0 si no
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Valor por defecto
    usuarioconpermisosclave = 0
    
    ' Consulta optimizada - más legible
    sql = "SELECT o.CodOperario " & _
          "FROM ((NivelesCargos nc " & _
          "INNER JOIN AsignacionNivelesCargos anc ON nc.CodNivelesCargos = anc.CodNivelesCargos) " & _
          "INNER JOIN Operarios o ON anc.CodOperario = o.CodOperario) " & _
          "WHERE nc.PermisosLider = True " & _
          "AND o.clave = '" & Replace(claveOperario, "'", "''") & "' " & _
          "AND anc.Fecha <= Date() " & _
          "AND (anc.Fin IS NULL OR anc.Fin >= Date())"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        If Not IsNull(rs!CodOperario) Then
            usuarioconpermisosclave = rs!CodOperario
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    usuarioconpermisosclave = 0
    If Not rs Is Nothing Then
        'If rs.State <> adStateClosed Then rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function

