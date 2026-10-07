' ==========================================================
' Modulo  : Actualizar Sistema
' Tipo    : 1  |  Lineas: 1153
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:08
' ==========================================================

Option Compare Database




Function MinimizarIntroduccion() As String
On Error GoTo Nulo
DoCmd.Close
Exit Function

Nulo:
End Function

Function NombreSistema() As String
Dim wshNetwork As Object 'New wshNetwork
Set wshNetwork = CreateObject("WScript.Network")
NombreSistema = wshNetwork.UserName
Set wshNetwork = Nothing
End Function

Function NombrePC() As String
Dim wshNetwork As Object 'New wshNetwork
Set wshNetwork = CreateObject("WScript.Network")
NombrePC = wshNetwork.ComputerName
Set wshNetwork = Nothing
End Function

Function actualizartablas(moda As Integer, loc As Integer)
'moda
'1: Aplicacion Gestion Solo Lectura MAIN VISTA + CENTRAL VISTA + SUCURSAL VISTA
'2: Aplicacion Gestion Edicion      MAIN EDICION + CENTRAL VISTA + SUCIRSAL VISTA
'3: Aplicacion Local                MAIN VISTA + CENTRAL VISTA + SUCURSAL EDICION
'Call eliminarvinculos
Call eliminartablas
Select Case moda
    Case 1
        Call importartablasmain
        Call importartablascentral(loc)
        Call importartablaslocal(loc)
        If APIDisponible() Then
            Call importartablasweb
        Else
            Call importartablaespecifica("RRHH", "Operarios", "Operarios", 3)
            Call importartablaespecifica("RRHH", "AsignacionNivelesCargos", "AsignacionNivelesCargos", 3)
            Call importartablaespecifica("RRHH", "NivelesCargos", "NivelesCargos", 3)
        End If
    Case 2
        Call vinculartablasmain
        Call importartablascentral(loc)
        Call importartablaslocal(loc)
        If APIDisponible() Then
            Call importartablasweb
        Else
            Call importartablaespecifica("RRHH", "Operarios", "Operarios", 3)
            Call importartablaespecifica("RRHH", "AsignacionNivelesCargos", "AsignacionNivelesCargos", 3)
            Call importartablaespecifica("RRHH", "NivelesCargos", "NivelesCargos", 3)
        End If
    Case 3
        
        Call eliminararchivosllavelectura
        Call desvinculartablas(loc) ' copiar archivo de sucursal a una copia eliminar la ofriginal y copiar la copia como nuevo
        Call importartablasmain
        Call importartablascentral(loc)
        Call vinculartablaslocal(loc)
        If APIDisponible() Then
            Call importartablasweb
        Else
            Call importartablaespecifica("RRHH", "Operarios", "Operarios", 3)
            Call importartablaespecifica("RRHH", "AsignacionNivelesCargos", "AsignacionNivelesCargos", 3)
            Call importartablaespecifica("RRHH", "NivelesCargos", "NivelesCargos", 3)
        End If
End Select

End Function

Sub actualizartablasmodulo(modu As String)
Call eliminararchivosllavelectura
Call desvinculartablasmodulo(modu) ' copiar archivo de mpdulo a una copia eliminar la ofriginal y copiar la copia como nuevo

Call eliminartablas

Call vinculartablasmodulo(modu)
Call importartablasmain
Call importartablaslocal(0)
Call importartablascentral(0)
Call importartablasweb

MsgBox "Datos de modulo vinculados correctamente"
End Sub
Sub actualizartablasmixedglobal()

Dim nombrearchi As String
nombrearchi = CurrentProject.Name

If nombrearchi Like "Modulo*" Or nombrearchi Like "Pitaya_System*" Then
    MsgBox "es raiz"
    Call eliminartablasMixed(-1)
    Call importartablasMainMixed
    Call importartablasLocalMixed(-1)
    Call actualizardatosmodulosxmixedglobal
        
Else ' de sucursal especifica incluyendo 0
    Dim indi As String
    indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
    'Call eliminararchivosllavelectura
    Call eliminartablasMixed(CInt(indi))  ' el codigo de local no sirve para nada, se borra todo menos vinculadas que solo aplica a sucursales que se vncule en el mixed
    Call importartablasMainMixed
    Call importartablasLocalMixed(CInt(indi))
    If indi <> 0 Then ' ya sean modulos, sistema raiz o sistema de sucursales, en estos casos no tenemos pitaya 0 vinculado en la db mixed asi que chancamos con datos de modulos infdeentdinetes
        Call actualizardatosmodulosxmixedglobal
    End If
    
End If

End Sub
Sub actualizardatosmodulosxmixedglobal()

On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim tota As Integer

Dim tablac As String
Dim moduloc As String
Dim tipoc As String

miSQL = "SELECT TablasModulos.Tabla, TablasModulos.Modulo, TablasModulos.Tipo FROM TablasModulos"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
tota = rst.RecordCount
rst.MoveFirst

For I = 1 To tota
    tablac = rst("Tabla")
    moduloc = rst("Modulo")
    tipoc = rst("Tipo")
    
    Select Case tipoc
        Case "Central"
            Call importartablaespecificaAMixed(moduloc, tablac, tablac)
        Case "Pitaya0"
            Call importartablaespecificaAMixed(moduloc, tablac, tablac & "0")
        Case Else
    End Select

    rst.MoveNext
Next I

rst.Close

Exit Sub

Nulo:
rst.Close
MsgBox "Los datos no se descargaron correctamente"
End Sub
Sub actualizardatosclubmixedglobal()
    On Error GoTo Nulo
    Dim ax As Integer
    Dim nombrearchi As String
    Dim condi As Integer
    Dim indi As String
    Dim intentos As Integer
    Dim maxIntentos As Integer
    Dim sucursalesex As Integer
    
    sucursalesex = cantidadsucursalesexistentes()
    nombrearchi = CurrentProject.Name
    condi = 1
    maxIntentos = 5 ' Número máximo de intentos por sucursal
    
    If nombrearchi Like "Modulo*" Or nombrearchi = "Pitaya_System.accdb" Then
        indi = 0
        condi = 1
    Else
        indi = Split(Split(CurrentProject.Name, "_")(0), "Pitaya")(1)
        condi = 2
    End If
    
    ax = 1
    While ax <= sucursalesex
        intentos = 0
        Dim exito As Boolean
        exito = False
        
        ' Intentar hasta maxIntentos veces
        While intentos < maxIntentos And Not exito
            On Error Resume Next
            If condi = 2 And indi = ax Then
                ' Hay excepción cuando es pitaya de sucursal y coincide
                exito = True ' Consideramos éxito ya que no se procesa
            Else
                Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "ClientesClub", "ClientesClub" & ax)
                If Not APIDisponible() Then
                    'cuando no hay conexion tambien descargar datos de subpedido y ntoa de pedido
                    Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "NotaDePedido", "NotaDePedido" & ax)
                    Call importartablaespecificaAMixed("Pitaya" & ax & "_DB", "SubPedido", "SubPedido" & ax)
                End If
                If Err.Number = 0 Then
                    exito = True
                Else
                    intentos = intentos + 1
                    Err.Clear
                    If intentos < maxIntentos Then
                        ' Esperar un momento antes de reintentar (opcional)
                        Sleep 2000
                    End If
                End If
            End If
            On Error GoTo Nulo
        Wend
        
        If exito Then
            ax = ax + 1 ' Solo avanzar si fue exitoso
        Else
            ' Si falló después de todos los intentos, registrar error y avanzar
            'Call EnviarMensajeTelegram("Error al descargar Club Pitaya de sucursal " & ax & " después de " & maxIntentos & " intentos", grupotoperaciones())
            ax = ax + 1 ' Avanzar para no quedarse en bucle infinito
        End If
    Wend
    
    'Call EnviarMensajeTelegram("Datos de Club Pitaya actualizados correctamente", grupotgerencia())
    Exit Sub

Nulo:
    MsgBox "No se ha podido descargar las membresias de la sucursal" & ax & ", Actualiza nuevamente las membresias de otras sucursales"
    'Call EnviarMensajeTelegram("Problemas al descargar Club Pitaya de sucursal " & ax, grupotoperaciones())
End Sub

Function actualizarsistema()
If MsgBox("Desea actualizar el sistema?", vbYesNo, "ACTUALIZACION DE SISTEMA") = vbYes Then
    Call eliminarconsultas
    Call eliminarformularios
    Call eliminarinformes
    'Call eliminarmacros
    Call eliminarmodulos
    Call importarobjetos
    MsgBox "Sistema Actualizado Correctamente"
End If
End Function
Sub eliminararchivosllavelectura()
Dim temparch As String
temparch = Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\*.laccdb")
Do Until temparch = ""
    Kill ("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & temparch)
    temparch = Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\*.laccdb")
Loop
End Sub
Sub eliminartablas()
Dim Tabla As TableDef

For Each Tabla In CurrentDb.TableDefs
    If Left(Tabla.Name, 4) <> "USys" And Left(Tabla.Name, 4) <> "MSys" And Left(Tabla.Name, 4) <> "~TMP" Then
        DoCmd.DeleteObject acTable, Tabla.Name
    End If
Next
End Sub
Sub eliminarvinculos()
Dim db As DAO.Database
Dim rel As DAO.Relation

Set db = CurrentDb
For Each rel In db.Relations
    If rel.Table = sTable Or rel.ForeignTable = sTable Then
        MsgBox rel.Name
        db.Relations.Delete (rel.Name)
        
    End If
Next rel
End Sub

Sub eliminarmacros()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

Do While Not dbs.AllMacros.Count = 0
    For Each obj In dbs.AllMacros
        On Error Resume Next
        DoCmd.DeleteObject acMacro, obj.Name
    Next
Loop
End Sub

Sub eliminarformularios()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

Do While Not dbs.AllForms.Count = 0
    For Each obj In dbs.AllForms
        On Error Resume Next
        DoCmd.DeleteObject acForm, obj.Name
    Next
Loop
End Sub

Sub eliminarmodulos()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

Do While Not dbs.AllModules.Count = 1 'modulo que quedan despues de borraar todo
    For Each obj In dbs.AllModules
        On Error Resume Next
        If Not obj.Name = "Actualizar Sistema" Then 'cnatidad de modulos que debn quedar
        DoCmd.DeleteObject acModule, obj.Name
        End If
    Next
Loop
End Sub

Sub eliminarinformes()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

Do While Not dbs.AllReports.Count = 0
    For Each obj In dbs.AllReports
        On Error Resume Next
        DoCmd.DeleteObject acReport, obj.Name
    Next
Loop
End Sub

Sub eliminarconsultas()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentData

Do While Not dbs.AllQueries.Count = 0
    For Each obj In dbs.AllQueries
        On Error Resume Next
        DoCmd.DeleteObject acQuery, obj.Name
    Next
Loop
End Sub
Sub importarobjetos()
Dim appAccess As Access.Application
Dim fileNameStr As String
Dim frmSearch As Variant
Dim repSearch As Variant
Dim queSearch As Variant
Dim macSearch As Variant
Dim modSearch As Variant

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya_System.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Pitaya_System.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Set appAccess = New Access.Application
MsgBox "Mantener Presionado SHIFT"
appAccess.OpenCurrentDatabase h
MsgBox "Soltar SHIFT"
 
For Each frmSearch In appAccess.CurrentProject.AllForms
    DoCmd.TransferDatabase acImport, "Microsoft Access", h, acForm, frmSearch.Name, frmSearch.Name
Next frmSearch

For Each repSearch In appAccess.CurrentProject.AllReports
    DoCmd.TransferDatabase acImport, "Microsoft Access", h, acReport, repSearch.Name, repSearch.Name
Next repSearch

For Each queSearch In appAccess.CurrentData.AllQueries
    DoCmd.TransferDatabase acImport, "Microsoft Access", h, acQuery, queSearch.Name, queSearch.Name
Next queSearch

'For Each macSearch In appAccess.CurrentData.AllMacros
'   DoCmd.TransferDatabase acImport, "Microsoft Access", H, acMacro, macSearch.Name, macSearch.Name
'Next macSearch

For Each modSearch In appAccess.CurrentProject.AllModules
    If modSearch.Name <> "Actualizar Sistema" Then
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acModule, modSearch.Name, modSearch.Name
    End If
Next modSearch
 
appAccess.CloseCurrentDatabase
Set appAccess = Nothing
Set fs = Nothing
Kill (h)
End Sub

Sub eliminartablasMixed(loci As Integer)
' borra todo menos las tablas vicnuladas .connect >0
Dim db As Database
Dim td As TableDef
Dim M As String
M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"
Set db = DBEngine.Workspaces(0).OpenDatabase(M, True)
For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "USys" And Left(td.Name, 4) <> "~TMP" Then
        
        If Len(td.Connect) > 0 Then ' cuando se va a borrar las tablas mixed  no borra las tablas vinculadas, solo sucursales tendran vinculado el mixed
            'No hace nada
            
        Else
            db.Execute "DROP TABLE [" & td.Name & "]", dbFailOnError
            'DoCmd.DeleteObject acTable, td.Name
            'db.TableDefs.Delete td.Name
        End If
        
    End If
Next
db.Close
End Sub

Sub EliminarTablaDBExternaCarpetaUsuario(dbexterna As String, tabala As String)
Dim db As Database
Dim td As TableDef
Dim M As String
'M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"
M = "C:\Users\" & NombreSistema() & dbexterna
Set db = DBEngine.Workspaces(0).OpenDatabase(M, True)
For Each td In db.TableDefs
    If td.Name = tabala Then
        db.Execute "DROP TABLE [" & td.Name & "]", dbFailOnError
        'DoCmd.DeleteObject acTable, td.Name
        'db.TableDefs.Delete td.Name
    End If
Next
db.Close
End Sub

Sub eliminartablasmain()
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Main_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs 'recorrer nombres de tablas main copia
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        If existeTabla(td.Name) = 1 Then 'existe tabla
           DoCmd.DeleteObject acTable, td.Name
        End If
    End If
Next

db.Close
Set fs = Nothing
Kill (h)
End Sub


Sub eliminartablascentral(locaz As Integer)
'tipo: 1=completo, 2=si es modulo no elimina la de su modulo, 3 = sucursales
'
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim tota As Integer

Dim tablac As String
Dim moduloc As String
Dim esmodulo As String
Dim tipocaso As Integer
Dim tipotabla As String

If CurrentProject.Name Like "Modulo*" Then
    esmodulo = Split(Split(CurrentProject.Name, "Modulo")(1), ".accdb")(0)
    tipocaso = 2
ElseIf CurrentProject.Name = "Pitaya_System.accdb" Then 'pitaya raiz
    tipocaso = 1
ElseIf CurrentProject.Name Like "Pitaya*" Then 'sucursal
    tipocaso = 3
Else
    MsgBox "Archivo de sistema no permitido para mover datos"
    Exit Sub
End If


miSQL = "SELECT TablasModulos.Tabla, TablasModulos.Modulo, TablasModulos.Tipo FROM TablasModulos"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
tota = rst.RecordCount
rst.MoveFirst

For I = 1 To tota
    tablac = rst("Tabla")
    moduloc = rst("Modulo")
    tipotabla = rst("Tipo")
    
    Select Case tipocaso
        Case 1 'pitaya raiz solo elimina las tblas cntral , queda las de pitaya0 porque esas vans egun la sucursal abierta
            If locaz = 0 Then
                If existeTabla(tablac) = 1 Then 'Eliminar todas las de tipo central y pitaya0
                    DoCmd.DeleteObject acTable, tablac
                End If
            Else
                If tipotabla <> "Pitaya0" Then
                    If existeTabla(tablac) = 1 Then 'Eliminar todas las de tipo central solamente
                        DoCmd.DeleteObject acTable, tablac
                    End If
                End If
            End If
        Case 2 'modulos
            If moduloc <> esmodulo And existeTabla(tablac) = 1 Then ' solo borra si no es del modulo en cuestion incluyendo pitaya0 y central
                DoCmd.DeleteObject acTable, tablac
            End If
        Case 3
            If tipotabla <> "Pitaya0" Then
                If existeTabla(tablac) = 1 Then 'Eliminar todas las de tipo central, deja las de sucursal porque ya esta vinuclada de supropia sucursal
                    DoCmd.DeleteObject acTable, tablac
                End If
            End If
        Case Else
            MsgBox "Tipo asignado sin funcion especifica"
    End Select
    rst.MoveNext
Next I

rst.Close
Exit Sub

Nulo:
rst.Close
MsgBox "Los datos no se descargaron correctamente"

End Sub


Sub desvinculartablas(loc As Integer) ' copiar archivo de sucursal a una copia eliminar la ofriginal y copiar la copia como nuevo
Dim D, h As String
Dim fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & loc & "_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & loc & "local_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True
Set fs = Nothing
Kill (D)

Sleep 1000

Dim j, K As String
Dim fs2 As Object
j = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & loc & "local_DB.accdb"
K = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & loc & "_DB.accdb"
Set fs2 = CreateObject("Scripting.FileSystemObject")
fs2.copyfile j, K, True
Set fs2 = Nothing
Kill (j)
End Sub

Sub desvinculartablasmodulo(modu As String) ' copiar archivo de modulo a una copia eliminar la ofriginal y copiar la copia como nuevo
Dim D, h As String
Dim fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & modu & ".accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & modu & "local_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True
Set fs = Nothing
Kill (D)

Sleep 1000

Dim j, K As String
Dim fs2 As Object
j = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & modu & "local_DB.accdb"
K = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & modu & ".accdb"
Set fs2 = CreateObject("Scripting.FileSystemObject")
fs2.copyfile j, K, True
Set fs2 = Nothing
Kill (j)
End Sub

Sub ImportarTablaDBExternaToDBExternaCarpetaUsuario(loc As Integer)
On Error GoTo Nulo
Dim D, h As String
Dim fs As Object
Dim cod As Integer
Dim M As String
Dim db As Database
Dim td As TableDef
M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"

For cod = 0 To cantidadsucursalesexistentes()
    D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & cod & "_DB.accdb"
    h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\CopiaMixed" & NombrePC() & NombreSistema() & "Pitaya" & cod & "_DB.accdb"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.copyfile D, h, True
    
    Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
            If cod = loc And cod <> 0 Then ' en lectura siempre, solo cuando es local 1 para arriba copia sus tablas en vinculados
                Dim ac As Access.Application
                Set ac = New Access.Application
                ac.OpenCurrentDatabase (M)
                ac.DoCmd.TransferDatabase acLink, "Microsoft Access", D, acTable, td.Name, td.Name & cod
                ac.CloseCurrentDatabase
                ac.Quit acQuitSaveNone
                Set ac = Nothing
            Else
                db.Execute "SELECT * INTO [" & td.Name & cod & "] IN '" & M & "' FROM [" & td.Name & "]"
            End If
        End If
    Next
    db.Close
    Set fs = Nothing
    Kill (h)
Next cod
Exit Sub

Nulo:
MsgBox td.Name
MsgBox "Error al importar, interntar neuvamente"
End Sub

Sub importartablasLocalMixed(loc As Integer)
On Error GoTo Nulo
Dim D, h As String
Dim fs As Object
Dim cod As Integer
Dim M As String
Dim db As Database
Dim td As TableDef
M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"
   
For cod = 0 To cantidadsucursalesexistentes()

    D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & cod & "_DB.accdb"
    h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\CopiaMixed" & NombrePC() & NombreSistema() & "Pitaya" & cod & "_DB.accdb"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.copyfile D, h, True
    
    Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
            If cod = loc Then ' la base de datos de loc se copia como vinculado
                'ya no se importan las tablas de loc porqu quedo vinculado
                
                'Dim ac As Access.Application
                'Set ac = New Access.Application
                'ac.OpenCurrentDatabase (M)
                'ac.DoCmd.TransferDatabase acLink, "Microsoft Access", D, acTable, td.Name, td.Name & cod
                'ac.CloseCurrentDatabase
                'ac.Quit acQuitSaveNone
                'Set ac = Nothing
            Else
                db.Execute "SELECT * INTO [" & td.Name & cod & "] IN '" & M & "' FROM [" & td.Name & "]"
            End If
        End If
    Next
    
    On Error GoTo Nulo
    
    db.Close
    Set fs = Nothing
    Kill (h)
Next cod

'Call EnviarMensajeTelegram("Datos de sucursales actualizados correctamente", grupotgerencia())
MsgBox "Datos de sucursales actualizados correctamente"
Exit Sub

Nulo:
db.Close
Set fs = Nothing
'Call EnviarMensajeTelegram("Problemas al descargar datos de sucursales importarlocalmixed", grupotgerencia())
MsgBox "Problemas al descargar datos de sucursales"
End Sub



Sub vinculartablasLocalMixed(loc As Integer)
On Error GoTo Nulo
Dim D, h As String
Dim fs As Object

Dim M As String
Dim db As Database
Dim td As TableDef
M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"


D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & loc & "_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\CopiaMixed" & NombrePC() & NombreSistema() & "Pitaya" & loc & "_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        'importan las tablas de loc porqu quedo vinculado
        
        Dim ac As Access.Application
        Set ac = New Access.Application
        ac.OpenCurrentDatabase (M)
        ac.DoCmd.TransferDatabase acLink, "Microsoft Access", D, acTable, td.Name, td.Name & loc
        ac.CloseCurrentDatabase
        ac.Quit acQuitSaveNone
        Set ac = Nothing
    End If
Next
db.Close
Set fs = Nothing
Kill (h)



Exit Sub

Nulo:

'MsgBox td.Name
MsgBox "Error al importar, intentar nuevamente"
End Sub
Sub importartablasMainMixed()
Dim D, h As String
Dim fs As Object
Dim cod As Integer
Dim M As String
Dim db As Database
Dim td As TableDef
M = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\CopiaMixed" & NombrePC() & NombreSistema() & "Main_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        db.Execute "SELECT * INTO [" & td.Name & "] IN '" & M & "' FROM [" & td.Name & "]"
    End If
Next
db.Close
Set fs = Nothing
Kill (h)
End Sub

Sub importartablasmain()
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Main_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs ' recorro la base de datos main copia y escrbo todo
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, td.Name, td.Name
    End If
Next

db.Close
Set fs = Nothing
Kill (h)
End Sub


Sub importartablascentral(locaz As Integer)
'tipo: 1=completo vista, 2=si es modulo no sube tablas ya que hay vinculado
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String
Dim tota As Integer

Dim moduloc As String
Dim esmodulo As String
Dim tipocaso As Integer
Dim tipotabla As String

If CurrentProject.Name Like "Modulo*" Then
    esmodulo = Split(Split(CurrentProject.Name, "Modulo")(1), ".accdb")(0)
    tipocaso = 2
ElseIf CurrentProject.Name = "Pitaya_System.accdb" Then 'pitaya raiz
    tipocaso = 1
ElseIf CurrentProject.Name Like "Pitaya*" Then 'sucursal
    tipocaso = 3
Else
    MsgBox "Archivo de sistema no permitido para mover datos"
    Exit Sub
End If

miSQL = "SELECT TablasModulos.Modulo, TablasModulos.Tipo FROM TablasModulos GROUP BY TablasModulos.Modulo, TablasModulos.Tipo"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
rst.MoveLast
tota = rst.RecordCount
rst.MoveFirst

For I = 1 To tota
    moduloc = rst("Modulo")
    tipotabla = rst("Tipo")
    
    Select Case tipocaso
        Case 1 'Pitaya systema raiz
            If locaz = 0 Then ' si tinee cargado el sistema 0 en sistema raiz
                Call ImportarDBEspecifica(moduloc, 1)
            Else
                If tipotabla <> "Pitaya0" Then
                    Call ImportarDBEspecifica(moduloc, 1)
                End If
            End If
        Case 2 ' para modulos
            If moduloc = esmodulo Then
            ' No descarga ninguna tabla ya que estan ya vinculadas
            Else
                Call ImportarDBEspecifica(moduloc, 1)
            End If
        Case 3 'sicirsales sistema pitaya...
            If tipotabla <> "Pitaya0" Then 'importa todas las de tipo central si no existe
                Call ImportarDBEspecifica(moduloc, 1)
            End If
    End Select
    rst.MoveNext
Next I

rst.Close
Exit Sub

Nulo:
MsgBox tablac & moduloc
rst.Close
MsgBox "Los datos no se descargaron correctamente"

End Sub

Sub vinculartablasmain()
Dim archivo As String
Dim db As Database
Dim td As TableDef

archivo = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Main_DB.accdb"
Set db = DBEngine.Workspaces(0).OpenDatabase(archivo, True)

For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        DoCmd.TransferDatabase acLink, "Microsoft Access", archivo, acTable, td.Name, td.Name
    End If
Next
db.Close
End Sub

Sub importartablaslocal(cod As Integer)

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & cod & "_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Pitaya" & cod & "_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        If cod = 0 And DLookup("[Tipo]", "[TablasModulos]", "[Tabla]='" & td.Name & "'") = "Pitaya0" Then
            'No descarga nada porque en la formula de datos centrales ya decargo del tipo pitaya0, aplica solo pitaya central
        Else
            DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, td.Name, td.Name
        End If
    End If
Next
db.Close
Set fs = Nothing
Kill (h)

'If cod = 0 Then 'Cuando descarga de la central primero descagra datos de los modulos luego encima de la base datos 0
'    Call descargardatoscentralcopia("todos")
'End If

End Sub

Sub vinculartablaslocal(cod As Integer)
Dim archivo As String
Dim db As Database
Dim td As TableDef

archivo = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya" & cod & "_DB.accdb"
Set db = DBEngine.Workspaces(0).OpenDatabase(archivo, True)

For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        DoCmd.TransferDatabase acLink, "Microsoft Access", archivo, acTable, td.Name, td.Name
    End If
Next
db.Close
End Sub

Sub vinculartablasmodulo(modu As String)
Dim archivo As String
Dim db As Database
Dim td As TableDef

archivo = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & modu & ".accdb"
Set db = DBEngine.Workspaces(0).OpenDatabase(archivo, True)

For Each td In db.TableDefs
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        DoCmd.TransferDatabase acLink, "Microsoft Access", archivo, acTable, td.Name, td.Name
    End If
Next
db.Close
End Sub

Sub importarpreingresossistema0()
If codigoLocal() = 0 Then
    Exit Sub
End If

If DCount("*", "MSysObjects", "Type=1 AND Name='tPreIngresoPitaya'") > 0 Then
DoCmd.RunSQL "DROP TABLE [tPreIngresoPitaya]"
End If
If DCount("*", "MSysObjects", "Type=1 AND Name='tSubPreIngresosPitaya'") > 0 Then
DoCmd.RunSQL "DROP TABLE [tSubPreIngresosPitaya]"
End If

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\Pitaya0_DB.accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & "Pitaya0_DB.accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True
    
Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs
    If td.Name = "PreIngresoPitaya" Or td.Name = "SubPreIngresosPitaya" Then
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, td.Name, "t" & td.Name
    End If
Next

db.Close
Set fs = Nothing
Kill (h)

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM PreIngresoPitaya"
DoCmd.RunSQL "DELETE * FROM SubPreIngresosPitaya"
DoCmd.SetWarnings True

'Crear carpeta temporal
DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO PreIngresoPitaya SELECT * FROM [tPreIngresoPitaya]"
DoCmd.RunSQL "INSERT INTO SubPreIngresosPitaya SELECT * FROM [tSubPreIngresosPitaya]"
DoCmd.SetWarnings True
End Sub

Sub importartablaespecificaAMixed(dborigen As String, tablaorigen As String, tablafinal As String)
'ttablatemporal es comodin de cualquier tabla que se copie

Dim dborigenreal As String

'Verificar si existe archivo, si no existe usar el (1)
If Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigen & ".accdb") <> "" Then
    dborigenreal = dborigen
ElseIf Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigen & " (1)" & ".accdb") <> "" Then
    dborigenreal = dborigen & " (1)"
Else
    dborigenreal = dborigen & " (2)"
End If

'eliminar tabla temporal copia
If DCount("*", "MSysObjects", "Type=1 AND Name='ttablatemporal'") > 0 Then
    DoCmd.RunSQL "DROP TABLE [ttablatemporal]"
End If

'copia la base de datos a una copia
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigenreal & ".accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & dborigenreal & ".accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True
    
'buscar la tabla en la copia de base de datos, cuando la encuentra crea la tabla temporal con esa base de datos, modo copiado normal no vvinculado
DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, tablaorigen, "ttablatemporal"
Kill (h)

'Eliminar la tabla existente en la base de datos Mixed
On Error Resume Next
Dim db2 As Database
Dim M2 As String
M2 = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Pitaya_Mixed_System.accdb"
Set db2 = DBEngine.Workspaces(0).OpenDatabase(M2, True)
db2.Execute "DROP TABLE [" & tablafinal & "]", dbFailOnError
db2.Close
On Error GoTo 0

'traspasala tabla temporal a la tabla existente en mixed
DoCmd.SetWarnings False
DoCmd.TransferDatabase acExport, "Microsoft Access", M2, acTable, "ttablatemporal", tablafinal
DoCmd.SetWarnings True
End Sub

Sub importartablaespecifica(dborigen As String, tablaorigen As String, tablafinal As String, modo As Integer)
'ttablatemporal es comodin de cualquier tabla que se copie
'modo: 1 = copiar sobreescribir, si no existe lo crea, 2 = anexar, 3 = crea nueva tabla desde la creada temporal eliminando si ya existe

Dim dborigenreal As String

'Verificar si existe archivo, si no existe usar el (1)
If Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigen & ".accdb") <> "" Then
    dborigenreal = dborigen
Else
    dborigenreal = dborigen & " (1)"
End If
    
'eliminar tabla temporal copia
If DCount("*", "MSysObjects", "Type=1 AND Name='ttablatemporal'") > 0 Then
    DoCmd.RunSQL "DROP TABLE [ttablatemporal]"
End If

'copia la base de datos a una copia
Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigenreal & ".accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & dborigenreal & ".accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True


Select Case modo

    Case 1
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, tablaorigen, "ttablatemporal"
        
        'Vacia la tabla existente en la base de datos local
        DoCmd.SetWarnings False
       
        DoCmd.RunSQL "DELETE * FROM [" & tablafinal & "]"
        DoCmd.SetWarnings True
        
        'traspasa todos los datos de la tabla temporal a la tabla existente , si esta vinvulada igual la copia encima
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [" & tablafinal & "] SELECT * FROM [ttablatemporal]"
        DoCmd.SetWarnings True
    
    Case 2
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, tablaorigen, "ttablatemporal"
        
        'traspasa todos los datos de la tabla temporal a la tabla existente , si esta vinvulada igual la copia encima
        DoCmd.SetWarnings False
        DoCmd.RunSQL "INSERT INTO [" & tablafinal & "] SELECT * FROM [ttablatemporal]"
        DoCmd.SetWarnings True

    Case 3
        'eliminar tabla si ya existe
        If DCount("*", "MSysObjects", "Type=1 AND Name='" & tablafinal & "'") > 0 Then
            DoCmd.RunSQL "DROP TABLE [" & tablafinal & "]"
        End If
        
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, tablaorigen, tablafinal

End Select

Kill (h)

End Sub

Sub ImportarDBEspecifica(dborigen As String, modo As Integer)
Dim dborigenreal As String

'Verificar si existe archivo, si no existe usar el (1)
If Dir("C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigen & ".accdb") <> "" Then
    dborigenreal = dborigen
Else
    dborigenreal = dborigen & " (1)"
End If

Dim D, h As String, fs As Object
D = "C:\Users\" & NombreSistema() & "\Google Drive BP\Base de Datos Pitaya\" & dborigenreal & ".accdb"
h = "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Copia" & NombrePC() & NombreSistema() & dborigenreal & ".accdb"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.copyfile D, h, True

Dim db As Database
Dim td As TableDef
Set db = DBEngine.Workspaces(0).OpenDatabase(h, True)
For Each td In db.TableDefs ' recorro la base de datos del modulo copia y escrbo todo
    If Left(td.Name, 4) <> "MSys" And Left(td.Name, 4) <> "~TMP" And Left(td.Name, 4) <> "USys" Then
        DoCmd.TransferDatabase acImport, "Microsoft Access", h, acTable, td.Name, td.Name
    End If
Next

db.Close
Set fs = Nothing
Kill (h)
End Sub


Sub descargardatoscentralcopia(modulo As String)
Call eliminartablascentral(codigoLocal())
Call importartablascentral(codigoLocal())

End Sub


Function existeTabla(nomTabla As String) As Integer
Dim db As Object
Dim tdf As Object

Set db = CurrentDb

On Error GoTo TablaNoExiste
Set tdf = db.TableDefs(nomTabla)
existeTabla = 1
Set tdf = Nothing
Set db = Nothing
Exit Function

TablaNoExiste:
existeTabla = 0
Set db = Nothing
Set tdf = Nothing
End Function


Sub ReiniciarGoogleDrive()
Dim strKillCommand As String
Dim strStartCommand As String
Dim taskKillResult As Double
Dim startResult As Double

' Cierra el proceso de Google Drive (DriveFS.exe o GoogleDriveFS.exe)
strKillCommand = "taskkill /F /IM GoogleDriveFS.exe"
taskKillResult = Shell("cmd.exe /c " & strKillCommand, vbHide)

Sleep 5000

strKillCommand = "taskkill /F /IM GoogleDriveFS.exe"
taskKillResult = Shell("cmd.exe /c " & strKillCommand, vbHide)

' Espera unos segundos para asegurarse de que se cerró correctamente
Sleep 5000

' Inicia Google Drive nuevamente (ajusta la ruta si es distinta)
Dim rutadrive As String
rutadrive = BuscarRutaArchivo("GoogleDriveFS.exe", "C:\Program Files\Google\Drive File Stream\")
If rutadrive = "" Then
    MsgBox "No se encuentra archivo ejecutable de Google Drive"
Else
    strStartCommand = rutadrive
    startResult = Shell(strStartCommand, vbNormalFocus)
    MsgBox "Google Drive reiniciado correctamente"
End If

End Sub

Function fechaultimasubidaarchivodrive(archi As String) As Date
Dim fso As Object
Dim archivo As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(archi) Then
    Set archivo = fso.GetFile(archi)
    fechaultimasubidaarchivodrive = archivo.DateLastModified
Else
    fechaultimasubidaarchivodrive = 0
End If
End Function
