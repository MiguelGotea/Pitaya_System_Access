' ==========================================================
' Modulo  : ExportarTablaEstructuraSQL
' Tipo    : 1
' Lineas  : 86
' Proyecto: Database3
' Exportado: 2026-10-07 06:21:30
' ==========================================================
Option Compare Database

Sub ExportarDDL()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim sql As String
    Dim Linea As String
    Dim archivo As Integer
    
    Set db = CurrentDb()
    archivo = FreeFile
    Open "C:\Users\" & NombreSistema() & "\Desktop\Sistema\Estructura_DB_PitayaSystem_Access.sql" For Output As #archivo
    
    For Each tdf In db.TableDefs
        ' Saltar tablas del sistema
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            sql = "CREATE TABLE [" & tdf.Name & "] (" & vbCrLf
            Dim campos() As String
            ReDim campos(tdf.Fields.Count - 1)
            
            Dim I As Integer
            I = 0
            For Each fld In tdf.Fields
                Linea = "    [" & fld.Name & "] "
                
                Select Case fld.Type
                    Case 1:  Linea = Linea & "TINYINT"       ' Boolean
                    Case 2:  Linea = Linea & "TINYINT"       ' Byte
                    Case 3:  Linea = Linea & "INT"           ' Integer (entero normal 16-bit)
                    Case 4:  Linea = Linea & "INT"           ' Long Integer (entero largo 32-bit)
                    Case 5:  Linea = Linea & "DOUBLE"        ' Single
                    Case 6:  Linea = Linea & "DOUBLE"        ' Double
                    Case 7:  Linea = Linea & "DOUBLE"        ' Currency
                    Case 8:
                        Dim fmt As String
                        fmt = ""
                        On Error Resume Next
                        fmt = fld.Properties("Format").Value
                        On Error GoTo 0
                        
                        Select Case LCase(fmt)
                            Case "short date", "fecha corta", "d/m/yyyy", "dd/mm/yyyy", "m/d/yyyy"
                                Linea = Linea & "DATE"
                            Case "long date", "fecha larga"
                                Linea = Linea & "DATE"
                            Case "short time", "hora corta", "h:nn", "hh:nn"
                                Linea = Linea & "TIME"
                            Case "long time", "hora larga", "h:nn:ss", "hh:nn:ss"
                                Linea = Linea & "TIME"
                            Case "general date", "fecha general", ""
                                Linea = Linea & "DATETIME"
                            Case Else
                                ' Si tiene hora en el formato, es DATETIME o TIME
                                If InStr(LCase(fmt), "h") > 0 And InStr(LCase(fmt), "d") > 0 Then
                                    Linea = Linea & "DATETIME"
                                ElseIf InStr(LCase(fmt), "h") > 0 Then
                                    Linea = Linea & "TIME"
                                ElseIf InStr(LCase(fmt), "d") > 0 Or InStr(LCase(fmt), "m") > 0 Or InStr(LCase(fmt), "y") > 0 Then
                                    Linea = Linea & "DATE"
                                Else
                                    Linea = Linea & "DATETIME"  ' default seguro
                                End If
                        End Select
                    Case 10: Linea = Linea & "VARCHAR(" & fld.size & ")"  ' Text
                    Case 11: Linea = Linea & "BIT"           ' OLE (aprox)
                    Case 12: Linea = Linea & "TEXT"          ' Memo
                    Case 15: Linea = Linea & "UNIQUEIDENTIFIER" ' GUID
                    Case 16: Linea = Linea & "DOUBLE" ' Decimal
                    Case Else: Linea = Linea & "VARCHAR(255)"
                End Select
                
                If (fld.Attributes And 16) Then Linea = Linea & " NOT NULL"
                
                campos(I) = Linea
                I = I + 1
            Next fld
            
            sql = sql & Join(campos, "," & vbCrLf) & vbCrLf & ");" & vbCrLf & vbCrLf
            Print #archivo, sql
        End If
    Next tdf
    
    Close #archivo
    MsgBox "Exportado a C:\Estructura_DB_PitayaSystem_Access.sql"
End Sub
