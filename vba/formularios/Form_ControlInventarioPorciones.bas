' ==========================================================
' Modulo  : Form_ControlInventarioPorciones
' Tipo    : 100  |  Lineas: 90
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:09
' ==========================================================

Option Compare Database

Private Sub aproducto_Change()
Me.atipo = ""
End Sub

Private Sub atipo_Change()
Me.aproducto = ""
End Sub



Private Sub Comando172_Click()
If IsNull(Me.semanaac) = True Or Me.semanaac = "" Then
    MsgBox "Ingresar numero de semana"
Else
    Me.Filter = "[tipo]='" & Me.atipo & "'"
    Me.FilterOn = True
    Me.aproducto = ""
    Me.Requery
End If
End Sub

Private Sub Comando192_Click()
If IsNull(Me.semanaac) = True Or Me.semanaac = "" Then
    MsgBox "Ingresar numero de semana"
Else
    Me.Filter = "[CodIngrediente]= '" & Me.aproducto & "'"
    Me.FilterOn = True
    Me.atipo = ""
    Me.Requery
End If
End Sub

Public Sub Comando196_Click()
Me.Requery
Me.Printer.Orientation = acPRORPortrait
Me.Printer.LeftMargin = 0
Me.Printer.RightMargin = 0
Me.Printer.BottomMargin = 0
Me.Printer.TopMargin = 0

Dim Direccion, Direccion2 As String
Dim archivo As String

If codigoLocal() = 0 Then
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo CDS\Almacen\Reporte Semanal"
Else
    Direccion = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal() & "\" & Me.semanaac
    Direccion2 = "C:\Users\" & NombreSistema() & "\Google Drive BP\Modulo Operaciones\Reporte Semanal\" & nombrelocal()
End If

If Dir(Direccion2, vbDirectory) = "" Then
    MkDir Direccion2
    MkDir Direccion
End If
If Dir(Direccion, vbDirectory) = "" Then
    MkDir Direccion
End If

archivo = Direccion & "\" & "Control Porciones Desde " & Me.semanadc & " Hasta " & Me.semanaac & ".pdf"
DoCmd.OutputTo acOutputForm, "ControlInventarioPorciones", acFormatPDF, archivo, False

On Error GoTo CerradoAutomatico
DoCmd.Close acForm, "ControlInventarioPorciones"
Exit Sub

CerradoAutomatico:
DoCmd.Close acForm, "ControlInventarioPorciones"
End Sub

Private Sub Comando331_Click()
DoCmd.OpenForm "HistorialPreIngresosLocal"
Call Forms("[HistorialPreIngresosLocal]").modocentral
End Sub

Private Sub Comando65_Click()
Me.aproducto = ""
Me.atipo = ""
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Form.Caption = "¦¦ " & nombrelocal() & " - " & ciudadsistema() & " ¦¦"
Me.ShortcutMenu = False
Me.semanaant = numerosemana(Date)
Me.semanadc = numerosemana(Date)
Me.semanaac = numerosemana(Date)
End Sub
