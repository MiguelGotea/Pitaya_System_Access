' ==========================================================
' Modulo  : InformacionNutricional
' Tipo    : 1  |  Lineas: 127
' Proyecto: Database3
' Exportado: 2026-10-07 07:17:13
' ==========================================================

Option Compare Database

Function ComponenteNutricionalEnergia(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, EnergiaKCAL FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalEnergia = rst("EnergiaKCAL")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalEnergia = 0
End Function

Function ComponenteNutricionalProteina(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, ProteinasGR FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalProteina = rst("ProteinasGR")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalProteina = 0
End Function

Function ComponenteNutricionalGrasa(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, GrasasGR FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalGrasa = rst("GrasasGR")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalGrasa = 0
End Function

Function ComponenteNutricionalCarbohidrato(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, CarbohidratosGR FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalCarbohidrato = rst("CarbohidratosGR")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalCarbohidrato = 0
End Function

Function ComponenteNutricionalFibra(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, FibraGR FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalFibra = rst("FibraGR")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalFibra = 0
End Function

Function ComponenteNutricionalAzucar(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, AzucarGR FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalAzucar = rst("AzucarGR")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalAzucar = 0
End Function

Function ComponenteNutricionalVitaminaC(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, VitCMG FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalVitaminaC = rst("VitCMG")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalVitaminaC = 0
End Function

Function ComponenteNutricionalGluten(ingr As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, Gluten FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteNutricionalGluten = rst("Gluten")
rst.Close
Exit Function
Nulo:
ComponenteNutricionalGluten = 0
End Function

Function ComponenteCaloriasAzucarReceta(bat As String) As Double
On Error GoTo Nulo
Dim rst As DAO.Recordset
Dim miSQL As String

miSQL = "SELECT CodIngrediente, Gluten FROM ComposicionIngredientes WHERE CodIngrediente='" & ingr & "'"
Set rst = CurrentDb.OpenRecordset(miSQL, dbOpenDynaset)
ComponenteCaloriasAzucarReceta = rst("Gluten")
rst.Close
Exit Function
Nulo:
ComponenteCaloriasAzucarReceta = 0
End Function
