Attribute VB_Name = "Auxiliar"
Option Explicit


'//////////////////////////////////////////////////
'/              Auxiliar functions                /
'/                                                /
'//////////////////////////////////////////////////

Public Function Redondear(dblnToR As Double, Optional intCntDec As Integer) As Double
Dim dblPot As Double
Dim dblF As Double
If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
dblPot = 10 ^ intCntDec
Redondear = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot
End Function



