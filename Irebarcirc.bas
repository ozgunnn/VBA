Attribute VB_Name = "Module1"

Function IREBARCIRC(RadiusOfRebarSet, NoOfRebars, DiaOfRebars)
Dim i As Integer
Dim rr As Double
Dim nr As Integer
Dim dr As Double
Dim dummy As Double

rr = RadiusOfRebarSet
nr = NoOfRebars
dr = DiaOfRebars
Pi = WorksheetFunction.Pi()

i = 1
dummy = 0

Do While i <= nr
    dummy = dummy + Pi / 4 * (dr / 2) ^ 4 + (Pi * ((dr) ^ 2) / 4) * (Abs(rr * Sin(2 * Pi * (i - 1) / nr))) ^ 2
    i = i + 1
Loop

IREBARCIRC = dummy
End Function

