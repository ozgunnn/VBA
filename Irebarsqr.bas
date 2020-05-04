Attribute VB_Name = "Module2"
Function IREBARSQR(HRebarSet, NoOfRebars, DiaOfRebars)

Dim i As Integer
Dim a As Double
Dim h As Double
Dim nr As Integer
Dim noly As Integer
Dim dr As Double
Dim dummy As Double

h = HRebarSet
nr = NoOfRebars
dr = DiaOfRebars
Pi = WorksheetFunction.Pi()

i = 1
dummy = 0
a = h / (nr / 4)
noly = 1 + nr / 4

ReDim norebar(noly) As Double
ReDim Alayer(noly) As Double
ReDim I0layer(noly) As Double
ReDim Depth(noly) As Double
ReDim Dist(noly) As Double

Do While i <= noly
    If i = 1 Then
    norebar(i) = 1 + nr / 4
    ElseIf i = noly Then
    norebar(i) = 1 + nr / 4
    Else
    norebar(i) = 2
    End If
    i = i + 1
Loop
i = 1

Do While i <= noly
    Alayer(i) = Pi * (dr) ^ 2 / 4 * norebar(i)
    I0layer(i) = norebar(i) * Pi * (dr / 2) ^ 4 / 4
    Depth(i) = (i - 1) * a
    Dist(i) = Abs(h / 2 - Depth(i))
    i = i + 1
Loop
i = 1

Do While i <= noly
    dummy = dummy + I0layer(i) + Alayer(i) * (Dist(i)) ^ 2
    i = i + 1
Loop
    
IREBARSQR = dummy
End Function

