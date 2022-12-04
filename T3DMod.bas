Attribute VB_Name = "Module1"

Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
   
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R%, G%, B%, R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
R = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B3 = B1
    B4 = B2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 3 Then 'InsetInset
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = B2
End If
If Style3D = 4 Then 'No Border
R1 = R: R2 = R: R3 = R: R4 = R
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: B2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

