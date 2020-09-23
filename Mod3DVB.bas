Attribute VB_Name = "Mod3DVB"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Type POINTAPI
  X As Integer
  Y As Integer
End Type

Public Type tVector3D
  X As Integer
  Y As Integer
  z As Integer
End Type

Type tRGBcolor
  R As Byte
  G As Byte
  B As Byte
End Type

Public Type tPolygon4Point
  P3D(1 To 4) As tVector3D
  P2D(1 To 4) As POINTAPI
  Color As Long
End Type

Public Type tLightedPolygon4Point
  P3D(1 To 4) As tVector3D
  P2D(1 To 4) As POINTAPI
  Color As tRGBcolor
  Forecolor As ColorConstants
End Type

Public Type tObject6Face
  P3D(1 To 8) As tVector3D
  P2D(1 To 8) As POINTAPI
  Color(1 To 6) As Long
End Type

Public Sine(0 To 361) As Double
Public Cosine(0 To 361) As Double
Public FCount As Integer
Public Camera As tVector3D
Public TempColor As Long
Public TempRGBcolor As tRGBcolor
Public p As POINTAPI
Public TempP3D As tVector3D
Public DemoType As Byte
'types of demo
Public Const CUBE_WIREFRAME = 0
Public Const CUBE_FILLED = 1
Public Const CUBE_LIGHTED = 2

Public Const PI = 3.14159265358979 'obvious
Public Const CX = 200
Public Const CY = 200

Public Sub RotatePoint(p As tVector3D, X, Y)

'rotate on x-axis
TempP3D.Y = (p.Y * Cosine(X)) - (p.z * Sine(X))
p.z = (p.z * Cosine(X)) + (p.Y * Sine(X))
'copy new point
p.Y = TempP3D.Y

'rotate on y-axis
TempP3D.z = (p.z * Cosine(Y)) - (p.X * Sine(Y))
p.X = (p.X * Cosine(Y)) + (p.z * Sine(Y))
'copy new point
p.z = TempP3D.z

End Sub

Public Sub RotatePointX(p As tVector3D, rotation As Integer)
'rotate on x-axis
TempP3D.Y = (p.Y * Cosine(rotation)) - (p.z * Sine(rotation))
p.z = (p.z * Cosine(rotation)) + (p.Y * Sine(rotation))
'copy new point
p.Y = TempP3D.Y

End Sub

Public Sub RotatePointY(p As tVector3D, rotation As Integer)
'rotate on y-axis
TempP3D.z = (p.z * Cosine(rotation)) - (p.X * Sine(rotation))
p.X = (p.X * Cosine(rotation)) + (p.z * Sine(rotation))
'copy new point
p.z = TempP3D.z

End Sub

Public Sub RotatePointZ(p As tVector3D, rotation As Integer)

'rotate on z-axis
p.X = (p.X * Cosine(rotation)) - (p.Y * Sine(rotation))
p.Y = (p.Y * Cosine(rotation)) + (p.X * Sine(rotation))

End Sub


Public Sub PlotPoint(p As tVector3D, ScrP As POINTAPI)
'this converts a 3D co-ord to 2D co-ord to be plotted on the screen
On Error Resume Next
LensDivDistance = 256 / (p.z - Camera.z)
ScrP.X = (p.X * LensDivDistance) + 200
ScrP.Y = 200 - (p.Y * LensDivDistance)
End Sub

