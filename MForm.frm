VERSION 5.00
Begin VB.Form MForm 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3D VB"
   ClientHeight    =   4788
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4788
   ClipControls    =   0   'False
   Icon            =   "MForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MForm.frx":030A
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   360
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   4800
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "MForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Key As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rotation As tVector3D
Select Case KeyCode
Case vbKeyUp: Key = vbKeyUp
Case vbKeyDown: Key = vbKeyDown
Case vbKeyLeft: Key = vbKeyLeft
Case vbKeyRight: Key = vbKeyRight
Case vbKeyEscape: Key = vbKeyEscape
End Select
End Sub

Private Sub Form_Load()
Show
SForm.Hide
CreateTrigTable
Select Case DemoType
Case CUBE_WIREFRAME: DoWireframeCubeDemo
Case CUBE_FILLED: DoCubeDemo
Case CUBE_LIGHTED: DoLightedCubeDemo
End Select
End Sub

Sub DoWireframeCubeDemo()
Dim Cube As tObject6Face
Dim OriginalCube As tObject6Face
Camera.z = 300
CreateWireframeCube OriginalCube, 100

Do
DoEvents

PB.Picture = LoadPicture()
Cube = OriginalCube
For i = 1 To 8
RotatePoint Cube.P3D(i), Camera.X, Camera.Y
LensDivDistance = 256 / (Cube.P3D(i).z - Camera.z)
Cube.P2D(i).X = (Cube.P3D(i).X * LensDivDistance) + 200
Cube.P2D(i).Y = 200 - (Cube.P3D(i).Y * LensDivDistance)
Next
PB.Cls
MoveToEx PB.hdc, Cube.P2D(1).X, Cube.P2D(1).Y, p
LineTo PB.hdc, Cube.P2D(2).X, Cube.P2D(2).Y
LineTo PB.hdc, Cube.P2D(3).X, Cube.P2D(3).Y
LineTo PB.hdc, Cube.P2D(4).X, Cube.P2D(4).Y
LineTo PB.hdc, Cube.P2D(1).X, Cube.P2D(1).Y
LineTo PB.hdc, Cube.P2D(5).X, Cube.P2D(5).Y
LineTo PB.hdc, Cube.P2D(6).X, Cube.P2D(6).Y
LineTo PB.hdc, Cube.P2D(7).X, Cube.P2D(7).Y
LineTo PB.hdc, Cube.P2D(8).X, Cube.P2D(8).Y
LineTo PB.hdc, Cube.P2D(5).X, Cube.P2D(5).Y
MoveToEx PB.hdc, Cube.P2D(2).X, Cube.P2D(2).Y, p
LineTo PB.hdc, Cube.P2D(6).X, Cube.P2D(6).Y
MoveToEx PB.hdc, Cube.P2D(3).X, Cube.P2D(3).Y, p
LineTo PB.hdc, Cube.P2D(7).X, Cube.P2D(7).Y
MoveToEx PB.hdc, Cube.P2D(4).X, Cube.P2D(4).Y, p
LineTo PB.hdc, Cube.P2D(8).X, Cube.P2D(8).Y
'copy everything to the screen
BitBlt hdc, 0, 0, 400, 400, PB.hdc, 0, 0, vbSrcCopy
'display frame count
FCount = FCount + 1
    Select Case Key
    Case vbKeyUp
    Camera.X = Camera.X + 3
    If Camera.X >= 360 Then Camera.X = Camera.X - 360
    Case vbKeyDown
    Camera.X = Camera.X - 3
    If Camera.X <= 0 Then Camera.X = Camera.X + 360
    Case vbKeyLeft
    Camera.Y = Camera.Y + 3
    If Camera.Y >= 360 Then Camera.Y = Camera.Y - 360
    Case vbKeyRight
    Camera.Y = Camera.Y - 3
    If Camera.Y <= 0 Then Camera.Y = Camera.Y + 360
    End Select
Loop Until Key = vbKeyEscape
End
End Sub

Sub DoCubeDemo()
Dim TempPoly As tPolygon4Point
Dim Polygon(1 To 6) As tPolygon4Point
Dim OriginalPolygon(1 To 6) As tPolygon4Point
Camera.z = 400
CreateCubeFrom6Polygons OriginalPolygon(), 100
BitBlt hdc, 0, 0, 400, 400, PB.hdc, 0, 0, vbSrcCopy

PB.Picture = LoadPicture()
Do
DoEvents
PB.Cls

For i = 1 To 6
    Polygon(i) = OriginalPolygon(i)
For i2 = 1 To 4
    RotatePoint Polygon(i).P3D(i2), Camera.X, Camera.Y
    LensDivDistance = 256 / (Polygon(i).P3D(i2).z - Camera.z)
    Polygon(i).P2D(i2).X = (Polygon(i).P3D(i2).X * LensDivDistance) + 200
    Polygon(i).P2D(i2).Y = 200 - (Polygon(i).P3D(i2).Y * LensDivDistance)
Next
Next

'sort out the Z-order
For i = 1 To 6
For i2 = i + 1 To 6
  Select Case (Polygon(i).P3D(1).z + Polygon(i).P3D(2).z + Polygon(i).P3D(3).z + Polygon(i).P3D(4).z)
    Case Is >= (Polygon(i2).P3D(1).z + Polygon(i2).P3D(2).z + Polygon(i2).P3D(3).z + Polygon(i2).P3D(4).z)
        Select Case i
           Case Is > i2
           Case Else
           TempPoly = Polygon(i)
           Polygon(i) = Polygon(i2)
           Polygon(i2) = TempPoly
        End Select
    Case Is < (Polygon(i2).P3D(1).z + Polygon(i2).P3D(2).z + Polygon(i2).P3D(3).z + Polygon(i2).P3D(4).z)
        Select Case i
           Case Is < i2
           Case Else
           TempPoly = Polygon(i)
           Polygon(i) = Polygon(i2)
           Polygon(i2) = TempPoly
        End Select
End Select
Next
Next

'draw each polygon, from back to front of world
For i = 1 To 6
TempColor = Polygon(i).Color
PB.Forecolor = TempColor
PB.FillColor = TempColor
MoveToEx PB.hdc, Polygon(i).P2D(1).X, Polygon(i).P2D(1).Y, p
LineTo PB.hdc, Polygon(i).P2D(2).X, Polygon(i).P2D(2).Y
LineTo PB.hdc, Polygon(i).P2D(3).X, Polygon(i).P2D(3).Y
LineTo PB.hdc, Polygon(i).P2D(4).X, Polygon(i).P2D(4).Y
LineTo PB.hdc, Polygon(i).P2D(1).X, Polygon(i).P2D(1).Y
FloodFill PB.hdc, (Polygon(i).P2D(1).X + Polygon(i).P2D(2).X + Polygon(i).P2D(3).X + Polygon(i).P2D(4).X) \ 4, (Polygon(i).P2D(1).Y + Polygon(i).P2D(2).Y + Polygon(i).P2D(3).Y + Polygon(i).P2D(4).Y) \ 4, TempColor
Next

'copy everything to the screen
BitBlt hdc, 70, 70, 260, 260, PB.hdc, 70, 70, vbSrcCopy
'display frame count
FCount = FCount + 1

    Select Case Key
    Case vbKeyUp
    Camera.X = Camera.X + 7
    If Camera.X >= 360 Then Camera.X = Camera.X - 360
    Case vbKeyDown
    Camera.X = Camera.X - 7
    If Camera.X <= 0 Then Camera.X = Camera.X + 360
    Case vbKeyLeft
    Camera.Y = Camera.Y + 7
    If Camera.Y >= 360 Then Camera.Y = Camera.Y - 360
    Case vbKeyRight
    Camera.Y = Camera.Y - 7
    If Camera.Y <= 0 Then Camera.Y = Camera.Y + 360
    End Select

Loop Until Key = vbKeyEscape
End
End Sub

Sub DoLightedCubeDemo()
Dim TempPoly As tLightedPolygon4Point
Dim Polygon(1 To 6) As tLightedPolygon4Point
Dim OriginalPolygon(1 To 6) As tLightedPolygon4Point
Dim Shadow As Single

Camera.z = 300
CreateLightedCubeFrom6Polygons OriginalPolygon(), 100

PB.Picture = Picture
Do
DoEvents
PB.Cls

For i = 1 To 6
    Polygon(i) = OriginalPolygon(i)
For i2 = 1 To 4
    RotatePoint Polygon(i).P3D(i2), Camera.X, Camera.Y
    LensDivDistance = 256 / (Polygon(i).P3D(i2).z - Camera.z)
    Polygon(i).P2D(i2).X = (Polygon(i).P3D(i2).X * LensDivDistance) + 200
    Polygon(i).P2D(i2).Y = 200 - (Polygon(i).P3D(i2).Y * LensDivDistance)
Next
Next

'sort out the Z-order
For i = 1 To 6
For i2 = i + 1 To 6
  Select Case (Polygon(i).P3D(1).z + Polygon(i).P3D(2).z + Polygon(i).P3D(3).z + Polygon(i).P3D(4).z)
    Case Is >= (Polygon(i2).P3D(1).z + Polygon(i2).P3D(2).z + Polygon(i2).P3D(3).z + Polygon(i2).P3D(4).z)
        Select Case i
           Case Is > i2
           Case Else
           TempPoly = Polygon(i)
           Polygon(i) = Polygon(i2)
           Polygon(i2) = TempPoly
        End Select
    Case Is < (Polygon(i2).P3D(1).z + Polygon(i2).P3D(2).z + Polygon(i2).P3D(3).z + Polygon(i2).P3D(4).z)
        Select Case i
           Case Is < i2
           Case Else
           TempPoly = Polygon(i)
           Polygon(i) = Polygon(i2)
           Polygon(i2) = TempPoly
        End Select
End Select
Next
Next

'draw each polygon, from back to front of world
For i = 1 To 6
    TempRGBcolor = Polygon(i).Color
    Shadow = (Polygon(i).P2D(1).Y + Polygon(i).P2D(2).Y + Polygon(i).P2D(3).Y + Polygon(i).P2D(4).Y) / 370
    TempColor = RGB(TempRGBcolor.R / Shadow, TempRGBcolor.G / Shadow, TempRGBcolor.B / Shadow)
    PB.Forecolor = Polygon(i).Forecolor
    PB.FillColor = TempColor
    MoveToEx PB.hdc, Polygon(i).P2D(1).X, Polygon(i).P2D(1).Y, p
    LineTo PB.hdc, Polygon(i).P2D(2).X, Polygon(i).P2D(2).Y
    LineTo PB.hdc, Polygon(i).P2D(3).X, Polygon(i).P2D(3).Y
    LineTo PB.hdc, Polygon(i).P2D(4).X, Polygon(i).P2D(4).Y
    LineTo PB.hdc, Polygon(i).P2D(1).X, Polygon(i).P2D(1).Y
    ExtFloodFill PB.hdc, (Polygon(i).P2D(1).X + Polygon(i).P2D(2).X + Polygon(i).P2D(3).X + Polygon(i).P2D(4).X) \ 4, (Polygon(i).P2D(1).Y + Polygon(i).P2D(2).Y + Polygon(i).P2D(3).Y + Polygon(i).P2D(4).Y) \ 4, PB.Forecolor, 0
Next

'copy everything to the screen
BitBlt hdc, 0, 0, 400, 400, PB.hdc, 0, 0, vbSrcCopy
'display frame count
FCount = FCount + 1

    Select Case Key
    Case vbKeyUp
    Camera.X = Camera.X + 5
    If Camera.X >= 360 Then Camera.X = Camera.X - 360
    Case vbKeyDown
    Camera.X = Camera.X - 5
    If Camera.X <= 0 Then Camera.X = Camera.X + 360
    Case vbKeyLeft
    Camera.Y = Camera.Y + 5
    If Camera.Y >= 360 Then Camera.Y = Camera.Y - 360
    Case vbKeyRight
    Camera.Y = Camera.Y - 5
    If Camera.Y <= 0 Then Camera.Y = Camera.Y + 360
    End Select

Loop Until Key = vbKeyEscape
End
End Sub

Private Sub Timer1_Timer()
Caption = "3D VB - FPS : " & FCount
FCount = 0
End Sub
