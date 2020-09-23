Attribute VB_Name = "CreateStuff"
Public Sub CreateWireframeCube(Cube As tObject6Face, Size As Integer)
'creates a object6face cube based on the input
Cube.P3D(1).x = -Size
Cube.P3D(1).y = Size
Cube.P3D(1).z = Size
Cube.P3D(2).x = Size
Cube.P3D(2).y = Size
Cube.P3D(2).z = Size
Cube.P3D(3).x = Size
Cube.P3D(3).y = -Size
Cube.P3D(3).z = Size
Cube.P3D(4).x = -Size
Cube.P3D(4).y = -Size
Cube.P3D(4).z = Size
Cube.P3D(5).x = -Size
Cube.P3D(5).y = Size
Cube.P3D(5).z = -Size
Cube.P3D(6).x = Size
Cube.P3D(6).y = Size
Cube.P3D(6).z = -Size
Cube.P3D(7).x = Size
Cube.P3D(7).y = -Size
Cube.P3D(7).z = -Size
Cube.P3D(8).x = -Size
Cube.P3D(8).y = -Size
Cube.P3D(8).z = -Size
Cube.Color(1) = vbRed
Cube.Color(2) = vbGreen
Cube.Color(3) = vbBlue
Cube.Color(4) = vbMagenta
Cube.Color(5) = vbYellow
Cube.Color(6) = vbCyan
End Sub

Public Sub CreateCubeFrom6Polygons(Polygon() As tPolygon4Point, Size As Integer)
Dim Cube As tObject6Face
'creates a cube from 6 polygons based on the input
'first make a temporary cube object
Cube.P3D(1).x = -Size
Cube.P3D(1).y = Size
Cube.P3D(1).z = Size
Cube.P3D(2).x = Size
Cube.P3D(2).y = Size
Cube.P3D(2).z = Size
Cube.P3D(3).x = Size
Cube.P3D(3).y = -Size
Cube.P3D(3).z = Size
Cube.P3D(4).x = -Size
Cube.P3D(4).y = -Size
Cube.P3D(4).z = Size
Cube.P3D(5).x = -Size
Cube.P3D(5).y = Size
Cube.P3D(5).z = -Size
Cube.P3D(6).x = Size
Cube.P3D(6).y = Size
Cube.P3D(6).z = -Size
Cube.P3D(7).x = Size
Cube.P3D(7).y = -Size
Cube.P3D(7).z = -Size
Cube.P3D(8).x = -Size
Cube.P3D(8).y = -Size
Cube.P3D(8).z = -Size
Cube.Color(1) = vbRed
Cube.Color(2) = vbGreen
Cube.Color(3) = vbBlue
Cube.Color(4) = vbMagenta
Cube.Color(5) = vbYellow
Cube.Color(6) = vbCyan

'then copy the cube attributes to the 6 polygons
Polygon(1).P3D(1) = Cube.P3D(1)
Polygon(1).P3D(2) = Cube.P3D(2)
Polygon(1).P3D(3) = Cube.P3D(3)
Polygon(1).P3D(4) = Cube.P3D(4)

Polygon(2).P3D(1) = Cube.P3D(6)
Polygon(2).P3D(2) = Cube.P3D(5)
Polygon(2).P3D(3) = Cube.P3D(8)
Polygon(2).P3D(4) = Cube.P3D(7)

Polygon(3).P3D(1) = Cube.P3D(2)
Polygon(3).P3D(2) = Cube.P3D(6)
Polygon(3).P3D(3) = Cube.P3D(7)
Polygon(3).P3D(4) = Cube.P3D(3)

Polygon(4).P3D(1) = Cube.P3D(5)
Polygon(4).P3D(2) = Cube.P3D(1)
Polygon(4).P3D(3) = Cube.P3D(4)
Polygon(4).P3D(4) = Cube.P3D(8)

Polygon(5).P3D(1) = Cube.P3D(1)
Polygon(5).P3D(2) = Cube.P3D(5)
Polygon(5).P3D(3) = Cube.P3D(6)
Polygon(5).P3D(4) = Cube.P3D(2)

Polygon(6).P3D(1) = Cube.P3D(8)
Polygon(6).P3D(2) = Cube.P3D(4)
Polygon(6).P3D(3) = Cube.P3D(3)
Polygon(6).P3D(4) = Cube.P3D(7)

For i = 1 To 6
Polygon(i).Color = Cube.Color(i)
Next

End Sub

Public Sub CreateLightedCubeFrom6Polygons(Polygon() As tLightedPolygon4Point, Size As Integer)
Dim Cube As tObject6Face
'creates a cube from 6 polygons based on the input
'first make a temporary cube object
Cube.P3D(1).x = -Size
Cube.P3D(1).y = Size
Cube.P3D(1).z = Size
Cube.P3D(2).x = Size
Cube.P3D(2).y = Size
Cube.P3D(2).z = Size
Cube.P3D(3).x = Size
Cube.P3D(3).y = -Size
Cube.P3D(3).z = Size
Cube.P3D(4).x = -Size
Cube.P3D(4).y = -Size
Cube.P3D(4).z = Size
Cube.P3D(5).x = -Size
Cube.P3D(5).y = Size
Cube.P3D(5).z = -Size
Cube.P3D(6).x = Size
Cube.P3D(6).y = Size
Cube.P3D(6).z = -Size
Cube.P3D(7).x = Size
Cube.P3D(7).y = -Size
Cube.P3D(7).z = -Size
Cube.P3D(8).x = -Size
Cube.P3D(8).y = -Size
Cube.P3D(8).z = -Size

'then copy the cube attributes to the 6 polygons
Polygon(1).P3D(1) = Cube.P3D(1)
Polygon(1).P3D(2) = Cube.P3D(2)
Polygon(1).P3D(3) = Cube.P3D(3)
Polygon(1).P3D(4) = Cube.P3D(4)

Polygon(2).P3D(1) = Cube.P3D(6)
Polygon(2).P3D(2) = Cube.P3D(5)
Polygon(2).P3D(3) = Cube.P3D(8)
Polygon(2).P3D(4) = Cube.P3D(7)

Polygon(3).P3D(1) = Cube.P3D(2)
Polygon(3).P3D(2) = Cube.P3D(6)
Polygon(3).P3D(3) = Cube.P3D(7)
Polygon(3).P3D(4) = Cube.P3D(3)

Polygon(4).P3D(1) = Cube.P3D(5)
Polygon(4).P3D(2) = Cube.P3D(1)
Polygon(4).P3D(3) = Cube.P3D(4)
Polygon(4).P3D(4) = Cube.P3D(8)

Polygon(5).P3D(1) = Cube.P3D(1)
Polygon(5).P3D(2) = Cube.P3D(5)
Polygon(5).P3D(3) = Cube.P3D(6)
Polygon(5).P3D(4) = Cube.P3D(2)

Polygon(6).P3D(1) = Cube.P3D(8)
Polygon(6).P3D(2) = Cube.P3D(4)
Polygon(6).P3D(3) = Cube.P3D(3)
Polygon(6).P3D(4) = Cube.P3D(7)

Polygon(1).Color.R = 255
Polygon(1).Color.G = 100
Polygon(1).Color.B = 100

Polygon(2).Color.R = 100
Polygon(2).Color.G = 255
Polygon(2).Color.B = 100

Polygon(3).Color.R = 100
Polygon(3).Color.G = 100
Polygon(3).Color.B = 255

Polygon(4).Color.R = 255
Polygon(4).Color.G = 100
Polygon(4).Color.B = 255

Polygon(5).Color.R = 255
Polygon(5).Color.G = 255
Polygon(5).Color.B = 100

Polygon(6).Color.R = 100
Polygon(6).Color.G = 255
Polygon(6).Color.B = 255

Polygon(1).Forecolor = vbRed
Polygon(2).Forecolor = vbGreen
Polygon(3).Forecolor = vbBlue
Polygon(4).Forecolor = vbMagenta
Polygon(5).Forecolor = vbYellow
Polygon(6).Forecolor = vbCyan
End Sub


Sub SortZOrder() 'this algo sorts out the Z-Order!!!!!
End Sub

Public Sub CreateTrigTable()
For i = 0 To 361
Sine(i) = Sin(i / 180 * PI)
Cosine(i) = Cos(i / 180 * PI)
Next
End Sub

