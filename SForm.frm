VERSION 5.00
Begin VB.Form SForm 
   BackColor       =   &H00000000&
   Caption         =   "3D VB - Visit www.hispalace.fsbusiness.co.uk"
   ClientHeight    =   4536
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4104
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit 3D VB"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1452
   End
   Begin VB.CommandButton cmdLighted 
      Caption         =   "Lighting demo"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1452
   End
   Begin VB.CommandButton cmdFilled 
      Caption         =   "Flat color demo"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1452
   End
   Begin VB.CommandButton cmdWireframe 
      Caption         =   "Wireframe demo"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1452
   End
   Begin VB.Label Info 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"SForm.frx":0000
      Height          =   1752
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   2052
   End
   Begin VB.Image Logo 
      Height          =   2400
      Left            =   120
      Picture         =   "SForm.frx":00BE
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TriedIt As Boolean

Private Sub cmdExit_Click()
MsgBox "To see loads of quality VB games - 3D and 2D, visit my website : www.hispalace.fsbusiness.co.uk", vbApplicationModal + vbInformation, "Thankyou for using 3D VB"
Unload Me
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If TriedIt Then
Info = "Exits 3D VB. To see loads of quality VB games - both 3D and 2D, visit my website : www.hispalace.fsbusiness.co.uk. Also, if you have anything you would like to be put on the site, you can e-mail me : KingSimon@Hispalace.fsbusiness.co.uk"
Else
Info = "Don't quit yet! - you haven't even tried the program!"
End If
End Sub

Private Sub cmdFilled_Click()
DemoType = CUBE_FILLED
MForm.Visible = True
Unload Me
End Sub

Private Sub cmdFilled_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Info = "Use the arrow keys to spin around a 3D cube, with backface removal and flat color. On my 400Mhz PC, the frame rate is usually in excess of 20 frames/sec (FPS). Try this if you have an average PC."
End Sub

Private Sub cmdLighted_Click()
DemoType = CUBE_LIGHTED
MForm.Visible = True
Unload Me
End Sub

Private Sub cmdLighted_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Info = "Use the arrow keys to spin around a 3D cube, with backface removal and lighting effects (many shades of color). On my 400Mhz PC, the frame rate is usually in excess of 10 frames/sec (FPS). Try this if you have an fast PC."
End Sub

Private Sub cmdWireframe_Click()
DemoType = CUBE_WIREFRAME
MForm.Visible = True
Unload Me
End Sub

Private Sub cmdWireframe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Info = "Use the arrow keys to spin around a wireframe 3D cube. On my 400Mhz PC, the frame rate is usually in excess of 50 frames/sec (FPS). Try this if you have a slow PC."
End Sub

Private Sub Form_Unload(Cancel As Integer)
TriedIt = True
Hide
End Sub
