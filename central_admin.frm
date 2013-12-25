VERSION 5.00
Begin VB.Form central_admin 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Departmental Administrator"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Event Head"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Student Database"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Faculty Database"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   7170
      Left            =   9360
      Picture         =   "central_admin.frx":0000
      Top             =   2640
      Width           =   11970
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "central_admin.frx":45C12
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   18840
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "M.I.T. College of Engineering"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display Name"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "central_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public c_a_n As String ' central admin name , central admin id
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub command4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub LABEL2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label2.ForeColor = vbRed
    
End Sub


Private Sub command6_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub Command1_Click()
faculty_dept.n = c_a_n
login.Hide
faculty_dept.Show
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub Command2_Click()
login.Hide
Me.Hide
class_division.name_div = c_a_n
class_division.Show
End Sub

Private Sub Command3_Click()
login.Hide
event_admin.mode = 1   ' event head
Me.Hide
event_admin.Show

End Sub

Private Sub Command4_Click()
login.Hide
event_admin.mode = 2  'dept admin
Me.Hide
event_admin.Show
End Sub

Private Sub Command5_Click()
End
End Sub



Private Sub Command6_Click()
password.Show
End Sub

Private Sub Form_Activate()
Label1.Caption = "Hello " + c_a_n + "!"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbGreen
End Sub

Private Sub Label2_Click()
End
End Sub
