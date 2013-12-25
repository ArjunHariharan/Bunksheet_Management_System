VERSION 5.00
Begin VB.Form faculty_event_admin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "select login type"
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19995
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   19995
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   2145
      Left            =   16320
      Picture         =   "faculty_event_admin.frx":0000
      ScaleHeight     =   2085
      ScaleWidth      =   2790
      TabIndex        =   7
      Top             =   4560
      Width           =   2850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Administrator"
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
      Left            =   9360
      TabIndex        =   6
      Top             =   3720
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   13680
      Picture         =   "faculty_event_admin.frx":17ED
      ScaleHeight     =   2355
      ScaleWidth      =   2700
      TabIndex        =   5
      Top             =   1320
      Width           =   2760
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "faculty_event_admin.frx":31FE
      ScaleHeight     =   1440
      ScaleWidth      =   1230
      TabIndex        =   4
      Top             =   120
      Width           =   1290
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change password"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Faculty"
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
      Left            =   9360
      TabIndex        =   0
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   13920
      Picture         =   "faculty_event_admin.frx":3ED0
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   2775
      Left            =   4440
      Picture         =   "faculty_event_admin.frx":49589C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18120
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   840
      Picture         =   "faculty_event_admin.frx":498BE7
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login As"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   8880
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
End
Attribute VB_Name = "faculty_event_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mode As Integer
Public f_name, f_id As String
Public f_dpt, f_eve As Integer
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbGreen
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub label3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label3.ForeColor = vbRed
End Sub

Private Sub Command1_Click()

If mode = 1 Then
    bunksheet.fac_id = f_id
    bunksheet.fac_name = f_name
    bunksheet.event_id = f_eve
    bunksheet.fac_type = 1
    bunksheet.Show
Else  'admin mode
    d_admin.d_dpt = f_dpt
    d_admin.d_id = f_id
    d_admin.d_name = f_name
    d_admin.Show
End If
Me.Hide
faculty_home.Hide
bunksheet_display.Hide
End Sub

Private Sub Command2_Click()
faculty_home.fc_type = 1
faculty_home.fc_name = f_name
faculty_home.fc_dpt = f_dpt
Me.Hide
faculty_home.Show
bunksheet.Hide
d_admin.Hide
bunksheet_display.Hide
End Sub

Private Sub Command3_Click()
password.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Activate()

If mode = 2 Then  'department admin
    Command1.Caption = "Department Administrator"
    Command1.Font.Size = 12
    Command1.FontBold = True
Else
    Command1.Caption = "Event Head"
End If
Label1.Caption = "Hello " + f_name + "!"
End Sub

