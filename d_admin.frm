VERSION 5.00
Begin VB.Form d_admin 
   BackColor       =   &H80000005&
   Caption         =   "Department Administrator"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16335
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   16335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FACULTY SUBJECTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7440
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TIME TABLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "< < BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE BUNKSHEET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   11880
      Picture         =   "d_admin.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   14760
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   960
      Picture         =   "d_admin.frx":28AD1A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department Administrator"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "d_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Public d_id, d_name As String ' d_id has the faculty id(not required for timetable),d_name is just to display the name of the user
Public d_dpt As Integer  'd_dpt has the dept no of the faculty used in timetable and other forms related to dept admin
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

Private Sub Command3_Click()
faculty_event_admin.Show
Me.Hide
End Sub

Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub command4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label4.ForeColor = vbRed
End Sub
Private Sub command5_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlack 'change it back to original color
End Sub



Private Sub Label4_Click()
End
End Sub

Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label4.ForeColor = &H8000000D 'color on hovering
End Sub


Private Sub Command1_Click()
con.Close
password.Show
End Sub

Private Sub Command2_Click()
Set cmd = New ADODB.Command
   With cmd
        .ActiveConnection = con
        .CommandText = "attendance_upd"
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_dpt", adNumeric, adParamInput, 30, d_dpt)
   End With
cmd.Execute
MsgBox ("Attendance updated")
End Sub

Private Sub Command4_Click()
' me.hide
'call timetable. put all the commands on timetable form load in timetable activate
 ' add a done button in timetable or practical table which comes back to this form
'put a back button in all d forms. link all d timetable related forms(kulz forms)
'put a logout button in each form.
End Sub

'Private Sub Command6_Click()
'Label1.Caption = "Wassup Sir?"
'End
'End Sub

Private Sub Form_Activate()
password.Hide
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"

Label2.Caption = "Hello " + d_name + "!"
End Sub

