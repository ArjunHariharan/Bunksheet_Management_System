VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   DrawWidth       =   5
   FillColor       =   &H00FFFF80&
   FillStyle       =   2  'Horizontal Line
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      TabIndex        =   6
      Top             =   7680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15600
      TabIndex        =   5
      Top             =   6000
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15600
      TabIndex        =   4
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   1440
      Left            =   120
      Picture         =   "login.frx":0000
      Top             =   120
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   5010
      Left            =   1920
      Picture         =   "login.frx":0CD2
      Top             =   3480
      Width           =   7500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   11160
      TabIndex        =   3
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   11280
      TabIndex        =   2
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bunksheet Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Top             =   1680
      Width           =   11295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MIT College of Engineering"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   10095
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As New ADODB.Recordset
Public f_id As String
Public admin_type, event_head_type, dept As Integer
Public f_name As String
Dim flag As Integer
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
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox ("Enter Username!")
ElseIf Text2.Text = "" Then
MsgBox ("Enter Password!")
Else
    con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = "login_verify"   ' calls the stored procedure to verify the username and password
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_name", adVarChar, adParamInput, 30, Text1.Text)
        .Parameters.Append cmd.CreateParameter("v_password", adVarChar, adParamInput, 30, Text2.Text)
        .Parameters.Append cmd.CreateParameter("v_id", adVarChar, adParamOutput, 10)
        .Parameters.Append cmd.CreateParameter("v_event", adInteger, adParamOutput, 1)
        .Parameters.Append cmd.CreateParameter("v_admin", adInteger, adParamOutput, 1)
        .Parameters.Append cmd.CreateParameter("v_fname", adVarChar, adParamOutput, 50)
        .Parameters.Append cmd.CreateParameter("v_dept", adInteger, adParamOutput, 2)
    End With

        cmd.Execute
        f_name = cmd("v_fname")
        event_head_type = cmd("v_event")
        admin_type = cmd("v_admin")
        f_id = cmd("v_id")
        dept = cmd("v_dept")

        cmd.ActiveConnection = Nothing
        con.Close
        If f_id = "xxxxx" Then
            MsgBox ("Invalid username and password!!")
        Else
            password.pass_change_id = f_id
            If admin_type <> 0 Then  'admin
                If admin_type = 9 Then   'central admin
                    'central_admin.c_a_id = f_id
                    central_admin.c_a_n = f_name
'                    con.Close
                    Me.Hide
                    central_admin.Show
                Else 'dept admin
                    faculty_event_admin.mode = 2
                    faculty_event_admin.f_dpt = dept
                    faculty_event_admin.f_id = f_id
                    faculty_event_admin.f_name = f_name
                    faculty_event_admin.f_eve = admin_type
'                    con.Close
                    Me.Hide
                    faculty_event_admin.Show
                End If
            
            ElseIf event_head_type <> 0 Then  'event head
                faculty_event_admin.mode = 1
                faculty_event_admin.f_dpt = dept
                faculty_event_admin.f_id = f_id
                faculty_event_admin.f_name = f_name
                faculty_event_admin.f_eve = event_head_type
'                con.Close
                Me.Hide
                faculty_event_admin.Show
                
            ElseIf event_head_type = 0 And admin_type = 0 Then   'normal faculty
                faculty_home.fc_name = f_name
                faculty_home.fc_dpt = dept
                faculty_home.fc_id = f_id
                faculty_home.fc_type = 0
                faculty_home.Show
'                con.Close
                Me.Hide
            End If
        End If
End If
Text1.Text = ""
Text2.Text = ""
Me.Refresh

End Sub


Private Sub Form_Load()
flag = 0
Text2.PasswordChar = "*"

End Sub

