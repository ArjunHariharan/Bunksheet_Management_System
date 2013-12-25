VERSION 5.00
Begin VB.Form password 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PASSWORD "
   ClientHeight    =   8805
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "password.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1230
      TabIndex        =   9
      Top             =   120
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password  Management"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim flag As Integer
Public pass_change_id As String
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

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    
    MsgBox ("Fields should not be empty")
ElseIf Text2.Text = Text3.Text Then
    con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = "password_change"   ' calls the stored procedure to verify the username and password
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_id", adVarChar, adParamInput, 30, pass_change_id)
        .Parameters.Append cmd.CreateParameter("v_old_pass", adVarChar, adParamInput, 30, Text1.Text)
        .Parameters.Append cmd.CreateParameter("v_new_pass", adVarChar, adParamInput, 10, Text2.Text)
        .Parameters.Append cmd.CreateParameter("V_flag", adInteger, adParamOutput, 2)
    End With
        cmd.Execute
        flag = cmd("v_flag")
        If flag = 0 Then
            MsgBox ("Password incorrect")
        Else
            MsgBox ("Password changed")
            con.Close
            Me.Hide
        End If
Else
    MsgBox ("Entered password don't match")
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Text1.PasswordChar = "*"
Text2.PasswordChar = "*"
Text3.PasswordChar = "*"
End Sub

