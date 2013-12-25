VERSION 5.00
Begin VB.Form class_division 
   BackColor       =   &H00FFFFFF&
   Caption         =   "class_division"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   120
      Picture         =   "class_division.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   22
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   21
      Top             =   9480
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DIVISION"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   2040
      TabIndex        =   11
      Top             =   7320
      Width           =   16455
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   14280
         TabIndex        =   19
         Tag             =   "6"
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   12240
         TabIndex        =   18
         Tag             =   "6"
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   10200
         TabIndex        =   17
         Tag             =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   16
         Tag             =   "4"
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   6240
         TabIndex        =   15
         Tag             =   "3"
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   4320
         TabIndex        =   14
         Tag             =   "2"
         Top             =   780
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Tag             =   "1"
         Top             =   750
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Tag             =   "0"
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   16455
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CEES"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   13800
         TabIndex        =   10
         Tag             =   "1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MECHANICAL"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   9960
         TabIndex        =   9
         Tag             =   "7"
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ENTC"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   6960
         TabIndex        =   8
         Tag             =   "5"
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IT"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   7
         Tag             =   "3"
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPUTER"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Tag             =   "1"
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   2040
      TabIndex        =   0
      Top             =   5160
      Width           =   16455
      Begin VB.OptionButton fe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FE"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "F1"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BE"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   2
         Left            =   10080
         TabIndex        =   3
         Tag             =   "B4"
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TE"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   7080
         TabIndex        =   2
         Tag             =   "T3"
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SE"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   4200
         TabIndex        =   1
         Tag             =   "S2"
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   8640
      TabIndex        =   26
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   17520
      TabIndex        =   25
      Top             =   120
      Width           =   1215
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
      Height          =   375
      Left            =   18960
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Database"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   6240
      TabIndex        =   20
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "class_division"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Year As String    'gets the class ie fe/se/te/be
Dim section As Integer    'gets the value for roll no calculation
Dim branch, b As Integer '1,2,3,4,5 corresponding to each branch
Dim class As String  'stores TE-B
Dim Division As String  ' stores a/b/c/d.... division
Public name_div As String
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

Private Sub Label2_Click()
End
End Sub

Private Sub LABEL2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label2.ForeColor = vbRed
End Sub

Private Sub Label3_Click()
central_admin.Show
Me.Hide
End Sub

Private Sub label3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label3.ForeColor = vbRed
    
End Sub



Private Sub Command1_Click()
If branch = 0 Then
    MsgBox ("Please select department!!")
Else
    If Year = "" Then
        MsgBox ("Please select year!!")
    Else
        If Division = "" Then
        MsgBox ("Please select division!!")
        Else
            branch = branch + section
            Year = Year & Str(branch)
            Year = Replace(Year, " ", "")
            student_new.roll_start = Year
            If fe.Value = True Then
                class = fe.Caption
                branch = branch - section
                branch = branch + 8
            Else
                branch = branch - section
            End If
            student_new.name_admin = name_div
            student_new.branch_stu = b
            student_new.class_stu = class
            student_new.Division_stu = Division
            student_new.Show
            Me.Hide
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
central_admin.Show
Me.Hide
End Sub

Private Sub Form_Load()
Division = ""
branch = 0
section = 0
Year = ""
Label4.Caption = "Hello " + UCase(name_div)

i = 0

fe.Enabled = False

i = 0
Do While i < 3
Option1(i).Enabled = False
i = i + 1
Loop

Option4(0).Enabled = False
Option4(1).Enabled = False
Option3.Enabled = True

i = 0
Do While i < 6
Option5(i).Enabled = False
i = i + 1
Loop

i = 0
Do While i < 4
    Option2(i).Enabled = True
    i = i + 1
Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbGreen
Label2.ForeColor = vbGreen
End Sub

Private Sub Option1_Click(Index As Integer)
Year = Option1(Index).Tag
class = Option1(Index).Caption

End Sub

Private Sub Option2_Click(Index As Integer)
Option4(0).Enabled = True
Option4(1).Enabled = True
fe.Enabled = False
fe.Value = False


i = 0
Do While i < 3
Option1(i).Enabled = True
i = i + 1
Loop

i = 0
Do While i < 6
Option5(i).Enabled = False
i = i + 1
Loop

branch = Val(Option2(Index).Tag)
Year = ""
section = 0
b = Index + 1
End Sub

Private Sub Option3_Click()
Option4(0).Enabled = True
Option4(1).Enabled = True
fe.Enabled = True
fe.Value = True

i = 0
Do While i < 3
Option1(i).Enabled = False
i = i + 1
Loop

i = 0
Do While i < 6
Option5(i).Enabled = True
i = i + 1
Loop

branch = Val(Option3.Tag)
Year = fe.Tag
class = "FE"
b = 5
End Sub

Private Sub Option4_Click(Index As Integer)
section = Val(Option4(Index).Tag)
Division = Option4(Index).Caption

End Sub

Private Sub Option5_Click(Index As Integer)
section = Val(Option5(Index).Tag)
Division = Option5(Index).Caption

End Sub

