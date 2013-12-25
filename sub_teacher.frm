VERSION 5.00
Begin VB.Form sub_teacher 
   BackColor       =   &H80000005&
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   480
      Picture         =   "sub_teacher.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1230
      TabIndex        =   31
      Top             =   0
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   12600
      TabIndex        =   28
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   7
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   8760
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   6
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   8040
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   5
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7320
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   4
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6600
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   3
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5880
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   2
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5160
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   1
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton Add 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   7
      Top             =   9840
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   0
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   17280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16680
      TabIndex        =   33
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   17880
      TabIndex        =   32
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7080
      TabIndex        =   27
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   26
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   25
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   24
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   23
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   22
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   21
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   20
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Label15"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Division 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Division:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Department:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Subject Teacher"
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
      Left            =   10920
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Subject Name"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Class Teacher:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   15000
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Subject Details"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "sub_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim Sem_st As Integer
Dim query As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim cmd1 As New ADODB.Command
Dim rs1 As New Recordset
Dim i, flag As Integer
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbGreen
Label7.ForeColor = vbGreen
End Sub

Private Sub label6_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label6.ForeColor = vbRed
End Sub

Private Sub label7_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label7.ForeColor = vbRed
End Sub

Private Sub add_Click()   'saves the data


i = 0
If flag = 0 Then    'insert data
 For i = 0 To 7
    rs1.MoveFirst
    While rs1.Fields(0).Value <> Combo2(i).Text
        rs1.MoveNext
    Wend
  With rs2
    .AddNew
    .Fields(0) = UCase(Label16) + "-" + UCase(Label17)
    .Fields(1) = rs1.Fields(1)
    .Fields(2) = Label5(i).Caption
 End With
 Next i

Else   'edit data
    rs2.MoveFirst
    For i = 0 To 7
        rs1.MoveFirst
        While rs1.Fields(0).Value <> Combo2(i).Text
            rs1.MoveNext
        Wend  ' this gets the faculty id against the name
    rs2.Fields(1).Value = rs1.Fields(1).Value     'insert the name only as the other details are already entered
    rs2.MoveNext
    Next i
End If
rs2.MoveFirst
rs2.update  'edits or adds new data
Set DataGrid1.DataSource = rs2

rs1.MoveFirst
    While rs1.Fields(0).Value <> Combo1.Text
        rs1.MoveNext
Wend
Set cmd = New ADODB.Command
With cmd
.ActiveConnection = con
.CommandText = "update login set class_teacher = '" + UCase(Label16) + "-" + UCase(Label17) + "' where faculty_id = '" + rs1.Fields(1) + "'"
End With
cmd.Execute
End Sub

Private Sub Form_Load()
'Combo1.Text = ""
'For i = 0 To 7
'Combo2(i).Text = ""
'Next i

flag = 0
Label15.Caption = dept_admins.dept_da
Label16.Caption = dept_admins.yr_da
Label17.Caption = dept_admins.div_da

con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"

With cmd1
  .ActiveConnection = con
 .CommandText = "select first_name||' '||last_name,id from faculty where dept = " + Str(dept_admins.dept_int)
 .CommandType = adCmdText
End With

With rs1
.CursorLocation = adUseClient
.CursorType = adOpenStaticval
.LockType = adLockPessimistic
.Open cmd1
End With

i = 0

'With rs1

Do While i < rs1.RecordCount   'adding the names of faculty to combo box
Combo1.AddItem rs1.Fields(0).Value
For j = 0 To 7
Combo2(j).AddItem rs1.Fields(0).Value
Next j
rs1.MoveNext
i = i + 1
Loop


'End With

End Sub

Private Sub Option1_Click(Index As Integer)

If Option1(0).Value = True Then
    Sem_st = Val(Option1(0).Caption)
ElseIf Option1(1).Value Then
    Sem_st = Val(Option1(1).Caption)
End If

Set cmd = New ADODB.Command

With cmd
.ActiveConnection = con
.CommandText = "select abbr from sem_subjects where dept = " + Str(dept_admins.dept_int) + " and class='" + dept_admins.yr_da + "' and sem = " + Str(Sem_st)
.CommandType = adCmdText
End With

With rs  'subjects
.CursorLocation = adUseClient
.CursorType = adOpenStaticval
.LockType = adLockPessimistic
.Open cmd
End With


Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = con
    .CommandText = "select * from subject where class='" + UCase(Label16) + "-" + UCase(Label17) + "'"
    .CommandType = adCmdText
End With
With rs2
    .LockType = adLockPessimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .Open cmd
End With  'rs2 has the subject table with the rows

If rs2.RecordCount > 0 Then   ' it means the subject faculty relation was entered before. open now to edit
    flag = 1   ' this indicates edit mode
  'rs2 may or may not contain subjects in the same orders as rs
    rs2.MoveFirst
    
    For i = 0 To 7  'hence we assign the labels according to rs2
        rs1.MoveFirst
        Do While rs1.Fields(1) <> rs2.Fields(1)  ' this gets the name of the faculty against the id
            rs1.MoveNext
        Loop
        Label5(i).Caption = rs2.Fields(2).Value
        Combo2(i).Text = rs1.Fields(0).Value
        rs2.MoveNext
    Next i
Else  'insert new data
    With rs 'assigning the labels when there is no data
    i = 0
    Do While i < rs.RecordCount   'this assins the value of the labels
        Label5(i).Caption = .Fields(0).Value
        rs.MoveNext
        i = i + 1
    Loop
    End With
End If
End Sub
