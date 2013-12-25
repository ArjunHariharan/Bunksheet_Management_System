VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form faculty_home 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   2280
      TabIndex        =   21
      Top             =   5880
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   6800
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "faculty_home.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1230
      TabIndex        =   17
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   10080
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   20415
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   5520
         TabIndex        =   20
         Top             =   840
         Width           =   4215
         _Version        =   524288
         _ExtentX        =   7435
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2012
         Month           =   10
         Day             =   9
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Attendance"
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
         Left            =   10440
         TabIndex        =   15
         Top             =   3480
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   12600
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   17040
         TabIndex        =   9
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton search 
         Caption         =   "Bunksheet"
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
         Left            =   5400
         TabIndex        =   6
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   17520
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12840
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lecture Time:"
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
         Left            =   10800
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Roll No.:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class:"
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
         Height          =   375
         Left            =   15360
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Date:"
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
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last Name:"
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
         Left            =   15480
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "First Name:"
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
         Left            =   10560
         TabIndex        =   1
         Top             =   1920
         Width           =   1935
      End
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   18960
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   17640
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MITCOE Bunksheet Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   13815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
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
      Left            =   7320
      TabIndex        =   11
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "faculty_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag1, flag, response, i, r, c As Integer
Dim select_date As Date
Dim bn_roll, clss, a, lec, typ As String
Dim subquery, query, event_name, faculty_name, dt As String
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim rs_event As ADODB.Recordset
Public fc_dpt, fc_type As Integer
Public fc_name, fc_id As String
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = vbGreen
Label10.ForeColor = vbGreen
End Sub

Private Sub Label10_Click()
End
End Sub

Private Sub LABEL9_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label9.ForeColor = vbRed
End Sub

Private Sub LABEL10_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label10.ForeColor = vbRed
End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    End Sub
Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    End Sub
Private Sub SEARCH_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    End Sub


Private Sub Add()
    If Combo2.Text = "" Then
        MSFlexGrid1.Cols = 12
    Else
        MSFlexGrid1.Cols = 5
    End If
    MSFlexGrid1.Rows = rs_event.RecordCount
    
End Sub

Private Sub Calendar1_Click()
flag1 = 1
select_date = Str(Calendar1.Day) + "-" + Str(Calendar1.Month) + "-" + Str(Calendar1.Year)
Select Case Calendar1.Month
Case 1
    dt = Str(Calendar1.Day) + "-jan-" + Str(Calendar1.Year)
Case 2
    dt = Str(Calendar1.Day) + "-feb-" + Str(Calendar1.Year)
Case 3
    dt = Str(Calendar1.Day) + "-mar-" + Str(Calendar1.Year)
Case 4
    dt = Str(Calendar1.Day) + "-apr-" + Str(Calendar1.Year)
Case 5
    dt = Str(Calendar1.Day) + "-may-" + Str(Calendar1.Year)
Case 6
    dt = Str(Calendar1.Day) + "-jun-" + Str(Calendar1.Year)
Case 7
    dt = Str(Calendar1.Day) + "-jul-" + Str(Calendar1.Year)
Case 8
    dt = Str(Calendar1.Day) + "-aug-" + Str(Calendar1.Year)
Case 9
    dt = Str(Calendar1.Day) + "-sep-" + Str(Calendar1.Year)
Case 10
    dt = Str(Calendar1.Day) + "-oct-" + Str(Calendar1.Year)
Case 11
    dt = Str(Calendar1.Day) + "-nov-" + Str(Calendar1.Year)
Case 12
    dt = Str(Calendar1.Day) + "-dec-" + Str(Calendar1.Year)
End Select
dt = Replace(dt, " ", "")
End Sub

Private Sub Combo1_LostFocus()
    clss = Combo1.Text
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
If fc_type = 0 Then
    con.Close
    password.Show
ElseIf fc_type = 1 Then
    faculty_event_admin.mode = 1
    faculty_event_admin.Show
ElseIf fc_type = 2 Then
    faculty_event_admin.mode = 2
    faculty_event_admin.Show
End If
End Sub

Private Sub Command3_Click()  'attendance
typ = "attendance"

If Combo2.Text <> "" Then
    lecture
Else
    query = "select student.roll,student.first_name|| ' ' || student.last_name,attendance.first,attendance.second,attendance.third,attendance.fourth,attendance.fifth,attendance.sixth,attendance.seventh,attendance.eighth from student, attendance where student.roll = attendance.roll "
End If
'MsgBox (query)
flag = 0

generate_query

If flag = 1 Or Text1.Text <> "" Then   'flag = 1 indicates student detail entered, text1 indicates roll no,

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = query
        .CommandType = adCmdText
    End With
    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        .Open cmd
    End With
    flag1 = 0
    If rs.RecordCount > 0 Then
        attendance
    Else
        MsgBox ("No data found")
        MSFlexGrid1.Clear
    End If
    'attendance
Else
    MsgBox ("Insufficient data")
End If

'Check1.Value = Unchecked
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Form_Activate()
Label6.Caption = "Hello " + fc_name + "!"
password.Hide
If fc_dpt = 5 Then
    Combo1.AddItem ("FE-A")
    Combo1.AddItem ("FE-B")
    Combo1.AddItem ("FE-C")
    Combo1.AddItem ("FE-D")
    Combo1.AddItem ("FE-E")
    Combo1.AddItem ("FE-F")
    Combo1.AddItem ("FE-G")
    Combo1.AddItem ("FE-H")
    Combo1.AddItem ("FE-I")
Else
    Combo1.AddItem ("SE-A")
    Combo1.AddItem ("SE-B")
    Combo1.AddItem ("TE-A")
    Combo1.AddItem ("TE-B")
    Combo1.AddItem ("BE-A")
    Combo1.AddItem ("BE-B")
End If

Combo2.AddItem ("8:30-9:30")
Combo2.AddItem ("9:30-10:30")
Combo2.AddItem ("10:45-11:45")
Combo2.AddItem ("11:45-12:45")
Combo2.AddItem ("1:30-2:30")
Combo2.AddItem ("2:30-3:30")
Combo2.AddItem ("3:30-4:30")
Combo2.AddItem ("4:30-5:30")


a = Str(fc_dpt)
a = Replace(a, " ", "")
select_date = Str(Calendar1.Month) + Str(Calendar1.Day) + Str(Calendar1.Year)
flag = 99
'eid = 5   'temporary
flag1 = 0
typ = ""
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"

Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = "select login.faculty_id, faculty.first_name|| ' ' || faculty.last_name from faculty,login where login.event_head <> 0"
        .CommandType = adCmdText
    End With
    
Set rs_event = New ADODB.Recordset  'faculty_id,faculty_name,
With rs_event
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockPessimistic
    .Open cmd
End With

If fc_type = 0 Then
    Command2.Caption = "Change Password"
ElseIf fc_type = 1 Or fc_type = 2 Then
    Command2.Caption = "Back"
End If
End Sub

Private Sub Option1_Click(Index As Integer)
dpt = Option1(Index).Tag
End Sub



Private Sub search_Click()
typ = "bunksheet"
'subquery = "select roll from student where dept = " + a
If Combo2.Text <> "" Then  'lecture selected
    lecture  ' gets the lecture no and generates the query
Else
    query = "select student.roll,student.first_name|| ' ' || student.last_name,bunksheet.event,bunksheet.faculty_incharge,bunksheet.first,bunksheet.second,bunksheet.third,bunksheet.fourth,bunksheet.fifth,bunksheet.sixth,bunksheet.seventh,bunksheet.eighth from student, bunksheet where student.roll = bunksheet.roll"
End If

flag = 0

generate_query


If flag = 1 Or Text1.Text <> "" Then ' Or Combo2.Text <> "" Then

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = query
        .CommandType = adCmdText
    End With
    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        .Open cmd
    End With
    flag1 = 0
    If rs.RecordCount > 0 Then
        bunksheet
    Else
        MsgBox ("No data found")
        MSFlexGrid1.Clear
    End If
Else
    MsgBox ("Insufficient data")
End If


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub Text1_LostFocus()
'call a stored procedure to find if the roll exists in the same department
If Len(Text1.Text) > 0 Then
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = "fac_search_roll_chk"
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_roll", adVarChar, adParamInput, 10, UCase(Text1.Text))
        .Parameters.Append cmd.CreateParameter("v_dpt", adInteger, adParamInput, 2, fc_dpt)
        .Parameters.Append cmd.CreateParameter("v_flag", adInteger, adParamOutput, 2)
    End With
    cmd.Execute
    i = cmd("v_flag")
    If i = 1 Then
        MsgBox ("Roll no not accessible")
        Text1.Text = ""
    ElseIf i = 2 Then
        MsgBox ("INVALID ROLL NO")
        Text1.Text = ""
    End If
End If
End Sub

Private Sub lecture()
Select Case Combo2.ListIndex
    Case 0
        lec = "first"
    Case 1
        lec = "second"
    Case 2
        lec = "third"
    Case 3
        lec = "fourth"
    Case 4
        lec = "fifth"
    Case 5
        lec = "sixth"
    Case 6
        lec = "seventh"
    Case 7
        lec = "eighth"
End Select

a = Str(fc_dpt)
a = Replace(a, " ", "")

If typ = "bunksheet" Then
    query = "Select student.roll, student.first_name|| ' ' || student.last_name,bunksheet.event,bunksheet.faculty_incharge,bunksheet." + lec + " from student, bunksheet where student.roll = bunksheet.roll"
Else
    query = "Select student.roll, student.first_name|| ' ' || student.last_name,attendance." + lec + " from student,attendance where student.roll = attendance.roll"
End If
End Sub

Private Sub generate_query()

a = Str(select_date)
a = Replace(a, " ", "")

subquery = "select roll from student where dept = " + Str(fc_dpt)


If Text1.Text <> "" Then   'search by roll

    query = query + " and student.roll = '" + UCase(Text1.Text) + "'"
    
    If flag1 = 1 Then  'date selected
        query = query + " and to_char(" + typ + ".day,'fmdd-mon-yyyy')='" + dt + "'"
    Else  'ask for date

        response = MsgBox("Do you want to search for current date?", vbYesNo + vbQuestion, "SEARCH BY DAY")
    If response = vbYes Then
        query = query + " and to_char(" + typ + ".day,'fmdd-mon-yyyy')='" + dt + "'"
        flag1 = 1
    End If
    
End If
    
Else

    If flag1 = 1 Then  'date selected
        query = query + " and to_char(" + typ + ".day,'fmdd-mon-yyyy')='" + dt + "'"
    Else  'ask for date
        response = MsgBox("Do you want to search for current date?", vbYesNo + vbQuestion, "SEARCH BY DAY")
    If response = vbYes Then
        query = query + " and to_char(" + typ + ".day,'fmdd-mon-yyyy')='" + dt + "'"
        flag1 = 1
    End If
    
    End If

    If clss <> "" Then  'class
        subquery = subquery + " and class = '" + clss + "'"
        flag = 1
    End If

    If Text2.Text <> "" Then   'first name
        subquery = subquery + " and first_name = '" + UCase(Text2.Text) + "'"
        flag = 1
    End If

    If Text3.Text <> "" Then  'last name
        subquery = subquery + " and last_name = '" + UCase(Text3.Text) + "'"
        flag = 1
    End If
End If

If flag = 1 Then   'subquery exists
    query = query + " and student.roll in (" + subquery + ")"
End If
If Text1.Text = "" And flag1 = 1 And flag = 0 Then  'only date selected
a = Str(fc_dpt)
a = Replace(a, " ", a)
    query = query + " and " + typ + ".roll = student.roll in(select * from student where dept = " + a
End If
'MsgBox (query)
End Sub

Public Sub bunksheet()
If Combo2.Text = "" Then
        MSFlexGrid1.Cols = 13
        MSFlexGrid1.Rows = rs.RecordCount + 1
        With MSFlexGrid1
            .row = 0
            .col = 0
            .Text = "S.No"
            .col = 1
            .Text = "Roll no"
            .col = 2
            .ColWidth(2) = 2000
            .Text = "Name"
            .col = 3
            .Text = "Event"
            .col = 4
            .ColWidth(4) = 3000
            .Text = "Faculty Incharge"
            .col = 5
            .Text = "8:30-9:30"
            .col = 6
            .Text = "9:30-10:30"
            .col = 7
            .Text = "10:45-11:45"
            .col = 8
            .Text = "11:45-12:45"
            .col = 9
            .Text = "1:30-2:30"
            .col = 10
            .Text = "2:30-3:30"
            .col = 11
            .Text = "3:30-4:30"
            .col = 12
            .Text = "4:30-5:30"
        End With
        
        r = 1
        c = 0
        rs.MoveFirst
        Do While r <= rs.RecordCount
            MSFlexGrid1.row = r
            MSFlexGrid1.col = c
            MSFlexGrid1.Text = r
            c = 1
            Do While c < 3
                With MSFlexGrid1
                    .row = r
                    .col = c
                    .Text = rs.Fields(c - 1)
                End With
                c = c + 1
            Loop
            name_eve_fac
            With MSFlexGrid1
                    .row = r
                    .col = 3  'faculty incharge
                    .Text = event_name
                    .col = 4  'event
                    .Text = faculty_name
            End With
            c = 5
            Do While c < 13
                With MSFlexGrid1
                    .row = r
                    .col = c
                    .Text = rs.Fields(c - 1)
                    If rs.Fields(c - 1).Value = "Y" Then
                        .CellForeColor = vbRed
                        .CellFontBold = True
                    End If
                End With
                c = c + 1
            Loop
        r = r + 1
        rs.MoveNext
        c = 0
        Loop
Else  ' lecture selected
        MSFlexGrid1.Cols = 6
        MSFlexGrid1.Rows = rs.RecordCount + 1
        With MSFlexGrid1
            .row = 0
            .col = 0
            .Text = "S.No"
            .col = 1
            .Text = "Roll no"
            .col = 2
            .ColWidth(2) = 2000
            .Text = "Name"
            .col = 3
            .Text = "Event"
            .col = 4
            .ColWidth(4) = 3000
            .Text = "Faculty Incharge"
            .col = 5
            .Text = Combo2.Text
        End With
        
        r = 1
        c = 0
        rs.MoveFirst
        Do While r <= rs.RecordCount
            name_eve_fac
            With MSFlexGrid1
                    .row = r
                    .col = 0
                    .Text = r
                    .col = 1
                    .Text = rs.Fields(0)
                    .col = 2
                    .Text = rs.Fields(1)
                    .col = 3  'faculty incharge
                    .Text = event_name
                    .col = 4  'event
                    .Text = faculty_name
                    .col = 5
                    .Text = rs.Fields(4).Value
                    If rs.Fields(4).Value = "Y" Then
                        .CellForeColor = vbRed
                        .CellFontBold = True
                    End If
            End With
        r = r + 1
        rs.MoveNext
        Loop
        
    End If

End Sub

Private Sub name_eve_fac()
rs_event.MoveFirst
i = 0
Do While Not rs_event.EOF
    If rs_event.Fields(0).Value = rs.Fields(3).Value Then
        faculty_name = rs_event.Fields(1).Value
        Exit Do
    End If
    rs_event.MoveNext
Loop

Select Case rs.Fields(2)
    Case 1
        event_name = "Aarohan"
    Case 2
        event_name = "Gracia"
    Case 3
        event_name = "Spandan"
    Case 4
        event_name = "Sports"
    Case 5
        event_name = "Tesla"
End Select

End Sub

Public Sub attendance()
If Combo2.Text = "" Then
        MSFlexGrid1.Cols = 11
        MSFlexGrid1.Rows = rs.RecordCount + 1
        With MSFlexGrid1
            .row = 0
            .col = 0
            .Text = "S.No"
            .col = 1
            .Text = "Roll no"
            .col = 2
            .ColWidth(1) = 2000
            .Text = "Name"
            .col = 3
            .ColWidth(2) = 3000
            .ColWidth(3) = 2000
            .ColWidth(4) = 2000
            .ColWidth(5) = 2000
            .ColWidth(6) = 2000
            .ColWidth(7) = 2000
            .ColWidth(8) = 2000
            .ColWidth(9) = 2000
            .ColWidth(10) = 2000
            .Text = "8:30-9:30"
            .col = 4
            .Text = "9:30-10:30"
            .col = 5
            .Text = "10:45-11:45"
            .col = 6
            .Text = "11:45-12:45"
            .col = 7
            .Text = "1:30-2:30"
            .col = 8
            .Text = "2:30-3:30"
            .col = 9
            .Text = "3:30-4:30"
            .col = 10
            .Text = "4:30-5:30"
        End With
        
        r = 1
        c = 0
        rs.MoveFirst
        Do While r <= rs.RecordCount
            MSFlexGrid1.row = r
            MSFlexGrid1.col = c
            MSFlexGrid1.Text = r
            c = 1
            Do While c < 11
                With MSFlexGrid1
                    .row = r
                    .col = c
                    .Text = rs.Fields(c - 1)
                    If rs.Fields(c - 1).Value = "P*" Then
                        .CellForeColor = vbRed
                        .CellFontBold = True
                    End If
                End With
                c = c + 1
            Loop
        r = r + 1
        rs.MoveNext
        Loop
        
Else  ' lecture selected
        MSFlexGrid1.Cols = 4
        MSFlexGrid1.Rows = rs.RecordCount + 1
        With MSFlexGrid1
            .row = 0
            .col = 0
            .Text = "S.No"
            .col = 1
            .Text = "Roll no"
            .col = 2
            .ColWidth(1) = 2000
            .Text = "Name"
            .col = 3
            .ColWidth(2) = 3000
            .Text = Combo2.Text
            .ColWidth(3) = 2000
        End With
        
        r = 1
        c = 0
        rs.MoveFirst
        Do While r <= rs.RecordCount
            'name_eve_fac
            With MSFlexGrid1
                    .row = r
                    .col = 0
                    .Text = r
                    .col = 1
                    .Text = rs.Fields(0)
                    .col = 2
                    .Text = rs.Fields(1)
                    .col = 3
                    .Text = rs.Fields(2)
                    If rs.Fields(2).Value = "P*" Then
                        .CellForeColor = vbRed
                        .CellFontBold = True
                    End If
            End With
        r = r + 1
        rs.MoveNext
        c = 0
        Loop
        
    End If

End Sub
