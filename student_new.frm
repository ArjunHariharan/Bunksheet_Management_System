VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form student_new 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   2400
      TabIndex        =   21
      Top             =   5880
      Width           =   15100
      _ExtentX        =   26644
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back "
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
      Left            =   10680
      TabIndex        =   18
      Top             =   9840
      Width           =   2295
   End
   Begin VB.CommandButton delete 
      Caption         =   "Delete"
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
      Left            =   4440
      TabIndex        =   14
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Student Database"
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
      Left            =   13680
      TabIndex        =   13
      Top             =   9840
      Width           =   3495
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
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
      Left            =   9480
      TabIndex        =   12
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton update 
      Caption         =   "Save"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   9840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   9120
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   12240
      TabIndex        =   8
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      DataField       =   "t"
      Height          =   495
      Left            =   6360
      MaxLength       =   11
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   14760
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "student_new.frx":0000
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   16080
      TabIndex        =   23
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   18240
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Height          =   495
      Left            =   12240
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Left            =   4680
      TabIndex        =   19
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   8160
      TabIndex        =   17
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
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
      Left            =   14160
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Email ID:"
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
      Left            =   10320
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Mobile No:"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Middle Name:"
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
      Left            =   6840
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
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
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Student Database"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "student_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As New ADODB.Recordset
Dim i, response, flag, exist, rll, upd As Integer
Dim a, temp As String
Dim old As String
Public roll_start As String
Public branch_stu As Integer '1,2,3,4,5 corresponding to each branch
Public class_stu, name_admin, query As String 'stores TE-B
Public Division_stu, dept As String ' stores a/b/c/d.... division

Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub ADD_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub delete_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub update_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub label12_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label12.ForeColor = vbRed
End Sub

Private Sub label13_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label13.ForeColor = vbRed
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbGreen
Label13.ForeColor = vbGreen
End Sub

Private Sub add_Click()
If Text1.Text = "" Then
    MsgBox ("Please fill first name!")
ElseIf Text3.Text = "" Then
    MsgBox ("Please fill last name!")
ElseIf Text4.Text = "" Or Len(Text4.Text) <> 10 Or Not IsNumeric(Text4.Text) Then
    MsgBox ("INVALID MOBILE NO!")
ElseIf Text5.Text = "" Then
    MsgBox ("Please fill email ID!")
Else
    With rs
    .AddNew
    If exist = 0 Then
        .Fields(0) = i  'temp roll
    Else  ' if exist is 1 then roll no should be added to end
        Set cmd = New ADODB.Command
        With cmd
            .ActiveConnection = con
            .CommandText = "stu_late"
            .CommandType = adCmdStoredProc
            .Parameters.Append cmd.CreateParameter("v_new", adNumeric, adParamOutput, 5)
            .Parameters.Append cmd.CreateParameter("v_class", adVarChar, adParamInput, 5, class_stu + "-" + Division_stu)
        End With
        cmd.Execute
        rll = cmd("v_new")
        temp = roll_start
            If rll < 9 Then
            temp = temp + "0"
            temp = temp + Str(rll)
            Else
                temp = temp + Str(rll)
            End If
        temp = Replace(temp, " ", "")
        .Fields(0) = UCase(temp)  'roll
    End If
    .Fields(1) = UCase(Text1.Text)  'first name
    .Fields(2) = UCase(Text2.Text) 'middle name
    .Fields(3) = UCase(Text3.Text) 'last name
    .Fields(4) = Text4.Text  ' mobile no
    .Fields(5) = class_stu + "-" + Division_stu   'class
    .Fields(6) = branch_stu   'branch
    .Fields(7) = LCase(Text5.Text)  'email id
    upd = 1  'indicates that roll no need to be revised
    End With
    If exist <> 0 Then
        rs.update
    End If
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    i = i + 1
    MsgBox ("Added")
    flag = 1
End If
End Sub

Private Sub Command1_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        update_Click
    End If
End If
student_view.query = query
student_view.Label4.Caption = "Hello " + name_admin + "!"
student_view.Label2.Caption = Me.Label5.Caption
student_view.Label3.Caption = Me.Label6.Caption
con.Close
student_view.Show
Me.Hide
End Sub




Private Sub Command3_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        update_Click
    End If
End If
con.Close
Me.Hide
class_division.Show
End Sub



Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
Case 1
    DataGrid1.Columns(ColIndex).Value = UCase$(DataGrid1.Columns(ColIndex).Value)
Case 2
    If DataGrid1.Columns(ColIndex).Text <> "" Then
        DataGrid1.Columns(ColIndex).Value = UCase$(DataGrid1.Columns(ColIndex).Value)
    Else
        DataGrid1.Columns(ColIndex).Value = ""
    End If
Case 3
    If exist = 0 Then 'indicates roll can be sorted
        DataGrid1.Columns(ColIndex).Value = UCase$(DataGrid1.Columns(ColIndex).Value)
        upd = 1 ' sort roll
    Else
        DataGrid1.Columns(ColIndex).Value = UCase$(old)
        MsgBox ("Record can't be updated now")
        upd = 0  'dont sort
    End If
Case 4
    If Len(DataGrid1.Columns(ColIndex).Value) <> 10 Or Not IsNumeric(DataGrid1.Columns(ColIndex).Value) Then
        MsgBox ("Invalid mobile no!")
        DataGrid1.Columns(ColIndex).Value = old
    End If
Case 6
    DataGrid1.Columns(ColIndex).Value = LCase$(DataGrid1.Columns(ColIndex).Value)
End Select

End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
flag = 1
upd = 0 ' dont't sort roll
old = OldValue
End Sub

Private Sub delete_Click()
If DataGrid1.ApproxCount = 0 Then
    MsgBox ("No data present!!")
Else
    rs.delete (adAffectCurrent)
    update_Click
    DataGrid1.Refresh
    flag = 0
End If
End Sub

Private Sub Form_Activate()

flag = 0
Label8.Caption = "Hello " + UCase(name_admin) + "!"
Label6 = class_stu + "-" + Division_stu
Select Case branch_stu
Case 1:
    Label5 = " COMPUTER"
Case 2:
    Label5 = " IT"
Case 3:
    Label5 = " ENTC"
Case 4:
    Label5 = " MECHANICAL"
Case 5:
    Label5 = " CEES"
End Select

con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
i = 1
Set cmd = New ADODB.Command
With cmd
.ActiveConnection = con
.CommandText = "select * from student where class = '" + class_stu + "-" + Division_stu + "'and dept = " + Str(branch_stu) + " order by roll"
.CommandType = adCmdText
End With
query = cmd.CommandText
With rs
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockPessimistic
.Open cmd
End With

Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(5).Visible = False
DataGrid1.Columns(6).Visible = False
Do While i < 8
    DataGrid1.Columns(i).Width = 3000
    DataGrid1.Columns(i).Alignment = dbgCenter
    i = i + 1
Loop

Set cmd = New ADODB.Command
With cmd
.ActiveConnection = con
.CommandText = "exist_student"
.CommandType = adCmdStoredProc
.Parameters.Append cmd.CreateParameter("v_class", adVarChar, adParamInput, 5, class_stu + "-" + Division_stu)
.Parameters.Append cmd.CreateParameter("v_cnt", adNumeric, adParamOutput, 5)
End With
cmd.Execute
exist = cmd("v_cnt")   ' if exist = 1 then means that new record should be added to last
upd = 0   ' dont sort roll
End Sub



Private Sub Label12_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        update_Click
    End If
End If
End
End Sub

Private Sub Label13_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        update_Click
    End If
End If
con.Close
Me.Hide
central_admin.Show
End Sub

Private Sub update_Click()
flag = 0
If exist = 0 And upd = 1 Then
    If DataGrid1.ApproxCount > 0 Then
        rs.update
    End If
    
    If upd = 1 Then   ' means new data added or last name updated,sort data
        rs.Sort = "last_name asc"
        If DataGrid1.ApproxCount = 0 Then
            MsgBox ("Data empty!")
        Else
            Dim temp As String
            i = 1
            rs.MoveFirst
            Do While rs.EOF <> True
                temp = roll_start
                If i < 9 Then
                temp = temp + "0"
                temp = temp + Str(i)
                Else
                    temp = temp + Str(i)
                End If
                rs.MoveNext
                temp = Replace(temp, " ", "")
                If rs.EOF <> True Then
                    rs.Fields(0).Value = Str(9999 - i)
                    rs.MovePrevious
                    rs.Fields(0).Value = temp
                Else
                    rs.MovePrevious
                    rs.Fields(0).Value = temp
                End If
                i = i + 1
                rs.Move (1)
            Loop
            rs.MoveFirst
            rs.update
            i = 1
            DataGrid1.Refresh
        End If
    End If
Else  ' new record added to end or column changed
    rs.update
End If
 MsgBox ("Saved")
End Sub
