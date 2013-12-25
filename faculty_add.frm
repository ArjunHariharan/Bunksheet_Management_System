VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form faculty_add 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   3480
      TabIndex        =   28
      Top             =   6600
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3836
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
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "faculty_add.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4680
      TabIndex        =   23
      Top             =   10200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4440
      TabIndex        =   22
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   16440
      TabIndex        =   21
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10800
      MaxLength       =   10
      TabIndex        =   20
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6000
      TabIndex        =   19
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16560
      TabIndex        =   18
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   17
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save changes"
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
      Left            =   13320
      TabIndex        =   15
      Top             =   10200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9000
      TabIndex        =   14
      Top             =   10200
      Width           =   2415
   End
   Begin VB.CommandButton add 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   6
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   5280
      TabIndex        =   4
      Top             =   4080
      Width           =   10095
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Professor"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   11
         Tag             =   "2"
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Asst. Professor"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   10
         Tag             =   "3"
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lab Asst."
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   9
         Tag             =   "4"
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clerk"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   8
         Tag             =   "5"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HOD"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Tag             =   "1"
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   18480
      TabIndex        =   27
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   17040
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   8520
      TabIndex        =   24
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   7440
      TabIndex        =   16
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   14400
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Middle Name:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   14040
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Faculty Database"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Faculty ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone No.:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "faculty_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim cmd_del As ADODB.Command
Dim cmd_chk As ADODB.Command
Dim rs As New ADODB.Recordset
Dim op, i, temp, flg, response, flag As Integer
Dim desig, old As String
Public fac_dept As Integer
Public f_add_name As String
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbGreen
Label11.ForeColor = vbGreen
End Sub

Private Sub Label10_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        Command4_Click
    End If
End If
con.Close
Me.Hide
central_admin.Show

End Sub

Private Sub LABEL10_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label10.ForeColor = vbRed
End Sub

Private Sub Label11_Click()
If flag = 1 Then

    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        Command4_Click
    End If
End If
End

End Sub

Private Sub LABEL11_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label11.ForeColor = vbRed
End Sub

Private Sub ADD_MouseMove(Button As Integer, Shift As Integer, X As Single, _
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
Private Sub command4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 End Sub

Private Sub check()
Set cmd_chk = New ADODB.Command
    With cmd_chk
        .ActiveConnection = con
        .CommandText = "search_id"
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd_chk.CreateParameter("v_id", adVarChar, adParamInput, 30, old)
        .Parameters.Append cmd_chk.CreateParameter("v_flag", adNumeric, adParamOutput, 30)
    End With
cmd_chk.Execute
flg = cmd_chk("v_flag")  'returns 1 if there exists a faculty with the same id
End Sub

Private Sub add_Click()
If Text1.Text = "" Then
    MsgBox ("Please fill faculty ID!")
ElseIf Text2.Text = "" Then
    MsgBox ("Please fill first name!")
ElseIf Text4.Text = "" Then
    MsgBox ("Please fill last name!")
ElseIf Len(Text5.Text) <> 10 Or Not IsNumeric(Text5.Text) Then
    MsgBox ("mobile number invalid!")
ElseIf Text6.Text = "" Then
    MsgBox ("Please fill email ID!")
Else
    old = UCase(Text1.Text)
    check
    If flg <> 1 Then
        With rs
            .AddNew
            .Fields(0) = UCase(Text1.Text)  'faculty id
            .Fields(1) = UCase(Text2.Text)   'first name
            .Fields(2) = UCase(Text3.Text)   'middle name
            .Fields(3) = UCase(Text4.Text)   'last name
            .Fields(4) = Text5.Text   'phone
            .Fields(5) = temp   'designation
            .Fields(6) = fac_dept  'faculty department
            .Fields(7) = Text6.Text  'email
        End With
        rs.update
        flag = 1

        MsgBox ("ADDED")
    'write a code to display the designation in datagrid(8)
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""

        Option2(0).Value = False
        Option2(1).Value = False
        Option2(2).Value = False
        Option2(3).Value = False
        Option2(4).Value = False
    Else
        MsgBox ("Faculty id exists!!!")
    End If
End If
DataGrid1.Refresh
End Sub

Private Sub Command1_Click()

If DataGrid1.ApproxCount = 0 Then
 MsgBox ("No data present!!")
Else
    i = 0
    rs.delete (adAffectCurrent)
    rs.update
    DataGrid1.Refresh
    MsgBox ("Deleted")
    flag = 0
End If
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
If flag = 1 Then
    response = MsgBox("Do you want to save?", vbYesNo, "SAVE")
    If response = vbYes Then
        Command4_Click
    End If
End If
faculty_dept.n = f_add_name
central_admin.c_a_n = f_add_name
Me.Refresh
con.Close
central_admin.Show
faculty_dept.Show
Me.Hide
End Sub

Private Sub Command4_Click()
If DataGrid1.ApproxCount = 0 Then
 MsgBox ("No data present!!")
Else
 rs.update
 DataGrid1.Refresh
 MsgBox ("Saved")
 flag = 0
End If
End Sub



Private Sub Command5_Click()
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
Case 1
    DataGrid1.Columns(ColIndex).Value = UCase$(DataGrid1.Columns(ColIndex).Value)
Case 3
    DataGrid1.Columns(ColIndex).Value = UCase$(DataGrid1.Columns(ColIndex).Value)
Case 4
    If Len(DataGrid1.Columns(ColIndex).Value) <> 10 Or Not IsNumeric(DataGrid1.Columns(ColIndex).Value) Then
        MsgBox ("Invalid mobile no!")
        DataGrid1.Columns(ColIndex).Value = old
    End If
Case 5
    If Str(DataGrid1.Columns(ColIndex).Value) > 5 Then
        MsgBox ("Invalid designation")
        DataGrid1.Columns(ColIndex).Value = old
    End If
Case 7
    DataGrid1.Columns(ColIndex).Value = LCase$(DataGrid1.Columns(ColIndex).Value)
End Select
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
flag = 1
old = OldValue
End Sub

Private Sub Form_Activate()

con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"

Set cmd = New ADODB.Command
With cmd
.ActiveConnection = con
.CommandText = "select * from faculty where dept = " + Str(fac_dept) + "order by id asc"
.CommandType = adCmdText
End With

With rs
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockPessimistic
.Open cmd
End With

Set DataGrid1.DataSource = rs
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(0).Locked = True

i = 0
Do While i < 8
    DataGrid1.Columns(i).Width = 2000
    DataGrid1.Columns(i).Alignment = dbgGeneral
    i = i + 1
Loop
DataGrid1.Columns(5).Alignment = dbgCenter
i = 1
Label9.Caption = "Hello " + f_add_name + "!"

'con.Close

End Sub

Private Sub Option2_Click(Index As Integer)
temp = Val(Option2(Index).Tag)
desig = Option2(Index).Caption
End Sub

