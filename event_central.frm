VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form event_admin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   6480
      TabIndex        =   23
      Top             =   7680
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   3413
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
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      Picture         =   "event_central.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Height          =   975
      Left            =   5520
      TabIndex        =   18
      Text            =   "Bunksheet  Management"
      Top             =   120
      Width           =   9615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8760
      TabIndex        =   14
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3135
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   14415
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   17
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   720
         TabIndex        =   4
         Top             =   1560
         Width           =   12975
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   480
            TabIndex        =   9
            Tag             =   "1"
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IT"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   2640
            TabIndex        =   8
            Tag             =   "2"
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ENTC"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   4680
            TabIndex        =   7
            Tag             =   "3"
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mechanical"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   7080
            TabIndex        =   6
            Tag             =   "4"
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CEES"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   10080
            TabIndex        =   5
            Tag             =   "5"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Label Label7 
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
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Left            =   6960
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   10080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   1
      Top             =   10080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
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
      Left            =   8400
      TabIndex        =   0
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label8 
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
      Left            =   16680
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   18240
      TabIndex        =   21
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
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
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select a Faculty:"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Event:"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   6000
      Width           =   2295
   End
End
Attribute VB_Name = "event_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mode As Integer
Public nm As String
Dim query, update, a As String
Dim con As New ADODB.Connection
Dim cmd_save, cmd, cmd_chk As ADODB.Command
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim row, flag, dpt, eve, response, t, i As Integer
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbGreen
Label8.ForeColor = vbRed
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label4.ForeColor = vbRed
End Sub
Private Sub label8_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label8.ForeColor = vbRed
    
    con.Close
    Me.Hide
    central_admin.Show
End Sub


Private Sub Combo1_LostFocus()
eve = Val(Combo1.ListIndex) + 1
End Sub

Private Sub Command1_Click()
    t = 0
    query = "select id,first_name || ' ' || last_name from faculty where designation < 4"
    If Text1.Text <> "" Then
        query = query + " and first_name like '%" + UCase(Text1.Text) + "%'"
    End If

    If Text2.Text <> "" Then
        query = query + " and last_name like '%" + UCase(Text2.Text) + "%'"
    End If

    If mode = 2 Then
        If dpt = 0 Then
            MsgBox ("Select department")
        Else
            query = query + "and dept = " + Str(dpt)
            t = 1
        End If
    Else
        If dpt <> 0 Then
            a = Str(dpt)
            a = Replace(a, " ", "")
            query = query + " and dept = " + a
        End If
    End If
If mode = 1 Or t = 1 Then

    If mode = 2 Then
        Label2.Visible = True
    End If
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

    If rs.RecordCount > 0 Then
    Set DataGrid1.DataSource = rs
        DataGrid1.Refresh
        DataGrid1.Columns(0).Width = 3000
        DataGrid1.Columns(0).Caption = "Faculty ID"
        DataGrid1.Columns(1).Width = 3000
        DataGrid1.Columns(1).Caption = "Faculty Name"
        DataGrid1.Columns(1).Locked = True
        DataGrid1.Columns(1).Locked = False
    Else
        MsgBox ("No data found")
        Set DataGrid1.DataSource = Nothing
    End If
    If mode = 1 Then
        i = 5
    While i < 10
        Option1(i).Value = False
        i = i + 1
    Wend
    End If
    Text1.Text = ""
    Text2.Text = ""
End If

End Sub

Private Sub Command2_Click()
If DataGrid1.row = -1 Then
    MsgBox ("SELECT FACULTY")
ElseIf mode = 1 And eve = 0 Then
    MsgBox ("SELECT EVENT")
ElseIf mode = 2 And dpt = 0 Then
    MsgBox ("SELECT DEPARTMENT")
Else

    i = 0
    rs.MoveFirst
    While i < row   ' move the cursor to the corresponding row
        rs.MoveNext
        i = i + 1
    Wend
    
    Set cmd_chk = New ADODB.Command
    
    With cmd_chk
        .ActiveConnection = con
        .CommandText = "select event_head,admin from login where faculty_id = '" + rs.Fields(0) + "'"
        .CommandType = adCmdText
    End With
    Set rs1 = New ADODB.Recordset
    With rs1
        .LockType = adLockPessimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open cmd_chk
    End With
    
    If rs1.Fields(1) <> 0 And mode = 1 Then  'checks if the same user is the dept admin
        MsgBox ("Department administrator can't be the event head")
    ElseIf rs1.Fields(0) <> 0 And mode = 2 Then
        MsgBox ("Event head can't be the department administrator")
    Else
    
        With cmd_chk
            .ActiveConnection = con
            .CommandText = "event_admin_search"
            .CommandType = adCmdStoredProc
            .Parameters.Append cmd_chk.CreateParameter("v_name", adVarChar, adParamOutput, 100)
            If mode = 1 Then
                .Parameters.Append cmd_chk.CreateParameter("v_type", adInteger, adParamInput, 2, eve)
                .Parameters.Append cmd_chk.CreateParameter("v_mode", adInteger, adParamInput, 2, mode)
            Else
                .Parameters.Append cmd_chk.CreateParameter("v_type", adInteger, adParamInput, 2, dpt)
                .Parameters.Append cmd_chk.CreateParameter("v_mode", adInteger, adParamInput, 2, mode)
            End If
        End With
        cmd_chk.Execute  'searches if event head or dept head exist
        a = cmd_chk("v_name")  'a = 'xxxxx; means no there is no head
        If a <> "xxxxx" Then
            If mode = 1 Then
                response = MsgBox("You want to replace " + a + " as event head?", vbYesNo + vbQuestion, "EVENT HEAD")
            Else
                response = MsgBox("You want to replace " + a + " as departarment administrator?", vbYesNo + vbQuestion, "DEPT ADMIN")
            End If
        Else
            Set cmd_chk = New ADODB.Command
            With cmd_chk
                    .ActiveConnection = con
                    If mode = 1 Then
                        a = Str(eve)
                        a = Replace(a, " ", "")
                        .CommandText = "update login set event_head = " + a + " where faculty_id = '" + rs.Fields(0) + "'"
                    Else
                        a = Str(dpt)
                        a = Replace(a, " ", "")
                        .CommandText = "update login set admin = " + a + " where faculty_id = '" + rs.Fields(0) + "'"
                    End If
                    .CommandType = adCmdText
            End With
            cmd_chk.Execute
            MsgBox ("saved")
        End If
        
    
        If response = vbYes Or a = "xxxxx" Then
    
            Set cmd_save = New ADODB.Command
    
            With cmd_save
                .ActiveConnection = con
                .CommandText = "event_head_admin"
                .CommandType = adCmdStoredProc
                .Parameters.Append cmd_save.CreateParameter("v_id", adVarChar, adParamInput, 10, rs.Fields(0))
                If mode = 1 Then
                    .Parameters.Append cmd_save.CreateParameter("v_type", adInteger, adParamInput, 2, eve)
                    .Parameters.Append cmd_save.CreateParameter("v_mode", adInteger, adParamInput, 2, mode)
                Else
                    .Parameters.Append cmd_save.CreateParameter("v_type", adInteger, adParamInput, 2, dpt)
                    .Parameters.Append cmd_save.CreateParameter("v_mode", adInteger, adParamInput, 2, mode)
                End If
            End With
            cmd_save.Execute
            MsgBox ("SAVED")
        'End If
    End If
    flag = 0
    dpt = 0
    eve = 0
    Label2.Visible = False

End If
End If
End Sub

Private Sub Command3_Click()
con.Close
Me.Hide
central_admin.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub DataGrid1_Click()
row = DataGrid1.row
End Sub

Private Sub Form_Activate()
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
flag = 0
dpt = 0
eve = 0
row = 0
Label2.Visible = False
query = "select id,first_name || ' ' || last_name from faculty where designation < 4"

Set DataGrid1.DataSource = Nothing
Text1.Text = ""
Text2.Text = ""

If mode = 2 Then    'department admin
    Label5.Visible = False
    Combo1.Visible = False
Else
    Label5.Visible = True
    Combo1.Visible = True
    
End If
Label3.Caption = "Hello " + nm + "!"

End Sub

Private Sub Form_Load()
If mode = 1 Then       'event head
    Combo1.AddItem "AAROHAN"
    Combo1.AddItem "GRACIA"
    Combo1.AddItem "SPANDAN"
    Combo1.AddItem "SPORTS"
    Combo1.AddItem "TESLA"
End If

End Sub



Private Sub Option1_Click(Index As Integer)
dpt = Val(Option1(Index).Tag)
Label2.Caption = Option1(Index).Caption
End Sub

Private Sub Text3_Change()
b
End Sub
