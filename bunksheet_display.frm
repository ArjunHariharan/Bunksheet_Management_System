VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bunksheet_display 
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   360
      TabIndex        =   24
      Top             =   6480
      Width           =   17000
      _ExtentX        =   29977
      _ExtentY        =   3201
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
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   120
      Picture         =   "bunksheet_display.frx":0000
      ScaleHeight     =   1440
      ScaleWidth      =   1230
      TabIndex        =   21
      Top             =   120
      Width           =   1290
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   630
      Left            =   7200
      TabIndex        =   20
      Text            =   "View  Bunksheet"
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8760
      TabIndex        =   18
      Top             =   10200
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save "
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
      Left            =   13560
      TabIndex        =   17
      Top             =   10080
      Width           =   2415
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
      Left            =   3720
      TabIndex        =   9
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
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   20415
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   5280
         TabIndex        =   25
         Top             =   1080
         Width           =   4575
         _Version        =   524288
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483633
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Department"
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
         Height          =   1335
         Left            =   10680
         TabIndex        =   11
         Top             =   1680
         Width           =   8775
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CEES"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   6960
            TabIndex        =   16
            Tag             =   "5"
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mechanical"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5160
            TabIndex        =   15
            Tag             =   "4"
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ENTC"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3600
            TabIndex        =   14
            Tag             =   "3"
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IT"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   13
            Tag             =   "2"
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   12
            Tag             =   "1"
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show All"
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
         Height          =   735
         Left            =   1680
         TabIndex        =   8
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   6
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   17040
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12720
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   525
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
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
         TabIndex        =   10
         Top             =   360
         Width           =   1935
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
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
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
         Left            =   15120
         TabIndex        =   2
         Top             =   600
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
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   17280
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   18360
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
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
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "bunksheet_display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag1, flag, dpt, response, i As Integer
Dim bn_roll, date_string, a As String
Dim subquery, query, date_input, dt As String
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As New ADODB.Recordset
Public eid As Integer
Public nme As String
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbGreen
Label7.ForeColor = vbGreen
End Sub

Private Sub Label6_Click()
End
End Sub

Private Sub SEARCH_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub label7_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label7.ForeColor = vbRed
End Sub
Private Sub command6_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label6.ForeColor = vbRed
End Sub

Private Sub Calendar1_Click()
flag1 = 1
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

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
con.Close
Me.Hide
bunksheet.Show
End Sub

Private Sub Command3_Click()
rs.update
End Sub

Private Sub Command4_Click()
rs.delete (adAffectCurrent)
rs.update
DataGrid1.Refresh
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Form_Activate()

Label4.Caption = "Hello " + nme
Calendar1_Click
flag = 99
'eid = 5   'temporary
dpt = 99
flag1 = 0
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
End Sub

Private Sub Label7_Click()
con.Close
faculty_event_admin.Show
End Sub

Private Sub Option1_Click(Index As Integer)
dpt = Option1(Index).Tag
End Sub

Private Sub search_Click()
a = Str(eid)
a = Replace(a, " ", "")
query = "select * from bunksheet where event=" + a
subquery = "select roll from student where"

If Check1.Value = 1 Then
 flag = 1
ElseIf Text1.Text <> "" Then
        query = query + " and roll='" + UCase(Text1.Text) + "'"
        If flag1 = 0 Then
            response = MsgBox("Do you want to search for current date only?", vbYesNo + vbQuestion, "SEARCH BY DAY")
        End If
        If response = vbYes Or flag1 = 1 Then
            query = query + " and to_char(day,'fmdd-mon-yyyy')='" + dt + "'"
        End If
        
        flag = 1
Else
        If Text2.Text <> "" Then  'first name
             flag = 1
             subquery = subquery + " first_name= '" + UCase(Text2.Text) + "'"
        End If
        
        If Text3.Text <> "" Then   'last name
           If flag = 1 Then
               subquery = subquery + " and last_name ='" + UCase(Text3.Text) + "'"
           Else
               flag = 1
               subquery = subquery + " last_name = '" + UCase(Text3.Text) + "'"
           End If
        End If
        
        If dpt <> 99 Then  'dept no
            a = Str(dpt)
            a = Replace(a, " ", "")
            If flag = 1 Then
                subquery = subquery + " and dept = " + a
            Else
                flag = 1
                subquery = subquery + " dept = " + a
            End If
        End If
        
        
        If flag1 = 1 Then 'date selected
            'a = Str(select_date)
            'a = Replace(a, " ", "")
            If flag = 1 Then
                subquery = subquery + " and to_char(day,'fmdd-mon-yyyy')='" + dt + "'"
            Else
                subquery = subquery + " to_char(day,'fmdd-mon-yyyy')='" + dt + "'"
                flag = 1
            End If
        Else  'if date not selected then ask
            response = MsgBox("Click yes to search for current day only. else click no", vbYesNo + vbQuestion, "SEARCH BY DAY")
                If response = vbYes Then
                    'a = Str(select_date)
                    'a = Replace(a, " ", "")
                    If flag = 1 Then
                      subquery = subquery + " and to_char(day,'fmdd-mon-yyyy')='" + dt + "'"
                    Else
                       
                       subquery = subquery + " and to_char(day,'fmdd-mon-yyyy')='" + dt + "'"
                        flag = 1
                    End If
                End If
        End If
        If flag = 1 Then
            query = query + " and roll in( " + subquery + " )"
        End If
End If
MsgBox (query)
If flag = 1 Then

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

    Set DataGrid1.DataSource = rs
    If DataGrid1.ApproxCount > 0 Then
    
    DataGrid1.Columns(3).Locked = True
    DataGrid1.Columns(2).Visible = False
    DataGrid1.Columns(3).Visible = False
    i = 0
Do While i < 12

DataGrid1.Columns(i).Width = 1700
DataGrid1.Columns(i).Locked = True
i = i + 1
Loop
Else
    Set DataGrid1.DataSource = Nothing
    MsgBox ("No data found")
End If
    
End If

Check1.Value = Unchecked
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

flag = 0
While flag < 5
    Option1(flag).Value = False
    flag = flag + 1
Wend

End Sub

