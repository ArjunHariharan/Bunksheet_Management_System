VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bunksheet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   30
      Top             =   7320
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   4471
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   2640
      TabIndex        =   29
      Top             =   3480
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2012
      Month           =   10
      Day             =   8
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
   Begin VB.CommandButton Command6 
      Caption         =   "<<BACK"
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
      Left            =   4800
      TabIndex        =   27
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      MaskColor       =   &H8000000F&
      TabIndex        =   25
      Top             =   5520
      Width           =   1335
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
      Left            =   8280
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search Bunksheet"
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
      Left            =   12960
      TabIndex        =   11
      Top             =   10200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mark All"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15840
      MaskColor       =   &H8000000F&
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
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
      TabIndex        =   9
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lecture Timings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   9600
      TabIndex        =   7
      Top             =   3600
      Width           =   7695
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4:30-5:30"
         Height          =   375
         Index           =   7
         Left            =   5520
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3:30-4:30"
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2:30-3:30"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1:30-2:30"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "11:45-12:45"
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "10:45-11:45"
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "9:30-10:30"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8:30-9:30"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "Bunksheet.frx":0000
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   18480
      TabIndex        =   28
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   8880
      TabIndex        =   26
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   16920
      TabIndex        =   16
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label8 
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
      Height          =   495
      Left            =   15480
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   10560
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select  Date:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Name:"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Roll no.:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Event Name:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   6720
      TabIndex        =   1
      Top             =   1200
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bunksheet Management"
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
      Height          =   855
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "bunksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As ADODB.Command
Dim cmd_del As ADODB.Command
Dim cmd1 As ADODB.Command
Dim rs As New ADODB.Recordset
Dim response As Integer
Dim a, check As String
Dim select_date As Date

Public fac_id, fac_name As String
Public fac_type, event_id As Integer
Dim i, j, t, temp1 As Integer

Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub label13_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    Label13.ForeColor = vbRed
End Sub

Private Sub ADD_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub command3_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbGreen
SetCursor LoadCursor(0, IDC_HAND)
End Sub

   ' SetCursor LoadCursor(0, IDC_HAND)

Private Sub add_Click()
i = 0

Do While i < 8  'checks if time entered
    If Check1(i).Value = Unchecked Then
    i = i + 1
    Else
    Exit Do
    End If
Loop

If Text1.Text = "" Then   ' if roll is empty
        MsgBox ("Please enter roll no.!!")
        Exit Sub
ElseIf i = 8 Then  'time not selected
     MsgBox ("Please select timing!!")
     Exit Sub
Else
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandText = "bunk_check_date"   'checks if a record exists for corresponding date and student
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_roll", adVarChar, adParamInput, 30, UCase(Text1.Text))
        .Parameters.Append cmd.CreateParameter("v_date", adDate, adParamInput, 30, select_date)
        .Parameters.Append cmd.CreateParameter("v_flag", adNumeric, adParamOutput, 3)
    End With

    cmd.Execute
    temp1 = cmd("v_flag")

    If temp1 = 1 Then   ' record exists
            response = MsgBox("You want to replace existing record?", vbYesNo + vbQuestion, "Replace")
            If response = vbYes Then
            
                If DataGrid1.ApproxCount > 0 Then  'check in the datagrid if the data exists
                    rs.MoveFirst
                    Do While rs.EOF = False
                        If rs.Fields(0).Value = UCase(Text1.Text) And rs.Fields(1).Value = select_date Then  'if data found the delete
                            rs.delete (adAffectCurrent)  'delete the data
                            'rs.update  'update the database
                            Exit Do
                        End If
                        rs.MoveNext
                    Loop
                End If
            
                If rs.EOF Then  ' if the data doesnt exist in the recordset then delete from the database
                    Set cmd = New ADODB.Command
                    With cmd  ' delete from table if the record doesnt exist in the recordset
                        .ActiveConnection = con
                        .CommandText = "bunk_del"
                        .CommandType = adCmdStoredProc
                        .Parameters.Append cmd.CreateParameter("v_dt", adDate, adParamInput, 30, select_date)
                        .Parameters.Append cmd.CreateParameter("v_roll", adVarChar, adParamInput, 5, UCase(Text1.Text))
                        End With
                    'MsgBox (cmd.CommandText)
                    cmd.Execute ' deletes the record from the dataset
                End If
            End If
    End If
    If response <> vbNo Then
        With rs
            .AddNew
            .Fields(0) = UCase(Text1.Text)
            .Fields(1) = select_date
            .Fields(2) = event_id
            .Fields(3) = fac_id    ' id from login
            i = 0
            Do While i < 8
                j = i + 4
                If Check1(i).Value = Checked Then
                    .Fields(j) = "Y"
                Else
                    .Fields(j) = "N"
                End If
                
                i = i + 1
            Loop

        End With
    End If
End If
i = 0
Do While i < 8
    Check1(i).Value = 0
    i = i + 1
Loop

If temp1 = 0 Then
End If
If response <> vbNo Or temp1 = 0 Then
    Set DataGrid1.DataSource = Nothing
    rs.update
    Set DataGrid1.DataSource = rs
    t = 0
    Do While t < 12
    DataGrid1.Columns(t).Width = 1700
    DataGrid1.Columns(t).Alignment = dbgCenter
    DataGrid1.Columns(i).Locked = True
    t = t + 1
    Loop
    DataGrid1.Columns(2).Visible = False
    DataGrid1.Columns(3).Visible = False

    MsgBox ("Data updated successfully!!")
End If


End Sub

    

Private Sub Calendar1_Click()
select_date = Str(Calendar1.Month) + Str(Calendar1.Day) + Str(Calendar1.Year)
End Sub

Private Sub Command1_Click()
If DataGrid1.ApproxCount > 0 Then
    rs.delete (adAffectCurrent)
    rs.update
    DataGrid1.Refresh
    MsgBox ("Data deleted successfully!!")
Else
    MsgBox ("No entry found!!")
End If
End Sub

Private Sub Command2_Click()
i = 0
Do While i < 8
 Check1(i).Value = 1
 i = i + 1
Loop
End Sub



Private Sub Command3_Click()
Text1.Text = ""
Label9.Caption = ""
Label10.Caption = ""

select_date = Str(Calendar1.Day) + Str(Calendar1.Month) + Str(Calendar1.Year)
i = 0
Do While i < 8
   Check1(i).Value = 0
i = i + 1
Loop
End Sub

Private Sub Command4_Click()
If DataGrid1.ApproxCount > 0 Then
    MsgBox ("Data saved")
End If
bunksheet_display.eid = event_id
con.Close
Me.Hide
bunksheet_display.Show
End Sub

Private Sub Command5_Click()
End Sub


Private Sub Command6_Click()
MsgBox ("Data saved")

con.Close
'Me.Hide
faculty_event_admin.Show
End Sub

Private Sub Form_Activate()
Label12.Caption = "Hello " + fac_name

If fac_type = 1 Then
    Command6.Visible = True
Else
    Command6.Visible = False
End If

If event_id = 1 Then
    Label3.Caption = "AAROHAN"
ElseIf event_id = 2 Then
    Label3.Caption = "GRACIA"
ElseIf event_id = 3 Then
    Label3.Caption = "SPANDAN"
ElseIf event_id = 4 Then
    Label3.Caption = "SPORTS"
ElseIf event_id = 5 Then
    Label3.Caption = "TESLA"
End If

Calendar1_Click
i = 0

temp1 = 0
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"

Set cmd1 = New ADODB.Command
With cmd1
.ActiveConnection = con
.CommandText = "select * from bunksheet where 1 = 2"
.CommandType = adCmdText
End With

With rs
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockPessimistic
.Open cmd1
End With
t = 0

Set DataGrid1.DataSource = rs
Do While t < 12
    DataGrid1.Columns(t).Width = 1700
    DataGrid1.Columns(t).Alignment = dbgCenter
    DataGrid1.Columns(t).Locked = True
    t = t + 1
Loop
DataGrid1.Columns(2).Visible = False
DataGrid1.Columns(3).Visible = False

i = 0


End Sub

Private Sub Label13_Click()
MsgBox ("Data saved")
End

End Sub

Private Sub Text1_LostFocus()
If Len(Text1.Text) = 5 Then
Set cmd = New ADODB.Command


With cmd
        .ActiveConnection = con
        .CommandText = "bunk_name_return"   ' calls the stored procedure to verify the username and password
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("v_roll", adVarChar, adParamInput, 30, UCase(Text1.Text))
        .Parameters.Append cmd.CreateParameter("v_name", adVarChar, adParamOutput, 30)
        .Parameters.Append cmd.CreateParameter("v_dept", adVarChar, adParamOutput, 20)
        .Parameters.Append cmd.CreateParameter("v_class", adVarChar, adParamOutput, 10)
End With
    
    cmd.Execute
        
        check = cmd("v_name")
        If check = "xxxxx" Then
             MsgBox ("Invalid roll no.!")
        Else
             Label9.Caption = check
             Label10.Caption = cmd("v_dept")
            Label11.Caption = cmd("v_class")
        End If
        cmd.Cancel
End If

End Sub

