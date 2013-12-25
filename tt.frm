VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Timetable 
   BackColor       =   &H80000005&
   Caption         =   "Time Table"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Height          =   735
      Left            =   960
      TabIndex        =   69
      Top             =   3240
      Width           =   14775
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   78
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "8.30-9.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   77
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "9.30-10.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "10.45-11.45"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   75
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "11.45-12.45"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6360
         TabIndex        =   74
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "1.30-2.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8160
         TabIndex        =   73
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "2.30-3.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   9840
         TabIndex        =   72
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "3.30-4.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   11400
         TabIndex        =   71
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "4.30-5.30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   12960
         TabIndex        =   70
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton proceed 
      Caption         =   " NEXT BATCH"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   68
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton savep 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   67
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame prac 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1800
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   13935
      Begin VB.ComboBox eroll 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   9960
         TabIndex        =   66
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox sroll 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4560
         TabIndex        =   65
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "BATCH 1:"
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
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "End roll no:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8160
         TabIndex        =   63
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Start roll no:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   30
      Left            =   4080
      ScaleHeight     =   30
      ScaleWidth      =   135
      TabIndex        =   60
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton save 
      Caption         =   "SAVE AND PROCEED"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   59
      Top             =   9360
      Width           =   2295
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   600
      TabIndex        =   39
      Top             =   6960
      Width           =   15375
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   31
         Left            =   12840
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   30
         Left            =   11160
         TabIndex        =   46
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   29
         Left            =   9480
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   28
         Left            =   7800
         TabIndex        =   44
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   27
         Left            =   6120
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   26
         Left            =   4440
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   25
         Left            =   2760
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   24
         Left            =   1080
         TabIndex        =   40
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Thursday:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   9
      Left            =   3360
      TabIndex        =   22
      Top             =   5280
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   15015
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   2
         Left            =   360
         TabIndex        =   49
         Top             =   4920
         Width           =   15375
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   39
            Left            =   12840
            TabIndex        =   57
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   38
            Left            =   11160
            TabIndex        =   56
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   37
            Left            =   9480
            TabIndex        =   55
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   36
            Left            =   7800
            TabIndex        =   54
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   35
            Left            =   6120
            TabIndex        =   53
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   34
            Left            =   4440
            TabIndex        =   52
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   33
            Left            =   2760
            TabIndex        =   51
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   32
            Left            =   1080
            TabIndex        =   50
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Friday:"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   15375
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   23
            Left            =   12840
            TabIndex        =   38
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   22
            Left            =   11160
            TabIndex        =   37
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   21
            Left            =   9480
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   20
            Left            =   7800
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   19
            Left            =   6120
            TabIndex        =   34
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   18
            Left            =   4440
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   17
            Left            =   2760
            TabIndex        =   32
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   16
            Left            =   1080
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Wednesday:"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   15375
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   15
            Left            =   12840
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   14
            Left            =   11160
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   13
            Left            =   9480
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   12
            Left            =   7800
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   11
            Left            =   6120
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   10
            Left            =   4440
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   8
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Tuesday:"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   15375
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   7
            Left            =   12840
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   6
            Left            =   11160
            TabIndex        =   15
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   5
            Left            =   9480
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   4
            Left            =   7800
            TabIndex        =   13
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   3
            Left            =   6120
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   2
            Left            =   4440
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   0
            Left            =   1080
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Monday:"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1335
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   13575
      Begin VB.ComboBox div 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   6600
         TabIndex        =   81
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton insert 
         Caption         =   "INSERT"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox class 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "DIV:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "CLASS:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3480
      Top             =   10320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18480
      TabIndex        =   80
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17280
      TabIndex        =   79
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   0
      Picture         =   "tt.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   16320
      Picture         =   "tt.frx":0CD2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4005
   End
   Begin VB.Label Label3 
      Caption         =   "8.30-9.30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   18
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "8.30-9.30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   17
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TIMETABLE"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Timetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd As New ADODB.Command
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Dim flag As Boolean
Dim flag1 As Boolean
Dim flag2 As Boolean
Dim flag3 As Boolean
Dim flag4 As Boolean
Dim flag5 As Boolean
Dim c1 As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim row As Integer
Dim col As Integer
Dim co As Integer
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s4 As String
Dim s5 As String
Dim a1 As String
Dim a2 As String
Option Explicit
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal _
    hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long


Private Sub Form_Activate()

j = 1
i = 0
flag = False
flag1 = False
flag2 = False
flag3 = False
flag4 = False
flag5 = False
prac.Visible = False
con.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl"
For i = 0 To 39
Combo1(i).AddItem "select"
Combo1(i).Text = "select"
Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = vbRed
Label10.ForeColor = vbRed
End Sub

Private Sub LABEL9_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label9.ForeColor = &H8000000D 'color on hovering
End Sub

Private Sub LABEL10_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label10.ForeColor = &H8000000D 'color on hovering
End Sub

Private Sub insert_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label10.ForeColor = &H8000000D 'color on hovering
End Sub

Private Sub savep_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label10.ForeColor = &H8000000D 'color on hovering
End Sub

Private Sub save_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label10.ForeColor = &H8000000D 'color on hovering
End Sub

Private Sub proceed_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
 '   Label4.ForeColor = vbBlue
       Label10.ForeColor = &H8000000D 'color on hovering
End Sub


Private Sub Combo1_Click(Index As Integer)
row = 0
col = 0

For i = 0 To 39
If i Mod 2 = 0 Then
If Combo1(i).Text = "practical" Then
Combo1(i + 1).Enabled = False
Combo1(i + 1).Text = "practical"
Else
If flag1 = False Then
Combo1(i + 1).Enabled = True
End If
End If
End If
Next i

If flag = False Then
i = 0
row = Index / 8
col = (Index + 3) Mod 8
rs1.MoveFirst
For i = 0 To row
rs1.MoveNext
Next i
rs1.MovePrevious
rs1.Fields(col).Value = Combo1(Index).Text
End If

End Sub



Private Sub edit_Click()
i = 0
flag = False
While i < 40

With rs1
   Combo1(i).Text = .Fields(3).Value
   Combo1(i + 1).Text = .Fields(4).Value
   Combo1(i + 2).Text = .Fields(5).Value
   Combo1(i + 3).Text = .Fields(6).Value
   Combo1(i + 4).Text = .Fields(7).Value
   Combo1(i + 5).Text = .Fields(8).Value
   Combo1(i + 6).Text = .Fields(9).Value
   Combo1(i + 7).Text = .Fields(10).Value
End With
i = i + 8
rs1.MoveNext
Wend
End Sub

Private Sub Form_Load()

End Sub

Private Sub insert_Click()
c1 = class(0).Text + "-"
c1 = c1 + div(1).Text
MsgBox (c1)
If flag4 = False Then
With cmd
.ActiveConnection = con
.CommandText = " Select abbr,type from sem_subject where yr='" + class(0).Text + "'"
.CommandType = adCmdText
End With
'MsgBox (cmd.CommandText)

    With rs
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockPessimistic
.Open cmd
    End With
Else
rs.Close
With cmd
.ActiveConnection = con
.CommandText = " Select abbr,type from sem_subject where yr='" + class(0).Text + "'"
.CommandType = adCmdText
End With
'MsgBox (cmd.CommandText)

With rs
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockPessimistic
.Open cmd
End With

End If
'Set DataGrid1.DataSource = rs

'for i=0 to

For i = 0 To 39
    rs.MoveFirst
    While rs.EOF <> True
        If Not IsNull(rs("abbr")) And rs("type").Value = 1 Then
        Combo1(i).AddItem rs("abbr")
    End If
    rs.MoveNext
    Combo1(i).Refresh
    Wend
    If i Mod 2 = 0 Then
        Combo1(i).AddItem "practical"
    End If
Next i
If flag4 = False Then
With cmd1
.ActiveConnection = con
.CommandText = "select * from time_table where yr='" + c1 + "'"
.CommandType = adCmdText
End With

'MsgBox (cmd1.CommandText)
With rs1
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockPessimistic
    .Open cmd1
End With
Else
rs1.Close
With cmd1
.ActiveConnection = con
.CommandText = "select * from time_table where yr='" + c1 + "'"
.CommandType = adCmdText
End With

'MsgBox (cmd1.CommandText)
With rs1
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockPessimistic
    .Open cmd1
End With
End If
If rs1.EOF = True Then
flag = True
End If

'Set DataGrid2.DataSource = rs1
'Set DataGrid2.DataSource = rs1
flag4 = True
i = 0
'flag = False
If rs1.EOF <> True Then
rs1.MoveFirst
While i < 40

With rs1
   Combo1(i).Text = .Fields(3).Value
   Combo1(i + 1).Text = .Fields(4).Value
   Combo1(i + 2).Text = .Fields(5).Value
   Combo1(i + 3).Text = .Fields(6).Value
   Combo1(i + 4).Text = .Fields(7).Value
   Combo1(i + 5).Text = .Fields(8).Value
   Combo1(i + 6).Text = .Fields(9).Value
   Combo1(i + 7).Text = .Fields(10).Value
End With
i = i + 8
rs1.MoveNext
Wend
End If
End Sub

Private Sub Label9_Click()
MsgBox ("HOME")
End Sub

Private Sub proceed_Click()

a1 = Left$(eroll.Text, 3)
a2 = Right$(eroll.Text, 2)
sroll.Clear
eroll.Clear
a2 = Val(a2)


a2 = a2 + 1
If a2 < 9 Then
a2 = Str(a2)
a2 = a2 + "0"
End If
a1 = a1 + a2
sroll.Text = a1
a2 = Val(a2)
For i = a2 To 60
If i < 10 Then
s5 = s1 + "0"
s5 = s5 & i
Else
s5 = s1 & i
End If
eroll.AddItem s5
Next i



For i = 0 To 39
If Combo1(i).Enabled = True Then
Combo1(i).Text = "practical"
End If
Next i
End Sub
    
Private Sub save_Click()
If flag = True Then
    i = 0
    j = 1
    While i < 40 And j < 6
        With rs1
        .AddNew
        .Fields(0).Value = c1
        .Fields(1).Value = 1
        .Fields(2).Value = j
        .Fields(3).Value = Combo1(i).Text
        .Fields(4).Value = Combo1(i + 1).Text
        .Fields(5).Value = Combo1(i + 2).Text
        .Fields(6).Value = Combo1(i + 3).Text
        .Fields(7).Value = Combo1(i + 4).Text
        .Fields(8).Value = Combo1(i + 5).Text
        .Fields(9).Value = Combo1(i + 6).Text
        .Fields(10).Value = Combo1(i + 7).Text
        j = j + 1
        i = i + 8
        End With
    Wend
End If
rs1.update
'DataGrid2.Visible = True
Frame1.Visible = False
prac.Visible = True

i = 0
j = 3


    With rs1
    rs1.MoveFirst
    While rs1.Fields(0) <> c1
    rs1.MoveNext
    Wend
    
        While i < 40 And j < 11
        If .Fields(3).Value <> "practical" Then
            Combo1(i).Enabled = False
        End If
        
        If .Fields(4).Value <> "practical" Then
            Combo1(i + 1).Enabled = False
        End If
        
        If .Fields(5).Value <> "practical" Then
            Combo1(i + 2).Enabled = False
        End If

        If .Fields(6).Value <> "practical" Then
            Combo1(i + 3).Enabled = False
        End If

        If .Fields(7).Value <> "practical" Then
            Combo1(i + 4).Enabled = False
        End If

        If .Fields(8).Value <> "practical" Then
            Combo1(i + 5).Enabled = False
        End If

        If .Fields(9).Value <> "practical" Then
            Combo1(i + 6).Enabled = False
        End If
        
        If .Fields(10).Value <> "practical" Then
            Combo1(i + 7).Enabled = False
        End If

    i = i + 8
    rs1.MoveNext
        
        Wend
        End With
        With cmd2
                    .ActiveConnection = con
                    .CommandText = "select * from practical where class='" + c1 + "'"
                    .CommandType = adCmdText
                End With
                                
                With rs2
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockPessimistic
                    .Open cmd2
                End With
                If rs2.EOF <> True Then
                flag5 = True
                rs2.MoveFirst
                 End If
           If flag5 = True Then
           While rs2.Fields(1) <> c1
            rs2.MoveNext
            Wend
    
            i = 0
        
            
                While i < 40
                    If i Mod 8 = 0 And i <> 0 Then
                            rs2.MoveNext
                    End If
                    sroll.Text = rs2.Fields(3).Value
                    eroll.Text = rs2.Fields(4).Value
                
                    With rs2
                        If Combo1(i).Text = "practical" Then
                         If .Fields(5).Value = "1" Then
                                Combo1(i).Text = .Fields(2).Value
            rs.MoveFirst
            
            While rs.EOF <> True
                If Not IsNull(rs("abbr")) And rs("type").Value = 0 Then
                    'If Combo1(i).Text = "practical" Then
                     Combo1(i).AddItem rs("abbr")
                    End If
                'End If
            rs.MoveNext
            Combo1(i).Refresh
            Wend
        
                         End If
                        End If
                      i = i + 1
   
                    End With

                Wend

            Else
                        
            End If

            
rs2.Close
'End If

flag1 = True

        For i = 0 To 39
            If Combo1(i).Text = "practical" Then
                Combo1(i).Clear
                Combo1(i).Text = "practical"
            End If
        Next i

        For i = 0 To 39
            rs.MoveFirst
            
            While rs.EOF <> True
                If Not IsNull(rs("abbr")) And rs("type").Value = 0 Then
                    If Combo1(i).Text = "practical" Then
                     Combo1(i).AddItem rs("abbr")
                    End If
                End If
            rs.MoveNext
            Combo1(i).Refresh
            Wend
        Next i
 
 
    s1 = class(0).Text    'combo1 - class(FE/SE/TE/BE)
        If s1 = "FE" Then
            s1 = "F"

        ElseIf s1 = "SE" Then
            s1 = "S"

        ElseIf s1 = "TE" Then
            s1 = "T"

        Else
            s1 = "B"
        End If
 
    s1 = s1 & "3"
    s2 = div(1).Text    'combo2 - div(A/B)
        If s2 = "A" Then
            s1 = s1 & "1"
    
        ElseIf s2 = "B" Then
            s1 = s1 & "2"
    
        ElseIf s2 = "C" Then
            s1 = s1 + "3"

        ElseIf s2 = "D" Then
            s1 = s1 + "4"

        ElseIf s2 = "E" Then
            s1 = s1 + "5"
    
        End If

        For i = 1 To 60
            If i < 10 Then
                s5 = s1 + "0"
                s5 = s5 & i
            Else
                s5 = s1 & i
            End If


                sroll.AddItem s5
                eroll.AddItem s5

        Next i
        
                
save.Visible = False

savep.Visible = True
proceed.Visible = True
End Sub


Private Sub savep_Click()
flag5 = True
i = 0
If flag3 = False Then
    With cmd2
    .ActiveConnection = con
    .CommandText = "select * from practical where class='" + c1 + "'"
    .CommandType = adCmdText
    End With

'MsgBox (cmd1.CommandText)
    With rs2
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockPessimistic
    .Open cmd2
    End With
Else
    rs2.MoveLast
    rs2.MoveNext
End If
'If rs2.EOF = True Then
'flag = True
'End If
 


    With rs2
    rs1.MoveFirst
    While i < 40
        If i Mod 8 = 0 And i <> 0 Then
            rs1.MoveNext
        End If
    If Combo1(i).Enabled = True Or Combo1(i).Text = "practical" Then
        .AddNew
        .Fields(0).Value = rs1.Fields(2).Value
        .Fields(1).Value = c1
     If Combo1(i).Enabled = False Or Combo1(i).Text = "practical" Then
        .Fields(2).Value = Combo1(i - 1).Text
     Else
        .Fields(2).Value = Combo1(i).Text
     End If
        .Fields(3).Value = sroll.Text
        .Fields(4).Value = eroll.Text
    If Label8.Caption = "BATCH 1:" Then
        .Fields(5).Value = "1"
    ElseIf Label8.Caption = "BATCH 2:" Then
        .Fields(5).Value = "2"
    Else
        .Fields(5).Value = "3"
    End If
        .Fields(6).Value = 1
            co = i Mod 8
        .Fields(7).Value = rs1.Fields(co + 3).Name
                    rs2.update
    End If
        i = i + 1
Wend
    End With
    
If Label8.Caption = "BATCH 3:" Then
proceed.Enabled = False
proceed.Visible = False

Else

    If flag2 = False Then
       Label8.Caption = "BATCH 2:"
        flag2 = True
    Else
        Label8.Caption = "BATCH 3:"
    End If
End If

flag3 = True
flag5 = True
End Sub

