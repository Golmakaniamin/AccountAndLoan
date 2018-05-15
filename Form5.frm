VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00C3CEC4&
   BorderStyle     =   0  'None
   Caption         =   "„ÊÃÊœÌ"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "⁄„·Ì«  ÃœÌœ"
      Height          =   2775
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   5775
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ò”—"
         Height          =   495
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "«›“«Ì‘"
         Height          =   495
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin KewlButtonz.KewlButtons KewlButtons4 
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "À» "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":10378
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ ⁄„·Ì«  :"
         Height          =   495
         Index           =   1
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ :"
         Height          =   495
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ :"
         Height          =   495
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ê÷⁄Ì "
      Height          =   3855
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6000
      Width           =   5775
      Begin VB.ListBox List4 
         Height          =   2130
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List5 
         Height          =   2130
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox List6 
         Height          =   2130
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin KewlButtonz.KewlButtons KewlButtons5 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":10394
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons6 
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ç«Å —”Ìœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":103B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons7 
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ç«Å"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":103CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ"
         Height          =   495
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€"
         Height          =   495
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ ⁄„·Ì« "
         Height          =   495
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¬Œ—Ì‰ „ÊÃÊœÌ :"
         Height          =   495
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C3CEC4&
      Caption         =   "‰Ê⁄ Õ”«»"
      Height          =   975
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3240
      Width           =   2895
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "⁄«œÌ"
         Height          =   495
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÊÌéÂ"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ã” ÃÊ"
      Height          =   4935
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4320
      Width           =   2895
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   3600
         Width           =   2655
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   135
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   238
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":103E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   135
         Left            =   1920
         TabIndex        =   19
         Top             =   3000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":10404
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   135
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   238
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form5.frx":10420
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‰ ÌÃÂ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Õ”«»"
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2280
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Form5.frx":1043C
      OLEDBString     =   $"Form5.frx":105F0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "Moneyvig"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Form5.frx":107A4
      OLEDBString     =   $"Form5.frx":10958
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "MoneyAdi"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin KewlButtonz.KewlButtons KewlButtons8 
      Height          =   495
      Left            =   11640
      TabIndex        =   13
      Top             =   9480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "»«“ê‘ "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form5.frx":10B0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   2280
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Form5.frx":10B28
      OLEDBString     =   $"Form5.frx":10BC4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "printhesab"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      DataField       =   "radif1"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„ÊÃÊœÌ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Boolean
Dim fso As New FileSystemObject

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Shell "C:\WINDOWS\system32\calc.exe"
End Sub

Private Sub KewlButtons1_Click()
Dim id(1000) As Integer, na(1000), idt, nat, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         idt = id(intq)
         nat = na(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         
         id(intw) = idt
         na(intw) = nat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons2_Click()
Dim id(1000) As Integer, na(1000), idt, nat, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If id(intq) > id(intw) Then
         idt = id(intq)
         nat = na(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         
         id(intw) = idt
         na(intw) = nat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons3_Click()
Dim na(1000), nat, count As String
For intq = 0 To List3.ListCount - 1
    na(intq) = List3.List(intq)
Next intq
count = List3.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         nat = na(intq)
         
         na(intq) = na(intw)
         
         na(intw) = nat
      End If
   Next intw
Next intq
List3.Clear
For intq = 0 To count
   List3.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons4_Click()
Dim lnq As Long
If Label12.Caption = "-" Then
  z = MsgBox("·ÿ›« Õ”«» ‘Œ’ „Ê—œ ‰Ÿ— —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Text1.Text = "") Or (Text2.Text = "") Then
  z = MsgBox("·ÿ›« ›Ì·œ Â«Ì „»·€ Ê Ì«  «—ÌŒ —« Å— ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If Label1.Caption = "«Ì‰ Õ”«» Õ–› ‘œÂ «” " Then
  z = MsgBox("«Ì‰ Õ”«» Õ–› ‘œÂ Ê ⁄„·Ì« Ì ‰„Ì  Ê«‰ —ÊÌ ¬‰ «‰Ã«„ œ«œ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Option3.Value = True) And (Amin.moneyaminnojoda(Text1.Text) > Amin.moneyaminnojoda(Label16.Caption)) Then
  z = MsgBox("„»·€ »—œ«‘ Ì ‘„« »Ì‘ «“ „ÊÃÊœÌ Õ”«» ‘„« „Ì »«‘œ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Option3.Value = False) And (Option4.Value = False) Then
  z = MsgBox("·ÿ›« ⁄„·Ì«  »—œ«‘  Ì« Ê«—Ì“ —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If
p = True
Adodc1.Recordset.MoveFirst
Do
  If Option4.Value = True Then
    If (Adodc1.Recordset.Fields!id = List1.List(List2.ListIndex)) And (Adodc1.Recordset.Fields!Date = Text2.Text) And (Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(Text1.Text)) And (Adodc1.Recordset.Fields!Amal = "«›“«Ì‘") Then
      p = False
      Exit Do
    End If
  End If
  If Option3.Value = True Then
    If (Adodc1.Recordset.Fields!id = List1.List(List2.ListIndex)) And (Adodc1.Recordset.Fields!Date = Text2.Text) And (Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(Text1.Text)) And (Adodc1.Recordset.Fields!Amal = "ò”—") Then
      p = False
      Exit Do
    End If
  End If
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True

If p = False Then
  z = MsgBox("«ÿ·«⁄«  Ê«—œ ‘œÂ  ò—«—Ì „Ì »«‘œ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If


If Option1.Value = True Then
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!id = List1.List(List1.ListIndex)
  Adodc1.Recordset.Fields!Date = Trim(Text2.Text)
  Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(Text1.Text)
  If Option3.Value = True Then Adodc1.Recordset.Fields!Amal = "ò”—"
  If Option4.Value = True Then Adodc1.Recordset.Fields!Amal = "«›“«Ì‘"
  Adodc1.Recordset.Fields!user = Form2.Label2.Caption
  Adodc1.Recordset.Update
  
  Form3.Adodc1.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
  If Option3.Value = True Then Form3.Adodc1.Recordset.Fields!Money = Val(Form3.Adodc1.Recordset.Fields!Money) - Amin.moneyaminnojoda(Text1.Text)
  If Option4.Value = True Then Form3.Adodc1.Recordset.Fields!Money = Val(Form3.Adodc1.Recordset.Fields!Money) + Amin.moneyaminnojoda(Text1.Text)
  Form3.Adodc1.Recordset.Update
  lnq = List2.ListIndex
  Call List2_Click
  z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
  If List2.ListCount < lnq + 1 Then List2.ListIndex = lnq + 1
  Call List2_Click
  List4.ListIndex = List4.ListCount - 1
  z12 = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  ç«Å ‘Êœ", vbMsgBoxRight + vbInformation + vbYesNo, "")
  If z12 = 6 Then
    Form18.Show
    Form18.CRViewer91.Refresh
  End If
End If

If Option2.Value = True Then
  Adodc2.Recordset.AddNew
  Adodc2.Recordset.Fields!id = List1.List(List1.ListIndex)
  Adodc2.Recordset.Fields!Date = Trim(Text2.Text)
  Adodc2.Recordset.Fields!Money = Amin.moneyaminnojoda(Text1.Text)
  If Option3.Value = True Then Adodc2.Recordset.Fields!Amal = "ò”—"
  If Option4.Value = True Then Adodc2.Recordset.Fields!Amal = "«›“«Ì‘"
  Adodc2.Recordset.Fields!user = Form2.Label2.Caption
  Adodc2.Recordset.Update
  
  Form4.Adodc1.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
  If Option3.Value = True Then Form4.Adodc1.Recordset.Fields!Money = Val(Form4.Adodc1.Recordset.Fields!Money) - Amin.moneyaminnojoda(Text1.Text)
  If Option4.Value = True Then Form4.Adodc1.Recordset.Fields!Money = Val(Form4.Adodc1.Recordset.Fields!Money) + Amin.moneyaminnojoda(Text1.Text)
  Form4.Adodc1.Recordset.Update
  lnq = List2.ListIndex
  Call List2_Click
  z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
  If List2.ListCount < lnq + 1 Then List2.ListIndex = lnq + 1
  Call List2_Click
  List4.ListIndex = List4.ListCount - 1
  z12 = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  ç«Å ‘Êœ", vbMsgBoxRight + vbInformation + vbYesNo, "")
  If z12 = 6 Then
    Form18.Show
    Form18.CRViewer91.Refresh
  End If
End If
akhar:
End Sub

Private Sub KewlButtons5_Click()
On Error Resume Next
If List2.ListIndex = -1 Then
  z = MsgBox("·ÿ›« Õ”«» ‘Œ’ „Ê—œ ‰Ÿ— —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If
z = 0
z = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
If z = 6 Then
  If Option1.Value = True Then
    Form3.Adodc1.Recordset.Find "id='" & List1.List(List2.ListIndex) & "'", , adSearchForward, 1
    If List4.List(List4.ListIndex) = "«›“«Ì‘" Then Form3.Adodc1.Recordset.Fields!Money = Val(Form3.Adodc1.Recordset.Fields!Money) - Val(Amin.moneyaminnojoda(List5.List(List5.ListIndex)))
    If List4.List(List4.ListIndex) = "ò”—" Then Form3.Adodc1.Recordset.Fields!Money = Val(Form3.Adodc1.Recordset.Fields!Money) + Val(Amin.moneyaminnojoda(List5.List(List5.ListIndex)))
    Form3.Adodc1.Recordset.Update
    
    Adodc1.Recordset.MoveFirst
    Do
      If (Adodc1.Recordset.Fields!id = List1.List(List2.ListIndex)) And (Adodc1.Recordset.Fields!Date = List6.List(List6.ListIndex)) And (Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(List5.List(List5.ListIndex))) And (Adodc1.Recordset.Fields!Amal = List4.List(List4.ListIndex)) Then
        Adodc1.Recordset.Delete
        Adodc1.Refresh
        Exit Do
      End If
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
    List4.RemoveItem (List4.ListIndex)
    List5.RemoveItem (List5.ListIndex)
    List6.RemoveItem (List6.ListIndex)
  End If
  
  If Option2.Value = True Then
    Form4.Adodc1.Recordset.Find "id='" & List1.List(List2.ListIndex) & "'", , adSearchForward, 1
    If List4.List(List4.ListIndex) = "«›“«Ì‘" Then Form4.Adodc1.Recordset.Fields!Money = Val(Form4.Adodc1.Recordset.Fields!Money) - Val(Amin.moneyaminnojoda(List5.List(List5.ListIndex)))
    If List4.List(List4.ListIndex) = "ò”—" Then Form4.Adodc1.Recordset.Fields!Money = Val(Form4.Adodc1.Recordset.Fields!Money) + Val(Amin.moneyaminnojoda(List5.List(List5.ListIndex)))
    Form4.Adodc1.Recordset.Update
    Adodc2.Recordset.MoveFirst
    Do
      If (Adodc2.Recordset.Fields!id = List1.List(List2.ListIndex)) And (Adodc2.Recordset.Fields!Date = List6.List(List6.ListIndex)) And (Adodc2.Recordset.Fields!Money = Amin.moneyaminnojoda(List5.List(List5.ListIndex))) And (Adodc2.Recordset.Fields!Amal = List4.List(List4.ListIndex)) Then
        Adodc2.Recordset.Delete
        Adodc2.Refresh
        Exit Do
      End If
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
    List4.RemoveItem (List4.ListIndex)
    List5.RemoveItem (List5.ListIndex)
    List6.RemoveItem (List6.ListIndex)
  End If
  Call List2_Click
  z = MsgBox("⁄„·Ì«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbMsgBoxRight + vbInformation, "")
End If
akhar:
End Sub

Private Sub KewlButtons6_Click()
If List4.ListIndex <> -1 Then
  Form18.Show
Else
  z = MsgBox("·ÿ›« Ã«»Ã«ÌÌ Õ”«» —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons7_Click()
If List2.ListIndex <> -1 Then
  fso.CopyFile "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\Data\info2.mdb", "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\info2.mdb", True
  Adodc3.Refresh
  If List4.ListCount >= 1 Then
    For q = 0 To List4.ListCount Step 2
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!radif1 = q + 1
      If List4.List(q) = "«›“«Ì‘" Then df = "«›“«Ì‘ „»·€ : " + List5.List(q) + " —Ì«· œ—  «—ÌŒ : " + List6.List(q)
      If List4.List(q) = "ò”—" Then df = "ò”— „»·€ : " + List5.List(q) + " —Ì«· œ—  «—ÌŒ : " + List6.List(q)
      Adodc3.Recordset.Fields!promp1 = df
      Adodc3.Recordset.Fields!radif2 = q + 2
      If List4.List(q + 1) = "«›“«Ì‘" Then df = "«›“«Ì‘ „»·€ : " + List5.List(q + 1) + " —Ì«· œ—  «—ÌŒ : " + List6.List(q + 1)
      If List4.List(q + 1) = "ò”—" Then df = "ò”— „»·€ : " + List5.List(q + 1) + " —Ì«· œ—  «—ÌŒ : " + List6.List(q + 1)
      Adodc3.Recordset.Fields!promp2 = df
      Adodc3.Recordset.Update
    Next q
  End If
  
  If (List4.ListCount Mod 2) = 1 Then
    Adodc3.Refresh
    Adodc3.Recordset.AddNew
    Adodc3.Recordset.Fields!radif1 = Val(List4.ListCount)
    If List4.List(q) = "«›“«Ì‘" Then df = "«›“«Ì‘ „»·€ : " + List5.List(List5.ListCount) + " —Ì«· œ—  «—ÌŒ : " + List6.List(List5.ListCount)
    If List4.List(q) = "ò”—" Then df = "ò”— „»·€ : " + List5.List(List5.ListCount) + " —Ì«· œ—  «—ÌŒ : " + List6.List(List5.ListCount)
    Adodc3.Recordset.Fields!promp1 = df
    Adodc3.Recordset.Fields!radif2 = 0
    Adodc3.Recordset.Fields!promp2 = ""
    Adodc3.Recordset.Update
  End If
  Adodc3.Recordset.Sort = "radif1"
  Form33.Show
Else
  z = MsgBox("·ÿ›« Õ”«» —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons8_Click()
Form2.Show
Me.Hide
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
Adodc1.Refresh
Adodc2.Refresh
List1.ListIndex = List2.ListIndex
Label12.Caption = List2.Text
List4.Clear
List5.Clear
List6.Clear
If (Option1.Value = True) And (Adodc1.Recordset.RecordCount > 0) Then
  Adodc1.Recordset.Sort = "date"
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!id = List1.List(List2.ListIndex) Then
      List4.AddItem Adodc1.Recordset.Fields!Amal
      List5.AddItem Amin.moneyaminjoda(Adodc1.Recordset.Fields!Money)
      List6.AddItem Adodc1.Recordset.Fields!Date
    End If
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
  
  Form3.Adodc1.Recordset.Find "id='" & List1.List(List2.ListIndex) & "'", , adSearchForward, 1
  Label16.Caption = Amin.moneyaminjoda(Form3.Adodc1.Recordset.Fields!Money)
  If Form3.Adodc1.Recordset.Fields!Delete = "‰‘œÂ" Then
    Label1.Caption = "«Ì‰ Õ”«» Õ–› ‰‘œÂ «” "
  Else
    Label1.Caption = "«Ì‰ Õ”«» Õ–› ‘œÂ «” "
  End If
End If

If (Option2.Value = True) And (Adodc2.Recordset.RecordCount > 0) Then
  Adodc1.Recordset.Sort = "date"
  Adodc2.Recordset.MoveFirst
  Do
    If Adodc2.Recordset.Fields!id = List1.List(List2.ListIndex) Then
      List4.AddItem Adodc2.Recordset.Fields!Amal
      List5.AddItem Amin.moneyaminjoda(Adodc2.Recordset.Fields!Money)
      List6.AddItem Adodc2.Recordset.Fields!Date
    End If
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
  
  Form4.Adodc1.Recordset.Find "id='" & List1.List(List2.ListIndex) & "'", , adSearchForward, 1
  Label16.Caption = Amin.moneyaminjoda(Form4.Adodc1.Recordset.Fields!Money)
  If Form4.Adodc1.Recordset.Fields!Delete = "‰‘œÂ" Then
    Label1.Caption = "«Ì‰ Õ”«» Õ–› ‰‘œÂ «” "
  Else
    Label1.Caption = "«Ì‰ Õ”«» Õ–› ‘œÂ «” "
  End If
End If
Text2.Text = Form2.Label5.Caption
Text1.Text = ""
End Sub

Private Sub List3_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
End Sub

Private Sub List4_Click()
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List4.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
End Sub

Private Sub Option1_Click()
Label12.Caption = "-"
Label16.Caption = "0"
Label1.Caption = ""
List1.Clear
List2.Clear
List4.Clear
List5.Clear
List6.Clear
If Form3.Adodc1.Recordset.RecordCount > 0 Then
  Form3.Adodc1.Recordset.MoveFirst
  Do
    List1.AddItem Form3.Adodc1.Recordset.Fields!id
    List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    Form3.Adodc1.Recordset.MoveNext
  Loop Until Form3.Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Option2_Click()
Label12.Caption = "-"
Label16.Caption = "0"
Label1.Caption = ""
List1.Clear
List2.Clear
List4.Clear
List5.Clear
List6.Clear
If Form4.Adodc1.Recordset.RecordCount > 0 Then
  Form4.Adodc1.Recordset.MoveFirst
  Do
    List1.AddItem Form4.Adodc1.Recordset.Fields!id
    List2.AddItem Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
    Form4.Adodc1.Recordset.MoveNext
  Loop Until Form4.Adodc1.Recordset.EOF = True
End If
End Sub


Private Sub Option3_Click()
Text2.Text = Form2.Label5.Caption
End Sub

Private Sub Option4_Click()
Text2.Text = Form2.Label5.Caption
End Sub


Private Sub Text1_Change()
Label18.Caption = Amin.moneyaminjoda(Text1.Text)
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Label18.Visible = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text1_LostFocus()
Label18.Visible = False
Text1.Text = Label18.Caption
End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List1.ListCount - 1
     If Trim(Text15.Text) = List1.List(q) Then List1.ListIndex = q
   Next q
End If
End Sub

Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   List3.Clear
   For q = 0 To List2.ListCount - 1
       If InStr(List2.List(q), Trim(Text16.Text)) <> 0 Then
          List3.AddItem List2.List(q)
       End If
   Next q
End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons4.SetFocus
End Sub
