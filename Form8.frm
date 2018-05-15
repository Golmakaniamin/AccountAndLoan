VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   Caption         =   "œ—Ì«›  «ﬁ”«ÿ"
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
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ê÷⁄Ì "
      Height          =   3255
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   6720
      Width           =   7935
      Begin VB.ListBox List1 
         Height          =   2130
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   2130
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Height          =   2130
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List8 
         Height          =   2130
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List9 
         Height          =   2130
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin KewlButtonz.KewlButtons KewlButtons6 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "Form8.frx":10378
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
         Left            =   1200
         TabIndex        =   11
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "Form8.frx":10394
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "—œÌ›"
         Height          =   495
         Index           =   14
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ œ—Ì«› "
         Height          =   495
         Index           =   6
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€"
         Height          =   495
         Index           =   7
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "”— —”Ìœ «ﬁ”«ÿ"
         Height          =   495
         Index           =   3
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«„ Ì«“"
         Height          =   495
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.ListBox List7 
      Height          =   1785
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ã” ÃÊ"
      Height          =   5055
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   4200
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
         TabIndex        =   5
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
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List6 
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
      Begin VB.ListBox List5 
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "B Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3480
         Width           =   2655
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   135
         Left            =   120
         TabIndex        =   39
         Top             =   4800
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
         MICON           =   "Form8.frx":103B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons4 
         Height          =   135
         Left            =   1920
         TabIndex        =   40
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
         MICON           =   "Form8.frx":103CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons5 
         Height          =   135
         Left            =   120
         TabIndex        =   41
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
         MICON           =   "Form8.frx":103E8
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
         TabIndex        =   44
         Top             =   3120
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
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Ê«„"
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
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "œ—Ì«›  ﬁ”ÿ"
      Height          =   2175
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4440
      Width           =   3495
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   480
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         MICON           =   "Form8.frx":10404
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ﬁ”ÿ"
         Height          =   495
         Index           =   1
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ œ—Ì«› "
         Height          =   495
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "«‰ Œ«» ‰Ê⁄ Ê«„ "
      Height          =   975
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   2895
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÊÌéÂ"
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "«÷ÿ—«—Ì"
         Height          =   495
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "⁄«œÌ"
         Height          =   495
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11280
      TabIndex        =   13
      Top             =   9360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "»«“ê‘ "
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
      MICON           =   "Form8.frx":10420
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
      Height          =   330
      Left            =   2280
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      Connect         =   $"Form8.frx":1043C
      OLEDBString     =   $"Form8.frx":105F0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "GvamVig"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2280
      Top             =   1800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"Form8.frx":107A4
      OLEDBString     =   $"Form8.frx":10958
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "GvamAz"
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
      Height          =   330
      Left            =   2280
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"Form8.frx":10B0C
      OLEDBString     =   $"Form8.frx":10CC0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "GvamAdi"
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» Ê«„ êÌ—‰œÂ :"
      Height          =   495
      Index           =   13
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Index           =   12
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Label9"
      DataField       =   "id"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ »«ﬁÌ„«‰œÂ :"
      Height          =   495
      Index           =   11
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ œ—Ì«› Ì :"
      Height          =   495
      Index           =   10
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Å—œ«Œ  :"
      Height          =   495
      Index           =   9
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ò«—„“œ :"
      Height          =   495
      Index           =   8
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «ﬁ”«ÿ :"
      Height          =   495
      Index           =   5
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ò·Ì Ê«„ :"
      Height          =   495
      Index           =   4
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ—Ì«›  «ﬁ”«ÿ"
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
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Boolean, strdate1 As String, strdate2 As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Shell "C:\WINDOWS\system32\calc.exe"
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons2_Click()
Dim q1 As Integer, w1 As Integer
If (Combo1.Text <> "") And (Text1(1).Text <> "") Then
  If Option1.Value = True Then
    If List1.ListCount = 0 Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc1.Recordset.Fields!rad = List1.ListCount + 1
      Adodc1.Recordset.Fields!Date = Text1(1).Text
      Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(Label4(2).Caption, "1")
      Adodc1.Recordset.Fields!saragsat = strdate2
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc1.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc1.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc1.Recordset.Update
    Else
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc1.Recordset.Fields!rad = List1.ListCount + 1
      Adodc1.Recordset.Fields!Date = Text1(1).Text
      Adodc1.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(List8.List(List8.ListCount - 1), "1")
      Adodc1.Recordset.Fields!saragsat = strdate2
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc1.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc1.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc1.Recordset.Update
    End If
  End If
    
  If Option2.Value = True Then
    If List1.ListCount = 0 Then
      Adodc2.Recordset.AddNew
      Adodc2.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc2.Recordset.Fields!rad = List1.ListCount + 1
      Adodc2.Recordset.Fields!Date = Text1(1).Text
      Adodc2.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(Label4(2).Caption, "1")
      Adodc2.Recordset.Fields!saragsat = strdate2
      Adodc2.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc2.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc2.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc2.Recordset.Update
    Else
      Adodc2.Recordset.AddNew
      Adodc2.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc2.Recordset.Fields!rad = List1.ListCount + 1
      Adodc2.Recordset.Fields!Date = Text1(1).Text
      Adodc2.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(List8.List(List8.ListCount - 1), "1")
      Adodc2.Recordset.Fields!saragsat = strdate2
      Adodc2.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc2.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc2.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc2.Recordset.Update
    End If
  End If
  If Option3.Value = True Then
    If List1.ListCount = 0 Then
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc3.Recordset.Fields!rad = List1.ListCount + 1
      Adodc3.Recordset.Fields!Date = Text1(1).Text
      Adodc3.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(Label4(2).Caption, "1")
      Adodc3.Recordset.Fields!saragsat = strdate2
      Adodc3.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc3.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc3.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc3.Recordset.Update
    Else
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!id = List6.List(List6.ListIndex)
      Adodc3.Recordset.Fields!rad = List1.ListCount + 1
      Adodc3.Recordset.Fields!Date = Text1(1).Text
      Adodc3.Recordset.Fields!Money = Amin.moneyaminnojoda(Combo1.Text)
      strdate2 = Amin.dateaminEzafeMoon(List8.List(List8.ListCount - 1), "1")
      Adodc3.Recordset.Fields!saragsat = strdate2
      Adodc3.Recordset.Fields!user = Form2.Label2.Caption
      
      If Text1(1).Text > strdate2 Then
        Adodc3.Recordset.Fields!emteyaz = -1 * (Amin.dateaminEktelaf(strdate2, Text1(1).Text))
      Else
        Adodc3.Recordset.Fields!emteyaz = Amin.dateaminEktelaf(Text1(1).Text, strdate2)
      End If
      Adodc3.Recordset.Update
    End If
  End If
  If Option1.Value = True Then q1 = 1
  If Option2.Value = True Then q1 = 2
  If Option3.Value = True Then q1 = 3
  Call List6_Click
  z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
  If Label4(1).Caption = List1.ListCount Then
    If Option1.Value = True Then
      Form7.Adodc1.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
      Form7.Adodc1.Recordset.Fields!tasvie = "‘œÂ"
      Form7.Adodc1.Recordset.Update
    End If
    
    If Option2.Value = True Then
      Form7.Adodc2.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
      Form7.Adodc2.Recordset.Fields!tasvie = "‘œÂ"
      Form7.Adodc2.Recordset.Update
    End If
    
    If Option3.Value = True Then
      Form7.Adodc3.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
      Form7.Adodc3.Recordset.Fields!tasvie = "‘œÂ"
      Form7.Adodc3.Recordset.Update
    End If
    z = MsgBox("«ﬁ”«ÿ ‘„« »Â Å«Ì«‰ —”ÌœÂ «” ", vbMsgBoxRight + vbInformation, "")
  End If
  If q1 = 1 Then Option1.Value = True
  If q1 = 2 Then Option2.Value = True
  If q1 = 3 Then Option3.Value = True
  Call List6_Click
  List1.ListIndex = List1.ListCount - 1
  
  z12 = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «ÿ·«⁄«  ç«Å ‘Êœ", vbMsgBoxRight + vbInformation + vbYesNo, "")
  If z12 = 6 Then
    Form17.Show
    Form17.CRViewer91.Refresh
  End If
Else
  z = MsgBox("·ÿ›« ›Ì·œ Â«Ì „—»ÊÿÂ —«  ò„Ì· ‰„«ÌÌœ", vbMsgBoxRight + vbCritical, "")
End If
End Sub

Private Sub KewlButtons3_Click()
Dim na(1000), nat, count As String
For intq = 0 To List4.ListCount - 1
    na(intq) = List4.List(intq)
Next intq
count = List4.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         nat = na(intq)
         
         na(intq) = na(intw)
         
         na(intw) = nat
      End If
   Next intw
Next intq
List4.Clear
For intq = 0 To count
   List4.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons4_Click()
Dim id(1000), na(1000) As Integer, idt, nat, count As String
For intq = 0 To List5.ListCount - 1
    id(intq) = List5.List(intq)
    na(intq) = List6.List(intq)
Next intq
count = List5.ListCount - 1
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
List5.Clear
List6.Clear
For intq = 0 To count
   List5.AddItem id(intq)
   List6.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons5_Click()
Dim id(1000), na(1000) As Integer, idt, nat, count As String
For intq = 0 To List5.ListCount - 1
    id(intq) = List5.List(intq)
    na(intq) = List6.List(intq)
Next intq
count = List5.ListCount - 1
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
List5.Clear
List6.Clear
For intq = 0 To count
   List5.AddItem id(intq)
   List6.AddItem na(intq)
Next intq
End Sub

Private Sub KewlButtons6_Click()
If List1.ListIndex <> -1 Then
  If q1 = 1 Then Option1.Value = True
  If q1 = 2 Then Option2.Value = True
  If q1 = 3 Then Option3.Value = True
  Text15.SetFocus
  SendKeys ("{enter}")
  Form17.Show
Else
  z = MsgBox("·ÿ›« —œÌ› ﬁ”ÿ —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons7_Click()
On Error Resume Next
If List2.ListIndex = -1 Then
  z = MsgBox("·ÿ›« Ê«„ „Ê—œ ‰Ÿ— —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If List2.ListIndex <> List2.ListCount - 1 Then
  z = MsgBox("·ÿ›« ¬Œ—Ì‰ ﬁ”ÿ —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

z = 0
z = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
If z = 6 Then
  If Option1.Value = True Then
    Adodc1.Recordset.MoveFirst
    Do
      If (Adodc1.Recordset.Fields!id = List6.List(List6.ListIndex)) And (Adodc1.Recordset.Fields!rad = List1.List(List1.ListIndex)) Then
        Adodc1.Recordset.Delete
        Adodc1.Refresh
        Exit Do
      End If
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
    List1.RemoveItem (List1.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    List8.RemoveItem (List8.ListIndex)
    List9.RemoveItem (List9.ListIndex)
    Adodc1.Refresh
  End If
  
  If Option2.Value = True Then
    Adodc2.Recordset.MoveFirst
    Do
      If (Adodc2.Recordset.Fields!id = List6.List(List6.ListIndex)) And (Adodc2.Recordset.Fields!rad = List1.List(List1.ListIndex)) Then
        Adodc2.Recordset.Delete
        Adodc2.Refresh
        Exit Do
      End If
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
    List1.RemoveItem (List1.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    List8.RemoveItem (List8.ListIndex)
    List9.RemoveItem (List9.ListIndex)
    Adodc2.Refresh
  End If
  
  If Option3.Value = True Then
    Adodc3.Recordset.MoveFirst
    Do
      If (Adodc3.Recordset.Fields!id = List6.List(List6.ListIndex)) And (Adodc3.Recordset.Fields!rad = List1.List(List1.ListIndex)) Then
        Adodc3.Recordset.Delete
        Adodc3.Refresh
        Exit Do
      End If
      Adodc3.Recordset.MoveNext
    Loop Until Adodc3.Recordset.EOF = True
    List1.RemoveItem (List1.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    List8.RemoveItem (List8.ListIndex)
    List9.RemoveItem (List9.ListIndex)
    Adodc3.Refresh
  End If
  
  If Option1.Value = True Then q1 = 1
  If Option2.Value = True Then q1 = 2
  If Option3.Value = True Then q1 = 3
  Call List6_Click
  If q1 = 1 Then Option1.Value = True
  If q1 = 2 Then Option2.Value = True
  If q1 = 3 Then Option3.Value = True
  Text15.SetFocus
  SendKeys ("{enter}")
End If
akhar:
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List8.ListIndex = List1.ListIndex
List9.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List8.ListIndex = List2.ListIndex
List9.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List8.ListIndex = List3.ListIndex
List9.ListIndex = List3.ListIndex
End Sub

Private Sub List8_Click()
List1.ListIndex = List8.ListIndex
List2.ListIndex = List8.ListIndex
List3.ListIndex = List8.ListIndex
List9.ListIndex = List8.ListIndex
End Sub

Private Sub List9_Click()
List1.ListIndex = List9.ListIndex
List2.ListIndex = List9.ListIndex
List8.ListIndex = List9.ListIndex
List3.ListIndex = List9.ListIndex
End Sub

Private Sub List4_Click()
For q = 0 To List5.ListCount - 1
   If List5.List(q) = List4.List(List4.ListIndex) Then
     List5.ListIndex = q
     Exit For
   End If
Next q
End Sub

Private Sub List5_Click()
List6.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
Dim id(1000) As Integer, na(1000), da(1000), q1(1000), q2(1000), idt, nat, dat, q1t, q2t, count As String
List5.ListIndex = List6.ListIndex
Combo1.Clear
If (Option1.Value = True) And (Form7.Adodc1.Recordset.RecordCount > 0) Then
  Form7.Adodc1.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
  Label4(0).Caption = Amin.moneyaminjoda(Form7.Adodc1.Recordset.Fields!moneyvam)
  Label4(1).Caption = Form7.Adodc1.Recordset.Fields!numberagsat
  Label4(2).Caption = Form7.Adodc1.Recordset.Fields!Date
  Label4(3).Caption = Amin.moneyaminjoda(Form7.Adodc1.Recordset.Fields!karmozd)
  Label4(6).Caption = Form7.Adodc1.Recordset.Fields!id1
  Label4(7).Caption = List5.List(List5.ListIndex)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc1.Recordset.Fields!moneyg1)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc1.Recordset.Fields!moneyg2)
  If Adodc1.Recordset.RecordCount > 0 Then
    List1.Clear
    List2.Clear
    List3.Clear
    List8.Clear
    List9.Clear
    Adodc1.Recordset.MoveFirst
    Do
      If Adodc1.Recordset.Fields!id = List6.List(List6.ListIndex) Then
        List1.AddItem Adodc1.Recordset.Fields!rad
        List2.AddItem Adodc1.Recordset.Fields!Date
        List3.AddItem Amin.moneyaminjoda(Adodc1.Recordset.Fields!Money)
        List8.AddItem Adodc1.Recordset.Fields!saragsat
        List9.AddItem Adodc1.Recordset.Fields!emteyaz
      End If
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
  End If
  Label4(4).Caption = 0
  For q = 0 To List3.ListCount - 1
     Label4(4).Caption = Val(Label4(4).Caption) + Amin.moneyaminnojoda(List3.List(q))
  Next q
  Label4(5).Caption = (Amin.moneyaminnojoda(Label4(0).Caption)) - (Label4(4).Caption)
  Label4(4).Caption = Amin.moneyaminjoda(Label4(4).Caption)
  Label4(5).Caption = Amin.moneyaminjoda(Label4(5).Caption)
  KewlButtons2.Enabled = True
End If

If (Option2.Value = True) And (Form7.Adodc2.Recordset.RecordCount > 0) Then
  Form7.Adodc2.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
  Label4(0).Caption = Amin.moneyaminjoda(Form7.Adodc2.Recordset.Fields!moneyvam)
  Label4(1).Caption = Form7.Adodc2.Recordset.Fields!numberagsat
  Label4(2).Caption = Form7.Adodc2.Recordset.Fields!Date
  Label4(3).Caption = Amin.moneyaminjoda(Form7.Adodc2.Recordset.Fields!karmozd)
  Label4(6).Caption = Form7.Adodc2.Recordset.Fields!id1
  Label4(7).Caption = List5.List(List5.ListIndex)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc2.Recordset.Fields!moneyg1)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc2.Recordset.Fields!moneyg2)
  If Adodc2.Recordset.RecordCount > 0 Then
    List1.Clear
    List2.Clear
    List3.Clear
    List8.Clear
    List9.Clear
    Adodc2.Recordset.MoveFirst
    Do
      If Adodc2.Recordset.Fields!id = List6.List(List6.ListIndex) Then
        List1.AddItem Adodc2.Recordset.Fields!rad
        List2.AddItem Adodc2.Recordset.Fields!Date
        List3.AddItem Amin.moneyaminjoda(Adodc2.Recordset.Fields!Money)
        List8.AddItem Adodc2.Recordset.Fields!saragsat
        List9.AddItem Adodc2.Recordset.Fields!emteyaz
      End If
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  Label4(4).Caption = 0
  For q = 0 To List3.ListCount - 1
     Label4(4).Caption = Val(Label4(4).Caption) + Amin.moneyaminnojoda(List3.List(q))
  Next q
  Label4(5).Caption = (Amin.moneyaminnojoda(Label4(0).Caption)) - (Label4(4).Caption)
  Label4(4).Caption = Amin.moneyaminjoda(Label4(4).Caption)
  Label4(5).Caption = Amin.moneyaminjoda(Label4(5).Caption)
  KewlButtons2.Enabled = True
End If

If (Option3.Value = True) And (Form7.Adodc3.Recordset.RecordCount > 0) Then
  Form7.Adodc3.Recordset.Find "id='" + List6.List(List6.ListIndex) + "'", , adSearchForward, 1
  Label4(0).Caption = Amin.moneyaminjoda(Form7.Adodc3.Recordset.Fields!moneyvam)
  Label4(1).Caption = Form7.Adodc3.Recordset.Fields!numberagsat
  Label4(2).Caption = Form7.Adodc3.Recordset.Fields!Date
  Label4(3).Caption = Amin.moneyaminjoda(Form7.Adodc3.Recordset.Fields!karmozd)
  Label4(6).Caption = Form7.Adodc3.Recordset.Fields!id1
  Label4(7).Caption = List5.List(List5.ListIndex)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc3.Recordset.Fields!moneyg1)
  Combo1.AddItem Amin.moneyaminjoda(Form7.Adodc3.Recordset.Fields!moneyg2)
  If Adodc3.Recordset.RecordCount > 0 Then
    List1.Clear
    List2.Clear
    List3.Clear
    List8.Clear
    List9.Clear
    Adodc3.Recordset.MoveFirst
    Do
      If Adodc3.Recordset.Fields!id = List6.List(List6.ListIndex) Then
        List1.AddItem Adodc3.Recordset.Fields!rad
        List2.AddItem Adodc3.Recordset.Fields!Date
        List3.AddItem Amin.moneyaminjoda(Adodc3.Recordset.Fields!Money)
        List8.AddItem Adodc3.Recordset.Fields!saragsat
        List9.AddItem Adodc3.Recordset.Fields!emteyaz
      End If
      Adodc3.Recordset.MoveNext
    Loop Until Adodc3.Recordset.EOF = True
  End If
  Label4(4).Caption = 0
  For q = 0 To List3.ListCount - 1
     Label4(4).Caption = Val(Label4(4).Caption) + Amin.moneyaminnojoda(List3.List(q))
  Next q
  Label4(5).Caption = (Amin.moneyaminnojoda(Label4(0).Caption)) - (Label4(4).Caption)
  Label4(4).Caption = Amin.moneyaminjoda(Label4(4).Caption)
  Label4(5).Caption = Amin.moneyaminjoda(Label4(5).Caption)
  KewlButtons2.Enabled = True
End If

'„— » ”«“Ì
  For intq = 0 To List1.ListCount - 1
     id(intq) = List1.List(intq)
     na(intq) = List2.List(intq)
     da(intq) = List3.List(intq)
     q1(intq) = List8.List(intq)
     q2(intq) = List9.List(intq)
  Next intq
  count = List1.ListCount - 1
  For intq = 0 To count
     For intw = intq To count
        If id(intq) > id(intw) Then
          idt = id(intq)
          nat = na(intq)
          dat = da(intq)
          q1t = q1(intq)
          q2t = q2(intq)

          id(intq) = id(intw)
          na(intq) = na(intw)
          da(intq) = da(intw)
          q1(intq) = q1(intw)
          q2(intq) = q2(intw)

          id(intw) = idt
          na(intw) = nat
          da(intw) = dat
          q1(intw) = q1t
          q2(intw) = q2t
        End If
     Next intw
  Next intq
  List1.Clear
  List2.Clear
  List3.Clear
  List8.Clear
  List9.Clear

  For intq = 0 To count
     List1.AddItem id(intq)
     List2.AddItem na(intq)
     List3.AddItem da(intq)
     List8.AddItem q1(intq)
     List9.AddItem q2(intq)
  Next intq
Text1(1).Text = Form2.Label5.Caption
End Sub

Private Sub Option1_Click()
List6.Clear
List7.Clear
List5.Clear
List4.Clear
If Form7.Adodc1.Recordset.RecordCount > 0 Then
  Form7.Adodc1.Recordset.MoveFirst
  Do
    If Form7.Adodc1.Recordset.Fields!tasvie = "‰‘œÂ" Then
      List6.AddItem Form7.Adodc1.Recordset.Fields!id
      List7.AddItem Form7.Adodc1.Recordset.Fields!id1
    End If
    Form7.Adodc1.Recordset.MoveNext
  Loop Until Form7.Adodc1.Recordset.EOF = True
End If
For q = 0 To List7.ListCount - 1
  Form3.Adodc1.Recordset.Find "id='" & List7.List(q) & "'", , adSearchForward, 1
  List5.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
Next q
End Sub

Private Sub Option2_Click()
List6.Clear
List7.Clear
List5.Clear
List4.Clear
If Form7.Adodc2.Recordset.RecordCount > 0 Then
  Form7.Adodc2.Recordset.MoveFirst
  Do
    If Form7.Adodc2.Recordset.Fields!tasvie = "‰‘œÂ" Then
      List6.AddItem Form7.Adodc2.Recordset.Fields!id
      List7.AddItem Form7.Adodc2.Recordset.Fields!id1
    End If
    Form7.Adodc2.Recordset.MoveNext
  Loop Until Form7.Adodc2.Recordset.EOF = True
End If
For q = 0 To List7.ListCount - 1
  Form3.Adodc1.Recordset.Find "id='" & List7.List(q) & "'", , adSearchForward, 1
  List5.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
Next q
End Sub

Private Sub Option3_Click()
List6.Clear
List7.Clear
List5.Clear
List4.Clear
If Form7.Adodc3.Recordset.RecordCount > 0 Then
  Form7.Adodc3.Recordset.MoveFirst
  Do
    If Form7.Adodc3.Recordset.Fields!tasvie = "‰‘œÂ" Then
      List6.AddItem Form7.Adodc3.Recordset.Fields!id
      List7.AddItem Form7.Adodc3.Recordset.Fields!id1
    End If
    Form7.Adodc3.Recordset.MoveNext
  Loop Until Form7.Adodc3.Recordset.EOF = True
End If
For q = 0 To List7.ListCount - 1
  Form4.Adodc1.Recordset.Find "id='" & List7.List(q) & "'", , adSearchForward, 1
  List5.AddItem Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
Next q
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Select Case Index
    Case 1
      KewlButtons2.SetFocus

  End Select
End If
End Sub

Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15.Text)
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List6.ListCount - 1
     If Trim(Text15.Text) = List6.List(q) Then List6.ListIndex = q
   Next q
End If
End Sub

Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   List4.Clear
   For q = 0 To List5.ListCount - 1
       If InStr(List5.List(q), Trim(Text16.Text)) <> 0 Then
          List4.AddItem List5.List(q)
       End If
   Next q
End If
End Sub
