VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Õ”«» ⁄«œÌ"
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
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons KewlButtons7 
      Height          =   495
      Left            =   2160
      TabIndex        =   50
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "Form3.frx":10378
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   11
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   13
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6960
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   12
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6360
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   1
      Left            =   4200
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      ItemData        =   "Form3.frx":10394
      Left            =   2280
      List            =   "Form3.frx":103A1
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Combo2"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form3.frx":103B8
      Left            =   4560
      List            =   "Form3.frx":103C2
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   6
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   9
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   10
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   8
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   7
      Left            =   2280
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   5
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   4
      Left            =   2280
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   3
      Left            =   4920
      MaxLength       =   25
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   2
      Left            =   7920
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   7440
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ã” ÃÊ"
      Height          =   4335
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   2895
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
         TabIndex        =   24
         Top             =   3000
         Width           =   2655
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
         Height          =   1260
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
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
         Height          =   1260
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   855
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
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
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
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   135
         Left            =   120
         TabIndex        =   17
         Top             =   4080
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
         MICON           =   "Form3.frx":103D3
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
         TabIndex        =   18
         Top             =   2400
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
         MICON           =   "Form3.frx":103EF
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
         TabIndex        =   19
         Top             =   2400
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
         MICON           =   "Form3.frx":1040B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         TabIndex        =   27
         Top             =   360
         Width           =   855
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
         TabIndex        =   26
         Top             =   360
         Width           =   1695
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
         TabIndex        =   25
         Top             =   2640
         Width           =   2655
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   1440
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
      Connect         =   $"Form3.frx":10427
      OLEDBString     =   $"Form3.frx":105DB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "AccountAdi"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin KewlButtonz.KewlButtons KewlButtons5 
      Height          =   375
      Left            =   10440
      TabIndex        =   29
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«›  «Õ Õ”«» ÃœÌœ"
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
      MICON           =   "Form3.frx":1078F
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
      Left            =   10440
      TabIndex        =   30
      Top             =   8160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› Õ”«»"
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
      MICON           =   "Form3.frx":107AB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons8 
      Height          =   375
      Left            =   10440
      TabIndex        =   31
      Top             =   9600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
      MICON           =   "Form3.frx":107C7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons9 
      Height          =   375
      Left            =   10440
      TabIndex        =   51
      Top             =   8520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ç«Å «ÿ·«⁄«  Õ”«»"
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
      MICON           =   "Form3.frx":107E3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons10 
      Height          =   375
      Left            =   10440
      TabIndex        =   54
      Top             =   8880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "’œÊ— œ› —çÂ Å” «‰œ«“"
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
      MICON           =   "Form3.frx":107FF
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
      Height          =   375
      Left            =   10440
      TabIndex        =   55
      Top             =   9240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ç«Å „‘Œ’«  Å—Ê‰œÂ"
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
      MICON           =   "Form3.frx":1081B
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
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ ò· «⁄÷« :"
      Height          =   495
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   8760
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰„Ê‰Â «„÷«¡"
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» ⁄«œÌ"
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
      TabIndex        =   48
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ „Õ· ò«—"
      Height          =   495
      Index           =   15
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰‘«‰Ì „Õ· ò«—"
      Height          =   495
      Index           =   14
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰‘«‰Ì „‰“·"
      Height          =   495
      Index           =   13
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «›  «Õ Õ”«»"
      Height          =   495
      Index           =   12
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ »”ÌÃ"
      Height          =   495
      Index           =   11
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ê÷⁄Ì   «Â·"
      Height          =   495
      Index           =   10
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ „·Ì"
      Height          =   495
      Index           =   9
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ Â„—«Â"
      Height          =   495
      Index           =   8
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ „‰“·"
      Height          =   495
      Index           =   7
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Õ’Ì·« "
      Height          =   495
      Index           =   6
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  Ê·œ"
      Height          =   495
      Index           =   5
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ‘‰«”‰«„Â"
      Height          =   495
      Index           =   4
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Åœ—"
      Height          =   495
      Index           =   3
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   495
      Index           =   2
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      Height          =   495
      Index           =   1
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Õ”«»"
      Height          =   495
      Index           =   0
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim o As Boolean
Private Sub Amin_1()
For q = 0 To 13
  Text1(q).Text = ""
Next q
Combo1.Text = ""
Combo2.Text = ""
Label6.Caption = ""
Image1.Picture = LoadPicture("")
Image2.Picture = LoadPicture("")
List1.Clear
List2.Clear
List3.Clear
If Adodc1.Recordset.RecordCount <> 0 Then
  Adodc1.Recordset.MoveFirst
  Do
     List1.AddItem Adodc1.Recordset.Fields!id
     List2.AddItem Adodc1.Recordset.Fields!Name + " " + Adodc1.Recordset.Fields!family
     Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo2.SetFocus
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1(9).SetFocus
End Sub

Private Sub Form_Load()
Call Amin_1
Label4.Caption = 0
Label7.Caption = List2.ListCount
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

Private Sub KewlButtons10_Click()
If List1.ListIndex <> -1 Then
  Form19.Show
Else
  z = MsgBox("·ÿ›« ‰«„ ‘Œ’Ì —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
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
If List1.ListIndex <> -1 Then
  Form22.Show
Else
  z = MsgBox("·ÿ›« ‰«„ ‘Œ’Ì —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons5_Click()
For q = 0 To 13
  Text1(q).Text = ""
Next q
Combo1.Text = ""
Combo2.Text = ""
Text1(0).SetFocus
Label4.Caption = 1
KewlButtons7.Enabled = True
Text1(1).Text = Form2.Label5.Caption
Label6.Caption = ""
End Sub

Private Sub KewlButtons6_Click()
If List1.ListIndex <> -1 Then
  z = 0
  z = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
  If z = 6 Then
    Adodc1.Recordset.MoveFirst
    Do
      If Adodc1.Recordset.Fields!id = List1.List(List1.ListIndex) Then
        Adodc1.Recordset.Fields!Delete = "‘œÂ"
        Adodc1.Recordset.Update
        z = MsgBox("Õ”«» „Ê—œ ‰Ÿ— «“ ”Ì” „ Å«ò ‘œ", vbInformation + vbMsgBoxRight, "")
        Exit Do
      End If
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
  End If
Else
  z = MsgBox("·ÿ›« ‰«„ ‘Œ’Ì —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub KewlButtons7_Click()
o = False
For q = 0 To 13
  If Text1(q) = "" Then o = True
Next q
If o = True Then
  z = MsgBox("·ÿ›« ›Ì·œ Â«Ì Œ«·Ì —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Len(Text1(7).Text) <> 10) Or (Len(Text1(1).Text) <> 10) Then
  z = MsgBox(" «—ÌŒ Â« —« »Â ‘ò· ’ÕÌÕ Ê«—œ ‰ò—œÂ «Ìœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If Label4.Caption = 1 Then
  If Adodc1.Recordset.RecordCount <> 0 Then
    o = False
    Adodc1.Recordset.MoveFirst
    Do
       If Adodc1.Recordset.Fields!id = Text1(0).Text Then o = True: Exit Do
       Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
    If o = True Then
      z = MsgBox("«Ì‰ Õ”«» ﬁ»·« À»  ‘œÂ «” ", vbCritical + vbMsgBoxRight, "")
    Else
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc1.Recordset.Fields!Name = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!family = Trim(Text1(3).Text)
      Adodc1.Recordset.Fields!fathername = Trim(Text1(4).Text)
      Adodc1.Recordset.Fields!shsh = Trim(Text1(5).Text)
      Adodc1.Recordset.Fields!birthdate = Trim(Text1(7).Text)
      Adodc1.Recordset.Fields!edate = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!basig = Trim(Combo2.Text)
      Adodc1.Recordset.Fields!madrak = Trim(Text1(8).Text)
      Adodc1.Recordset.Fields!addresshome = Trim(Text1(12).Text)
      Adodc1.Recordset.Fields!mobile = Trim(Text1(9).Text)
      Adodc1.Recordset.Fields!phonehome = Trim(Text1(10).Text)
      Adodc1.Recordset.Fields!addresswork = Trim(Text1(13).Text)
      Adodc1.Recordset.Fields!phonework = Trim(Text1(11).Text)
      Adodc1.Recordset.Fields!mellicode = Trim(Text1(6).Text)
      Adodc1.Recordset.Fields!taahol = Trim(Combo1.Text)
      Adodc1.Recordset.Fields!Delete = "‰‘œÂ"
      Adodc1.Recordset.Fields!Money = 0
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
    End If
  Else
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc1.Recordset.Fields!Name = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!family = Trim(Text1(3).Text)
      Adodc1.Recordset.Fields!fathername = Trim(Text1(4).Text)
      Adodc1.Recordset.Fields!shsh = Trim(Text1(5).Text)
      Adodc1.Recordset.Fields!birthdate = Trim(Text1(7).Text)
      Adodc1.Recordset.Fields!edate = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!basig = Trim(Combo2.Text)
      Adodc1.Recordset.Fields!madrak = Trim(Text1(8).Text)
      Adodc1.Recordset.Fields!addresshome = Trim(Text1(12).Text)
      Adodc1.Recordset.Fields!mobile = Trim(Text1(9).Text)
      Adodc1.Recordset.Fields!phonehome = Trim(Text1(10).Text)
      Adodc1.Recordset.Fields!addresswork = Trim(Text1(13).Text)
      Adodc1.Recordset.Fields!phonework = Trim(Text1(11).Text)
      Adodc1.Recordset.Fields!mellicode = Trim(Text1(6).Text)
      Adodc1.Recordset.Fields!taahol = Trim(Combo1.Text)
      Adodc1.Recordset.Fields!Delete = "‰‘œÂ"
      Adodc1.Recordset.Fields!Money = 0
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
  End If
End If
If Label4.Caption = 2 Then
      Adodc1.Recordset.Fields!Name = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!family = Trim(Text1(3).Text)
      Adodc1.Recordset.Fields!fathername = Trim(Text1(4).Text)
      Adodc1.Recordset.Fields!shsh = Trim(Text1(5).Text)
      Adodc1.Recordset.Fields!birthdate = Trim(Text1(7).Text)
      Adodc1.Recordset.Fields!edate = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!basig = Trim(Combo2.Text)
      Adodc1.Recordset.Fields!madrak = Trim(Text1(8).Text)
      Adodc1.Recordset.Fields!addresshome = Trim(Text1(12).Text)
      Adodc1.Recordset.Fields!mobile = Trim(Text1(9).Text)
      Adodc1.Recordset.Fields!phonehome = Trim(Text1(10).Text)
      Adodc1.Recordset.Fields!addresswork = Trim(Text1(13).Text)
      Adodc1.Recordset.Fields!phonework = Trim(Text1(11).Text)
      Adodc1.Recordset.Fields!mellicode = Trim(Text1(6).Text)
      Adodc1.Recordset.Fields!taahol = Trim(Combo1.Text)
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „  €ÌÌ— ÅÌœ« ò—œ", vbMsgBoxRight + vbInformation, "")
End If
Call Amin_1
Text1(0).SetFocus
akhar:
KewlButtons7.Enabled = False
End Sub

Private Sub KewlButtons8_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons9_Click()
If List1.ListIndex <> -1 Then


Else
  z = MsgBox("·ÿ›« ‰«„ ‘Œ’Ì —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
  Adodc1.Recordset.Find "id='" + List1.List(List2.ListIndex) + "'", , adSearchForward, 1
  Text1(0).Text = Adodc1.Recordset.Fields!id
  Text1(2).Text = Adodc1.Recordset.Fields!Name
  Text1(3).Text = Adodc1.Recordset.Fields!family
  Text1(4).Text = Adodc1.Recordset.Fields!fathername
  Text1(5).Text = Adodc1.Recordset.Fields!shsh
  Text1(7).Text = Adodc1.Recordset.Fields!birthdate
  Text1(1).Text = Adodc1.Recordset.Fields!edate
  Combo2.Text = Adodc1.Recordset.Fields!basig
  Text1(8).Text = Adodc1.Recordset.Fields!madrak
  Text1(12).Text = Adodc1.Recordset.Fields!addresshome
  Text1(9).Text = Adodc1.Recordset.Fields!mobile
  Text1(10).Text = Adodc1.Recordset.Fields!phonehome
  Text1(13).Text = Adodc1.Recordset.Fields!addresswork
  Text1(11).Text = Adodc1.Recordset.Fields!phonework
  Text1(6).Text = Adodc1.Recordset.Fields!mellicode
  If Adodc1.Recordset.Fields!Delete = "‰‘œÂ" Then
    Label6.Caption = "«Ì‰ Õ”«» Õ–› ‰‘œÂ «” "
  Else
    Label6.Caption = "«Ì‰ Õ”«» Õ–› ‘œÂ «” "
  End If
  
  Combo1.Text = Adodc1.Recordset.Fields!taahol
  Text1(0).SetFocus
  Label4.Caption = 2
  KewlButtons7.Enabled = True
  If basItemExist.ItemExist(App.Path + "\DATA BASE INFORMATION\Image\" + Text1(0).Text + ".bmp") = True Then
    Image1.Picture = LoadPicture(App.Path + "\DATA BASE INFORMATION\Image\" + Text1(0).Text + ".bmp")
  Else
    Image1.Picture = LoadPicture("")
  End If
  
  If basItemExist.ItemExist(App.Path + "\DATA BASE INFORMATION\Emza\" + Text1(0).Text + ".bmp") = True Then
    Image2.Picture = LoadPicture(App.Path + "\DATA BASE INFORMATION\Emza\" + Text1(0).Text + ".bmp")
  Else
    Image2.Picture = LoadPicture("")
  End If
End Sub

Private Sub List3_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
  Case 0
    Text1(1).SetFocus
  
  Case 1
    Text1(2).SetFocus
    
  Case 2
    Text1(3).SetFocus
    
  Case 3
    Text1(4).SetFocus
    
  Case 4
    Text1(5).SetFocus
  
  Case 5
    Text1(6).SetFocus
  
  Case 6
    Text1(7).SetFocus
  
  Case 7
    Text1(8).SetFocus
  
  Case 8
    Combo1.SetFocus
  
  Case 9
    Text1(10).SetFocus
  
  Case 10
    Text1(11).SetFocus
  
  Case 11
    Text1(12).SetFocus
    
  Case 12
    Text1(13).SetFocus
    
  Case 13
    If KewlButtons7.Enabled = True Then KewlButtons7.SetFocus
    
End Select
End If
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
       If InStr(1, List2.List(q), Trim(Text16.Text)) <> 0 Then
          List3.AddItem List2.List(q)
       End If
   Next q
End If
End Sub

