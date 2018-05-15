VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Õ”«» ÊÌéÂ"
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
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List4 
      Height          =   2130
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   8160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin KewlButtonz.KewlButtons KewlButtons11 
      Height          =   615
      Left            =   3000
      TabIndex        =   32
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "»Œ‘ „ÊÃÊœÌ"
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
      MICON           =   "Form4.frx":10378
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ã” ÃÊ"
      Height          =   5295
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3360
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
         TabIndex        =   1
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
         TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         Height          =   1560
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3480
         Width           =   2655
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   135
         Left            =   120
         TabIndex        =   25
         Top             =   5040
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
         MICON           =   "Form4.frx":10394
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
         Left            =   1920
         TabIndex        =   26
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
         MICON           =   "Form4.frx":103B0
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
         Height          =   135
         Left            =   120
         TabIndex        =   27
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
         MICON           =   "Form4.frx":103CC
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Height          =   3735
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4320
      Width           =   6975
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   4080
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   4080
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   2
         Left            =   4080
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin KewlButtonz.KewlButtons KewlButtons7 
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2760
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
         MICON           =   "Form4.frx":103E8
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
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Ã” ÃÊ"
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
         MICON           =   "Form4.frx":10404
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   495
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Õ”«» ÊÌéÂ"
         Height          =   495
         Index           =   0
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Õ”«» ⁄«œÌ"
         Height          =   495
         Index           =   1
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «›  «Õ Õ”«»"
         Height          =   495
         Index           =   2
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
         Height          =   495
         Index           =   3
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   495
         Index           =   4
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   495
         Index           =   5
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2040
         Width           =   2175
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons8 
      Height          =   495
      Left            =   11640
      TabIndex        =   12
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
      MICON           =   "Form4.frx":10420
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
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
      Connect         =   $"Form4.frx":1043C
      OLEDBString     =   $"Form4.frx":105F0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "Accountvig"
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
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      MICON           =   "Form4.frx":107A4
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
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
      MICON           =   "Form4.frx":107C0
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
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      MICON           =   "Form4.frx":107DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» ÊÌéÂ"
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
      TabIndex        =   15
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
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
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Boolean

Private Sub Form_Activate()
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

Private Sub Form_Load()
For q = 0 To 2
  Text1(q).Text = ""
Next q
Label2(4).Caption = "-"
Label2(5).Caption = "-"
Label4.Caption = 0
KewlButtons7.Enabled = False
Label5.Caption = ""
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

Private Sub KewlButtons10_Click()
Form14.Show
End Sub

Private Sub KewlButtons11_Click()
Form5.Show
End Sub

Private Sub KewlButtons2_Click()
If List1.ListIndex <> -1 Then
  Form3.Adodc1.Recordset.Find "id='" + Text1(1).Text + "'", , adSearchForward, 1
  List4.Clear
  List4.AddItem Form3.Adodc1.Recordset.Fields!fathername
  List4.AddItem Form3.Adodc1.Recordset.Fields!mellicode
  List4.AddItem Form3.Adodc1.Recordset.Fields!addresshome
  List4.AddItem Form3.Adodc1.Recordset.Fields!phonehome
  List4.AddItem Form3.Adodc1.Recordset.Fields!mobile
  Form20.Show
Else
  z = MsgBox("·ÿ›« ‰«„ ‘Œ’Ì —« «‰ Œ«» ”Å” «Ì‰ ⁄„·Ì«  —« «‰Ã«„ œÂÌœ", vbCritical + vbMsgBoxRight, "")
End If
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

Private Sub KewlButtons5_Click()
For q = 0 To 2
  Text1(q).Text = ""
Next q
Label2(4).Caption = "-"
Label2(5).Caption = "-"
Label5.Caption = ""
Text1(0).SetFocus
Label4.Caption = 1
KewlButtons7.Enabled = True
Text1(2).Text = Form2.Label5.Caption
End Sub

Private Sub KewlButtons6_Click()
If List2.ListIndex <> -1 Then
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

If (Len(Text1(2).Text) <> 10) Then
  z = MsgBox(" «—ÌŒ Â« —« »Â ‘ò· ’ÕÌÕ Ê«—œ ‰ò—œÂ «Ìœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If Label4.Caption = 1 Then
  If Adodc1.Recordset.RecordCount <> 0 Then
    o = False
    Adodc1.Recordset.MoveFirst
    Do
       If Adodc1.Recordset.Fields!id = Trim(Text1(0).Text) Then o = True: Exit Do
       Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
    If o = True Then
      z = MsgBox("«Ì‰ Õ”«» ﬁ»·« À»  ‘œÂ «” ", vbCritical + vbMsgBoxRight, "")
    Else
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc1.Recordset.Fields!idadi = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!edate = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!Name = Label2(4).Caption
      Adodc1.Recordset.Fields!family = Label2(5).Caption
      Adodc1.Recordset.Fields!Money = 0
      Adodc1.Recordset.Fields!Delete = "‰‘œÂ"
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
      Call Form_Activate
    End If
  Else
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc1.Recordset.Fields!idadi = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!edate = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!Name = Label2(4).Caption
      Adodc1.Recordset.Fields!family = Label2(5).Caption
      Adodc1.Recordset.Fields!Money = 0
      Adodc1.Recordset.Fields!Delete = "‰‘œÂ"
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
      Call Form_Activate
  End If
End If
If Label4.Caption = 2 Then
  a = List1.List(List1.ListIndex)
  
  Adodc1.Recordset.Find "id='" + a + "'", , adSearchForward, 1
  Adodc1.Recordset.Fields!idadi = Trim(Text1(1).Text)
  Adodc1.Recordset.Fields!edate = Trim(Text1(2).Text)
  Adodc1.Recordset.Fields!Name = Label2(4).Caption
  Adodc1.Recordset.Fields!family = Label2(5).Caption
  Adodc1.Recordset.Fields!user = Form2.Label2.Caption
  Adodc1.Recordset.Update
  z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „  €ÌÌ— ÅÌœ« ò—œ", vbMsgBoxRight + vbInformation, "")
  Call Form_Activate
End If

akhar:
End Sub

Private Sub KewlButtons8_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons9_Click()
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

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
  a = List1.List(List1.ListIndex)
  Adodc1.Recordset.Find "id='" + a + "'", , adSearchForward, 1
  Text1(0).Text = Adodc1.Recordset.Fields!id
  Label2(4).Caption = Adodc1.Recordset.Fields!Name
  Label2(5).Caption = Adodc1.Recordset.Fields!family
  Text1(1).Text = Adodc1.Recordset.Fields!idadi
  Text1(2).Text = Adodc1.Recordset.Fields!edate
  If Adodc1.Recordset.Fields!Delete = "‰‘œÂ" Then
    Label5.Caption = "«Ì‰ Õ”«» Õ–› ‰‘œÂ «” "
  Else
    Label5.Caption = "«Ì‰ Õ”«» Õ–› ‘œÂ «” "
  End If
  Text1(0).SetFocus
  Label4.Caption = 2
  KewlButtons7.Enabled = True
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
    p = False
    
    Form3.Adodc1.Recordset.MoveFirst
    Do
      If Form3.Adodc1.Recordset.Fields!id = Trim(Text1(1).Text) Then p = True: Exit Do
      Form3.Adodc1.Recordset.MoveNext
    Loop Until Form3.Adodc1.Recordset.EOF = True
    
    If p = False Then
      z = MsgBox("òœ Õ”«» ⁄«œÌ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      Label2(4).Caption = "-"
      Label2(5).Caption = "-"
    Else
      Label2(4).Caption = Form3.Adodc1.Recordset.Fields!Name
      Label2(5).Caption = Form3.Adodc1.Recordset.Fields!family
      Text1(2).SetFocus
    End If
    
  Case 2
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
       If InStr(List2.List(q), Trim(Text16.Text)) <> 0 Then
          List3.AddItem List2.List(q)
       End If
   Next q
End If
End Sub
