VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "À»  Ê«„"
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
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   9
      Left            =   2760
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8400
      Width           =   1575
   End
   Begin KewlButtonz.KewlButtons KewlButtons6 
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   9480
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form6.frx":10378
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
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      MICON           =   "Form6.frx":10394
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
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "ÊÌ—«Ì‘"
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
      MICON           =   "Form6.frx":103B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "ÃœÌœ"
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
      MICON           =   "Form6.frx":103CC
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
      Index           =   8
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   7
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   6
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7680
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2400
      Top             =   2040
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
      Connect         =   $"Form6.frx":103E8
      OLEDBString     =   $"Form6.frx":1059C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "vamVigt"
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
      Left            =   2400
      Top             =   1680
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
      Connect         =   $"Form6.frx":10750
      OLEDBString     =   $"Form6.frx":10904
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "vamazt"
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
      Left            =   2400
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   $"Form6.frx":10AB8
      OLEDBString     =   $"Form6.frx":10C6C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "vamadit"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "«‰ Œ«» ‰Ê⁄ Ê«„ "
      Height          =   975
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3240
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
         Width           =   735
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   495
      Left            =   11280
      TabIndex        =   19
      Top             =   9360
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "Form6.frx":10E20
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
      Height          =   4815
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4320
      Width           =   2655
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
         Height          =   3360
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   135
         Left            =   240
         TabIndex        =   28
         Top             =   4440
         Width           =   2295
         _ExtentX        =   4048
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
         MICON           =   "Form6.frx":10E3C
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
         Caption         =   "·Ì”  Ê«„ Â«Ì œ—ŒÊ«”  œ«œÂ ‘œÂ"
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   5
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   4
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   3
      Left            =   6600
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   2
      Left            =   6600
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   1
      Left            =   6600
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   6960
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ œ— ŒÊ«” "
      Height          =   495
      Index           =   9
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ò«—„“œ"
      Height          =   495
      Index           =   8
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ »ﬁÌÂ «ﬁ”«ÿ"
      Height          =   495
      Index           =   7
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ﬁ”ÿ «Ê·"
      Height          =   495
      Index           =   6
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      DataField       =   "id"
      DataSource      =   "Adodc3"
      Height          =   255
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «ﬁ”«ÿ"
      Height          =   495
      Index           =   5
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ò·Ì Ê«„"
      Height          =   495
      Index           =   4
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» ÷«„‰ œÊ„"
      Height          =   495
      Index           =   3
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» ÷«„‰ «Ê·"
      Height          =   495
      Index           =   2
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» Ê«„ êÌ—‰œÂ"
      Height          =   495
      Index           =   1
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Ê«„ "
      Height          =   495
      Index           =   0
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À»   Ê«„"
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
      TabIndex        =   20
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Boolean

Private Sub Form_Activate()
Call KewlButtons3_Click
KewlButtons6.Enabled = False
Label12.Caption = 0
Option1.Value = True
Call Option1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Shell "C:\WINDOWS\system32\calc.exe"
End Sub

Private Sub KewlButtons2_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons3_Click()
For q = 0 To 9
  Text1(q).Text = ""
Next q
Label7.Caption = "-"
Label9.Caption = "-"
Label11.Caption = "-"
Label12.Caption = 1
KewlButtons6.Enabled = True
Text1(0).SetFocus
Text1(9).Text = Form2.Label5.Caption
End Sub


Private Sub KewlButtons4_Click()
If List2.ListIndex = -1 Then
  z = MsgBox("·ÿ›« Ê«„ „Ê—œ ‰Ÿ— —« „‘Œ’ ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If
KewlButtons6.Enabled = True
Label12.Caption = 2
If (Option1.Value = True) And (Adodc1.Recordset.RecordCount > 0) Then
  Adodc1.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Text1(0).Text = Adodc1.Recordset.Fields!id
  Text1(1).Text = Adodc1.Recordset.Fields!id1
  Text1(2).Text = Adodc1.Recordset.Fields!idz1
  Text1(3).Text = Adodc1.Recordset.Fields!idz2
  Text1(4).Text = Amin.moneyaminjoda(Adodc1.Recordset.Fields!moneyvam)
  Text1(5).Text = Adodc1.Recordset.Fields!numberagsat
  Text1(6).Text = Amin.moneyaminjoda(Adodc1.Recordset.Fields!moneyg1)
  Text1(7).Text = Amin.moneyaminjoda(Adodc1.Recordset.Fields!moneyg2)
  Text1(8).Text = Adodc1.Recordset.Fields!karmozd
  Text1(9).Text = Adodc1.Recordset.Fields!datet
  Form3.Adodc1.Recordset.Find "id='" & Text1(1).Text & "'", , adSearchForward, 1
  Label7.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  Form3.Adodc1.Recordset.Find "id='" & Text1(2).Text & "'", , adSearchForward, 1
  Label9.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  Form3.Adodc1.Recordset.Find "id='" & Text1(3).Text & "'", , adSearchForward, 1
  Label11.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
End If

If (Option2.Value = True) And (Adodc2.Recordset.RecordCount > 0) Then
  Adodc2.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Text1(0).Text = Adodc2.Recordset.Fields!id
  Text1(1).Text = Adodc2.Recordset.Fields!id1
  Text1(2).Text = Adodc2.Recordset.Fields!idz1
  Text1(3).Text = Adodc2.Recordset.Fields!idz2
  Text1(4).Text = Adodc2.Recordset.Fields!moneyvam
  Text1(5).Text = Adodc2.Recordset.Fields!numberagsat
  Text1(6).Text = Adodc2.Recordset.Fields!moneyg1
  Text1(7).Text = Adodc2.Recordset.Fields!moneyg2
  Text1(8).Text = Adodc2.Recordset.Fields!karmozd
  Text1(9).Text = Adodc2.Recordset.Fields!datet
  Form3.Adodc1.Recordset.Find "id='" & Text1(1).Text & "'", , adSearchForward, 1
  Label7.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  Form3.Adodc1.Recordset.Find "id='" & Text1(2).Text & "'", , adSearchForward, 1
  Label9.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  Form3.Adodc1.Recordset.Find "id='" & Text1(3).Text & "'", , adSearchForward, 1
  Label11.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
End If

If (Option3.Value = True) And (Adodc3.Recordset.RecordCount > 0) Then
  Adodc3.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Text1(0).Text = Adodc3.Recordset.Fields!id
  Text1(1).Text = Adodc3.Recordset.Fields!id1
  Text1(2).Text = Adodc3.Recordset.Fields!idz1
  Text1(3).Text = Adodc3.Recordset.Fields!idz2
  Text1(4).Text = Adodc3.Recordset.Fields!moneyvam
  Text1(5).Text = Adodc3.Recordset.Fields!numberagsat
  Text1(6).Text = Adodc3.Recordset.Fields!moneyg1
  Text1(7).Text = Adodc3.Recordset.Fields!moneyg2
  Text1(8).Text = Adodc3.Recordset.Fields!karmozd
  Text1(9).Text = Adodc3.Recordset.Fields!datet
  Form4.Adodc1.Recordset.Find "id='" & Text1(1).Text & "'", , adSearchForward, 1
  Label7.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
  Form4.Adodc1.Recordset.Find "id='" & Text1(2).Text & "'", , adSearchForward, 1
  Label9.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
  Form4.Adodc1.Recordset.Find "id='" & Text1(3).Text & "'", , adSearchForward, 1
  Label11.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
End If
akhar:
End Sub

Private Sub KewlButtons5_Click()
If List2.ListIndex = -1 Then
  z = MsgBox("·ÿ›« Ê«„ „Ê—œ ‰Ÿ— —« „‘Œ’ ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Option1.Value = True) And (Adodc1.Recordset.RecordCount > 0) Then
  Adodc1.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Adodc1.Recordset.Delete
  Adodc1.Refresh
End If

If (Option2.Value = True) And (Adodc2.Recordset.RecordCount > 0) Then
  Adodc2.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Adodc2.Recordset.Delete
  Adodc2.Refresh
End If

If (Option3.Value = True) And (Adodc3.Recordset.RecordCount > 0) Then
  Adodc3.Recordset.Find "id='" & List2.List(List2.ListIndex) & "'", , adSearchForward, 1
  Adodc3.Recordset.Delete
  Adodc3.Refresh
End If
List2.RemoveItem (List2.ListIndex)
For q = 0 To 9
  Text1(q).Text = ""
Next q
Label7.Caption = "-"
Label9.Caption = "-"
Label11.Caption = "-"
akhar:
End Sub

Private Sub KewlButtons6_Click()
p = False
For q = 0 To 9
  If Text1(q).Text = "" Then p = True
Next q
If p = True Then
  z = MsgBox("·ÿ›« ›Ì·œ Â«Ì Œ«·Ì —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If Label12.Caption = 1 Then
  If Option1.Value = True Then
    p = False
    If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do
         If Adodc1.Recordset.Fields!id = Trim(Text1(0).Text) Then p = True: Exit Do
         Adodc1.Recordset.MoveNext
      Loop Until Adodc1.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc1.Recordset.Fields!id1 = Trim(Text1(1).Text)
      Adodc1.Recordset.Fields!idz1 = Trim(Text1(2).Text)
      Adodc1.Recordset.Fields!idz2 = Trim(Text1(3).Text)
      Adodc1.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
      Adodc1.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
      Adodc1.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
      Adodc1.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
      If Text1(8).Text = 0 Then
        Adodc1.Recordset.Fields!karmozd = 0
      Else
        Adodc1.Recordset.Fields!karmozd = ((Text1(8).Text * 2) * (Amin.moneyaminnojoda(Text1(4).Text) / 200))
      End If
      Adodc1.Recordset.Fields!datet = Text1(9).Text
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
      List2.AddItem Trim(Text1(0).Text)
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ  ò—«—Ì „Ì »«‘œ", vbMsgBoxRight + vbCritical, "")
      Text1(0).SetFocus
    End If
  End If
  
  If Option2.Value = True Then
    p = False
    If Adodc2.Recordset.RecordCount > 0 Then
      Adodc2.Recordset.MoveFirst
      Do
         If Adodc2.Recordset.Fields!id = Trim(Text1(0).Text) Then p = True: Exit Do
         Adodc2.Recordset.MoveNext
      Loop Until Adodc2.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc2.Recordset.AddNew
      Adodc2.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc2.Recordset.Fields!id1 = Trim(Text1(1).Text)
      Adodc2.Recordset.Fields!idz1 = Trim(Text1(2).Text)
      Adodc2.Recordset.Fields!idz2 = Trim(Text1(3).Text)
      Adodc2.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
      Adodc2.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
      Adodc2.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
      Adodc2.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
      If Text1(8).Text = 0 Then
        Adodc2.Recordset.Fields!karmozd = 0
      Else
        Adodc2.Recordset.Fields!karmozd = ((Text1(8).Text * 2) * (Amin.moneyaminnojoda(Text1(4).Text) / 200))
      End If
      Adodc2.Recordset.Fields!datet = Trim(Text1(9).Text)
      Adodc2.Recordset.Fields!user = Form2.Label2.Caption
      Adodc2.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
      List2.AddItem Trim(Text1(0).Text)
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ  ò—«—Ì „Ì »«‘œ", vbMsgBoxRight + vbCritical, "")
      Text1(1).SetFocus
    End If
  End If
  
  If Option3.Value = True Then
    p = False
    If Adodc3.Recordset.RecordCount > 0 Then
      Adodc3.Recordset.MoveFirst
      Do
         If Adodc3.Recordset.Fields!id = Text1(0).Text Then p = True: Exit Do
         Adodc3.Recordset.MoveNext
      Loop Until Adodc3.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!id = Trim(Text1(0).Text)
      Adodc3.Recordset.Fields!id1 = Trim(Text1(1).Text)
      Adodc3.Recordset.Fields!idz1 = Trim(Text1(2).Text)
      Adodc3.Recordset.Fields!idz2 = Trim(Text1(3).Text)
      Adodc3.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
      Adodc3.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
      Adodc3.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
      Adodc3.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
      If Text1(8).Text = 0 Then
        Adodc3.Recordset.Fields!karmozd = 0
      Else
        Adodc3.Recordset.Fields!karmozd = ((Text1(8).Text * 2) * (Amin.moneyaminnojoda(Text1(4).Text) / 200))
      End If
      Adodc3.Recordset.Fields!datet = Trim(Text1(9).Text)
      Adodc3.Recordset.Fields!user = Form2.Label2.Caption
      Adodc3.Recordset.Update
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
      List2.AddItem Trim(Text1(0).Text)
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ  ò—«—Ì „Ì »«‘œ", vbMsgBoxRight + vbCritical, "")
      Text1(2).SetFocus
    End If
  End If
End If
'
If Label12.Caption = 2 Then
  If Option1.Value = True Then
    Adodc1.Recordset.Find "id='" & Trim(List2.List(List2.ListIndex)) & "'", , adSearchForward, 1
    Adodc1.Recordset.Fields!id = Trim(Text1(0).Text)
    Adodc1.Recordset.Fields!id1 = Trim(Text1(1).Text)
    Adodc1.Recordset.Fields!idz1 = Trim(Text1(2).Text)
    Adodc1.Recordset.Fields!idz2 = Trim(Text1(3).Text)
    Adodc1.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
    Adodc1.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
    Adodc1.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
    Adodc1.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
    Adodc1.Recordset.Fields!karmozd = Trim(Text1(8).Text)
    Adodc1.Recordset.Fields!datet = Text1(9).Text
    Adodc1.Recordset.Fields!user = Form2.Label2.Caption
    Adodc1.Recordset.Update
    z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „  €ÌÌ— œ«œÂ ‘œ", vbMsgBoxRight + vbInformation, "")
    
  End If
  
  If Option2.Value = True Then
    Adodc2.Recordset.Find "id='" & Trim(List2.List(List2.ListIndex)) & "'", , adSearchForward, 1
    Adodc2.Recordset.Fields!id = Trim(Text1(0).Text)
    Adodc2.Recordset.Fields!id1 = Trim(Text1(1).Text)
    Adodc2.Recordset.Fields!idz1 = Trim(Text1(2).Text)
    Adodc2.Recordset.Fields!idz2 = Trim(Text1(3).Text)
    Adodc2.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
    Adodc2.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
    Adodc2.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
    Adodc2.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
    Adodc2.Recordset.Fields!karmozd = Trim(Text1(8).Text)
    Adodc2.Recordset.Fields!datet = Trim(Text1(9).Text)
    Adodc2.Recordset.Fields!user = Form2.Label2.Caption
    Adodc2.Recordset.Update
    z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „  €ÌÌ— œ«œÂ ‘œ", vbMsgBoxRight + vbInformation, "")
  End If
  
  If Option3.Value = True Then
    Adodc3.Recordset.Find "id='" & Trim(List2.List(List2.ListIndex)) & "'", , adSearchForward, 1
    Adodc3.Recordset.Fields!id = Trim(Text1(0).Text)
    Adodc3.Recordset.Fields!id1 = Trim(Text1(1).Text)
    Adodc3.Recordset.Fields!idz1 = Trim(Text1(2).Text)
    Adodc3.Recordset.Fields!idz2 = Trim(Text1(3).Text)
    Adodc3.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Trim(Text1(4).Text))
    Adodc3.Recordset.Fields!numberagsat = Trim(Text1(5).Text)
    Adodc3.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Trim(Text1(6).Text))
    Adodc3.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Trim(Text1(7).Text))
    Adodc3.Recordset.Fields!karmozd = Trim(Text1(8).Text)
    Adodc3.Recordset.Fields!datet = Trim(Text1(9).Text)
    Adodc3.Recordset.Fields!user = Form2.Label2.Caption
    Adodc3.Recordset.Update
    z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „  €ÌÌ— œ«œÂ ‘œ", vbMsgBoxRight + vbInformation, "")
  End If
End If

Call KewlButtons3_Click
KewlButtons6.Enabled = False

akhar:
KewlButtons6.Enabled = False
End Sub

Private Sub Option1_Click()
Call KewlButtons3_Click
List2.Clear
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.MoveFirst
  Do
     List2.AddItem Adodc1.Recordset.Fields!id
     Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Option2_Click()
Call KewlButtons3_Click
List2.Clear
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc2.Recordset.MoveFirst
  Do
     List2.AddItem Adodc2.Recordset.Fields!id
     Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If
End Sub

Private Sub Option3_Click()
Call KewlButtons3_Click
List2.Clear
If Adodc3.Recordset.RecordCount > 0 Then
  Adodc3.Recordset.MoveFirst
  Do
     List2.AddItem Adodc3.Recordset.Fields!id
     Adodc3.Recordset.MoveNext
  Loop Until Adodc3.Recordset.EOF = True
End If
End Sub

Private Sub Text1_Change(Index As Integer)
If (Index = 4) Or (Index = 6) Or (Index = 7) Then
  Label13.Caption = Amin.moneyaminjoda(Text1(Index).Text)
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
If (Index = 4) Or (Index = 6) Or (Index = 7) Then
  Label13.Visible = True
  Label13.Caption = Amin.moneyaminjoda(Text1(Index).Text)
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
  Case 0
    Text1(1).SetFocus
  
  Case 1
    If (Option1.Value = True) Or (Option2.Value = True) Then
      p = False
      Form3.Adodc1.Recordset.MoveFirst
      Do
        If Form3.Adodc1.Recordset.Fields!id = Trim(Text1(1).Text) Then p = True: Exit Do
        Form3.Adodc1.Recordset.MoveNext
      Loop Until Form3.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label7.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
        Text1(2).SetFocus
      Else
        Label7.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If
    
    If Option3.Value = True Then
      p = False
      Form4.Adodc1.Recordset.MoveFirst
      Do
        If Form4.Adodc1.Recordset.Fields!id = Trim(Text1(1).Text) Then p = True: Exit Do
        Form4.Adodc1.Recordset.MoveNext
      Loop Until Form4.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label7.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
        Text1(2).SetFocus
      Else
        Label7.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If
    
  Case 2
    If (Option1.Value = True) Or (Option2.Value = True) Then
      p = False
      Form3.Adodc1.Recordset.MoveFirst
      Do
        If Form3.Adodc1.Recordset.Fields!id = Trim(Text1(2).Text) Then p = True: Exit Do
        Form3.Adodc1.Recordset.MoveNext
      Loop Until Form3.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label9.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
        Text1(3).SetFocus
      Else
        Label9.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If

    If Option3.Value = True Then
      p = False
      Form4.Adodc1.Recordset.MoveFirst
      Do
        If Form4.Adodc1.Recordset.Fields!id = Trim(Text1(2).Text) Then p = True: Exit Do
        Form4.Adodc1.Recordset.MoveNext
      Loop Until Form4.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label9.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
        Text1(3).SetFocus
      Else
        Label7.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If

    
  Case 3
    If (Option1.Value = True) Or (Option2.Value = True) Then
      p = False
      Form3.Adodc1.Recordset.MoveFirst
      Do
        If Form3.Adodc1.Recordset.Fields!id = Trim(Text1(3).Text) Then p = True: Exit Do
        Form3.Adodc1.Recordset.MoveNext
      Loop Until Form3.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label11.Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
        Text1(4).SetFocus
      Else
        Label11.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If

    If Option3.Value = True Then
      p = False
      Form4.Adodc1.Recordset.MoveFirst
      Do
        If Form4.Adodc1.Recordset.Fields!id = Trim(Text1(3).Text) Then p = True: Exit Do
        Form4.Adodc1.Recordset.MoveNext
      Loop Until Form4.Adodc1.Recordset.EOF = True
    
      If p = True Then
        Label11.Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
        Text1(4).SetFocus
      Else
        Label7.Caption = "-"
        z = MsgBox("òœ Õ”«» Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
      End If
    End If

  Case 4
    Text1(5).SetFocus
  
  Case 5
    Text1(6).SetFocus
  
  Case 6
    Text1(7).SetFocus
  
  Case 7
    Text1(8).SetFocus
  
  Case 8
    Text1(9).SetFocus
  
  Case 9
    If KewlButtons6.Enabled = True Then KewlButtons6.SetFocus
  
End Select
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If (Index = 4) Or (Index = 6) Or (Index = 7) Then
  Text1(Index).Text = Label13.Caption
End If
Label13.Caption = ""
Label13.Visible = False
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List2.ListCount - 1
      If List2.List(q) = Trim(Text16.Text) Then List2.ListIndex = q
   Next q
End If
End Sub
