VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form27 
   BorderStyle     =   0  'None
   Caption         =   "ÅÌ«„ò"
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
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form27.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Connect         =   $"Form27.frx":10378
      OLEDBString     =   $"Form27.frx":1052C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "sendsms"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C3CEC4&
      Caption         =   "ÅÌ«„ Â«Ì Œ«’"
      Height          =   4935
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   6015
      Begin VB.ListBox List1 
         Height          =   3510
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   2295
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "«⁄÷«¡ Õ–› ‰‘œÂ"
         Height          =   495
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3720
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "«⁄÷«¡ Õ–› ‘œÂ"
         Height          =   495
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   " „«„Ì «⁄÷«¡"
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
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
         Height          =   3690
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin KewlButtonz.KewlButtons KewlButtons9 
         Height          =   135
         Left            =   3600
         TabIndex        =   13
         Top             =   4560
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Form27.frx":106E0
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
         Left            =   240
         TabIndex        =   18
         Top             =   4200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "«—”«·"
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
         MICON           =   "Form27.frx":106FC
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
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ :"
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "„ ‰ ÅÌ«„"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   615
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Œ—ÊÃÌ"
      Height          =   3615
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
      Begin KewlButtonz.KewlButtons KewlButtons6 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÅÌ«„ Â«Ì Œ«’"
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
         MICON           =   "Form27.frx":10718
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
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " «ŒÌ— «ﬁ”«ÿ"
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
         MICON           =   "Form27.frx":10734
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
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ò”—Ì Õﬁ ⁄÷ÊÌ "
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
         MICON           =   "Form27.frx":10750
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons11 
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "«—”«· ÅÌ«„ Œ«’"
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
         MICON           =   "Form27.frx":1076C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons12 
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "÷„«‰  Â«"
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
         MICON           =   "Form27.frx":10788
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "Ê—ÊœÌ"
      Height          =   2895
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "»«ﬁÌ„«‰œÂ «ﬁ”«ÿ"
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
         MICON           =   "Form27.frx":107A4
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
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÅÌ«„ Â«Ì „œÌ—"
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
         MICON           =   "Form27.frx":107C0
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "„ÊÃÊœÌ Õ”«» Â«"
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
         MICON           =   "Form27.frx":107DC
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
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "œ—Ì«›  «ﬁ”«ÿ"
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
         MICON           =   "Form27.frx":107F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11160
      TabIndex        =   0
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
      MICON           =   "Form27.frx":10814
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2280
      Top             =   1920
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
      Connect         =   $"Form27.frx":10830
      OLEDBString     =   $"Form27.frx":109E4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "smsread"
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "promp"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "promp"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÅÌ«„ò"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim w As String

Private Sub hide_1()
Frame3.Visible = False
End Sub

Private Sub Form_Activate()
Frame3.Visible = False
Unload Form29

End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Form29.Show
Me.Hide
End Sub

Private Sub KewlButtons10_Click()
If Text1.Text <> "" Then
    Form29.Timer1.Enabled = False
    Form29.Timer2.Enabled = False
    Form29.Adodc4.Recordset.AddNew
    Form29.Adodc4.Recordset.Fields!date1 = Form2.Label5.Caption
    Form29.Adodc4.Recordset.Update
    For q = 0 To List2.ListCount - 1
      If List2.Selected(q) = True Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!number = List1.List(q)
        Adodc1.Recordset.Fields!promp = Text1.Text
        Adodc1.Recordset.Fields!no = "ÅÌ«„ Œ«’"
        Adodc1.Recordset.Fields!no1 = "1"
        Adodc1.Recordset.Fields!user = Form2.Label2.Caption
        Adodc1.Recordset.Fields!Time = Form2.Label7.Caption
        Adodc1.Recordset.Fields!Date = Form2.Label5.Caption
        Adodc1.Recordset.Fields!send = "0"
        Adodc1.Recordset.Fields!delivery = "0"
        Adodc1.Recordset.Update
      End If
    Next q
    MsgBox "ÅÌ«„ Œ«’ »« „Ê›ﬁÌ  »—«Ì  „«„Ì «›—«œ «‰ Œ«» ‘œÂ «—”«· ‘œ", vbCritical + vbMsgBoxRight, ""
    Form29.Timer1.Enabled = True
Else
  MsgBox "„ ‰Ì „ÊÃÊœ ‰Ì” ", vbCritical + vbMsgBoxRight, ""
End If
End Sub

Private Sub KewlButtons11_Click()
Call hide_1
Frame3.Visible = True
List1.Clear
List2.Clear
Form3.Adodc1.Recordset.MoveFirst
Do
  If Len(Form3.Adodc1.Recordset.Fields!mobile) = 11 Then
    List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    List1.AddItem Form3.Adodc1.Recordset.Fields!mobile
  End If
  Form3.Adodc1.Recordset.MoveNext
Loop Until Form3.Adodc1.Recordset.EOF = True
For q = 0 To List2.ListCount - 1
  List2.Selected(q) = True
Next q
Label5.Caption = List2.ListCount
End Sub


Private Sub KewlButtons6_Click()
Call hide_1
Form30.Adodc1.CommandType = 8
Form30.Adodc1.RecordSource = "select * from sendsms"
Form30.Adodc1.Refresh
Form30.Adodc1.RecordSource = "select * from sendsms where no1='3'"
Form30.Adodc1.Refresh
Form30.DataGrid1.Caption = "ÅÌ«„ò Â«Ì «—”«· ‘œÂ «“ ÿ—› ’‰œÊﬁ"
Form30.DataGrid1.Refresh
Form30.Show
End Sub

Private Sub KewlButtons7_Click()
Call hide_1
Form30.Adodc1.CommandType = 8
Form30.Adodc1.RecordSource = "select * from sendsms"
Form30.Adodc1.Refresh
Form30.Adodc1.RecordSource = "select * from sendsms where no1='1'"
Form30.Adodc1.Refresh
Form30.DataGrid1.Caption = "ÅÌ«„ò Â«Ì «—”«· ‘œÂ »Â ⁄·   «ŒÌ— «ﬁ”«ÿ"
Form30.DataGrid1.Refresh
Form30.Show
End Sub

Private Sub KewlButtons8_Click()
Call hide_1
Form30.Adodc1.CommandType = 8
Form30.Adodc1.RecordSource = "select * from sendsms"
Form30.Adodc1.Refresh
Form30.Adodc1.RecordSource = "select * from sendsms where no1='2'"
Form30.Adodc1.Refresh
Form30.DataGrid1.Caption = "ÅÌ«„ò Â«Ì «—”«· ‘œÂ »Â ⁄·  ò”—Ì „ÊÃÊœÌ"
Form30.DataGrid1.Refresh
Form30.Show
End Sub

Private Sub KewlButtons9_Click()
Dim na(1000), nat, count As String
For intq = 0 To List2.ListCount - 1
    na(intq) = List2.List(intq)
Next intq
count = List2.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If na(intq) > na(intw) Then
         nat = na(intq)
         
         na(intq) = na(intw)
         
         na(intw) = nat
      End If
   Next intw
Next intq
List2.Clear
For intq = 0 To count
   List2.AddItem na(intq)
Next intq
End Sub

Private Sub Option1_Click()
List1.Clear
List2.Clear
Form3.Adodc1.Recordset.MoveFirst
Do
  If Len(Form3.Adodc1.Recordset.Fields!mobile) = 11 Then
    List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    List1.AddItem Form3.Adodc1.Recordset.Fields!mobile
  End If
  Form3.Adodc1.Recordset.MoveNext
Loop Until Form3.Adodc1.Recordset.EOF = True
For q = 0 To List2.ListCount - 1
  List2.Selected(q) = True
Next q
Label5.Caption = List2.ListCount
End Sub

Private Sub Option2_Click()
List1.Clear
List2.Clear
Form3.Adodc1.Recordset.MoveFirst
Do
  If (Form3.Adodc1.Recordset.Fields!Delete = "‘œÂ") And (Len(Form3.Adodc1.Recordset.Fields!mobile) = 11) Then
    List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    List1.AddItem Form3.Adodc1.Recordset.Fields!mobile
  End If
  Form3.Adodc1.Recordset.MoveNext
Loop Until Form3.Adodc1.Recordset.EOF = True
For q = 0 To List2.ListCount - 1
  List2.Selected(q) = True
Next q
Label5.Caption = List2.ListCount
End Sub

Private Sub Option3_Click()
List1.Clear
List2.Clear
Form3.Adodc1.Recordset.MoveFirst
Do
  If (Form3.Adodc1.Recordset.Fields!Delete = "‰‘œÂ") And (Len(Form3.Adodc1.Recordset.Fields!mobile) = 11) Then
    List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    List1.AddItem Form3.Adodc1.Recordset.Fields!mobile
  End If
  Form3.Adodc1.Recordset.MoveNext
Loop Until Form3.Adodc1.Recordset.EOF = True
For q = 0 To List2.ListCount - 1
  List2.Selected(q) = True
Next q
Label5.Caption = List2.ListCount
End Sub
