VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Å—œ«Œ  Ê«„"
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
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
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
      Connect         =   $"Form7.frx":10378
      OLEDBString     =   $"Form7.frx":1052C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "pvamvig"
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
      Connect         =   $"Form7.frx":106E0
      OLEDBString     =   $"Form7.frx":10894
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "pvamaz"
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
      Connect         =   $"Form7.frx":10A48
      OLEDBString     =   $"Form7.frx":10BFC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "pvamadi"
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
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   5160
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   2160
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6960
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "’› Ê«„"
      Height          =   3975
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4440
      Width           =   3975
      Begin VB.TextBox Text1 
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
         TabIndex        =   8
         Top             =   720
         Width           =   1575
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
         Height          =   2760
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1095
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
         Left            =   3000
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
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
         Height          =   2760
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   5
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
         Height          =   2760
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C3CEC4&
         Caption         =   " «—ÌŒ À» "
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
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C3CEC4&
         Caption         =   "òœ Õ”«»"
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C3CEC4&
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
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
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
      MICON           =   "Form7.frx":10DB0
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
      Left            =   2280
      TabIndex        =   12
      Top             =   9480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Å—œ«Œ  Ê«„"
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
      MICON           =   "Form7.frx":10DCC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "«‰ Œ«» ‰Ê⁄ Ê«„ "
      Height          =   975
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3360
      Width           =   2895
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
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "«÷ÿ—«—Ì"
         Height          =   495
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÊÌéÂ"
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «›—«œÌ òÂ œ— ’› Ê«„ ﬁ—«— œ«—‰œ :"
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   8520
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Label9"
      DataField       =   "id"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   2160
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
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Å—œ«Œ  Ê«„"
      Height          =   495
      Index           =   10
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ÿ·«⁄«  çò ÷„«‰ "
      Height          =   495
      Index           =   0
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   11
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   10
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» Ê«„ êÌ—‰œÂ :"
      Height          =   495
      Index           =   1
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» ÷«„‰ «Ê· :"
      Height          =   495
      Index           =   2
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Õ”«» ÷«„‰ œÊ„ :"
      Height          =   495
      Index           =   3
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ò·Ì Ê«„ :"
      Height          =   495
      Index           =   4
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ «ﬁ”«ÿ :"
      Height          =   495
      Index           =   5
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ﬁ”ÿ «Ê· :"
      Height          =   495
      Index           =   6
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ »ﬁÌÂ «ﬁ”«ÿ :"
      Height          =   495
      Index           =   7
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ò«—„“œ :"
      Height          =   495
      Index           =   8
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ œ— ŒÊ«”  :"
      Height          =   495
      Index           =   9
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Å—œ«Œ  Ê«„"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Boolean

Private Sub Form_Activate()
Label12.Caption = 0
Option1.SetFocus
Call Option1_Click
End Sub

Private Sub KewlButtons2_Click()
If List1.ListIndex = -1 Then
  z = MsgBox("·ÿ›« Ê«„ „Ê—œ ‰Ÿ— ŒÊœ —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Text2.Text <> "") And (Text3.Text <> "") Then
  If Option1.Value = True Then
    If Adodc1.Recordset.RecordCount > 0 Then
      p = False
      Adodc1.Recordset.MoveFirst
      Do
        If Adodc1.Recordset.Fields!id = List1.List(List1.ListIndex) Then p = True: Exit Do
        Adodc1.Recordset.MoveNext
      Loop Until Adodc1.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!id = List1.List(List1.ListIndex)
      Adodc1.Recordset.Fields!id1 = Label4(0).Caption
      Adodc1.Recordset.Fields!idz1 = Label4(1).Caption
      Adodc1.Recordset.Fields!idz2 = Label4(2).Caption
      Adodc1.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Label4(3).Caption)
      Adodc1.Recordset.Fields!numberagsat = Label4(9).Caption
      Adodc1.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Label4(4).Caption)
      Adodc1.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Label4(10).Caption)
      Adodc1.Recordset.Fields!karmozd = Label4(5).Caption
      Adodc1.Recordset.Fields!Date = Trim(Text3.Text)
      Adodc1.Recordset.Fields!Check = Trim(Text2.Text)
      Adodc1.Recordset.Fields!tasvie = "‰‘œÂ"
      Adodc1.Recordset.Fields!user = Form2.Label2.Caption
      Adodc1.Recordset.Update
    
      Form6.Adodc1.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
      Form6.Adodc1.Recordset.Delete
      List1.RemoveItem (List1.ListIndex)
      List2.RemoveItem (List2.ListIndex)
      List3.RemoveItem (List3.ListIndex)
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
    End If
  End If

  If Option2.Value = True Then
    If Adodc2.Recordset.RecordCount > 0 Then
      p = False
      Adodc2.Recordset.MoveFirst
      Do
        If Adodc2.Recordset.Fields!id = List1.List(List1.ListIndex) Then p = True: Exit Do
        Adodc2.Recordset.MoveNext
      Loop Until Adodc2.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc2.Recordset.AddNew
      Adodc2.Recordset.Fields!id = List1.List(List1.ListIndex)
      Adodc2.Recordset.Fields!id1 = Label4(0).Caption
      Adodc2.Recordset.Fields!idz1 = Label4(1).Caption
      Adodc2.Recordset.Fields!idz2 = Label4(2).Caption
      Adodc2.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Label4(3).Caption)
      Adodc2.Recordset.Fields!numberagsat = Label4(9).Caption
      Adodc2.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Label4(4).Caption)
      Adodc2.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Label4(10).Caption)
      Adodc2.Recordset.Fields!karmozd = Label4(5).Caption
      Adodc2.Recordset.Fields!Date = Trim(Text3.Text)
      Adodc2.Recordset.Fields!Check = Trim(Text2.Text)
      Adodc2.Recordset.Fields!tasvie = "‰‘œÂ"
      Adodc2.Recordset.Fields!user = Form2.Label2.Caption
      Adodc2.Recordset.Update
  
      Form6.Adodc2.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
      Form6.Adodc2.Recordset.Delete
      List1.RemoveItem (List1.ListIndex)
      List2.RemoveItem (List2.ListIndex)
      List3.RemoveItem (List3.ListIndex)
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
    End If
  End If

  If Option3.Value = True Then
    If Adodc3.Recordset.RecordCount > 0 Then
      p = False
      Adodc3.Recordset.MoveFirst
      Do
        If Adodc3.Recordset.Fields!id = List1.List(List1.ListIndex) Then p = True: Exit Do
        Adodc3.Recordset.MoveNext
      Loop Until Adodc3.Recordset.EOF = True
    Else
      p = False
    End If
    If p = False Then
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!id = List1.List(List1.ListIndex)
      Adodc3.Recordset.Fields!id1 = Label4(0).Caption
      Adodc3.Recordset.Fields!idz1 = Label4(1).Caption
      Adodc3.Recordset.Fields!idz2 = Label4(2).Caption
      Adodc3.Recordset.Fields!moneyvam = Amin.moneyaminnojoda(Label4(3).Caption)
      Adodc3.Recordset.Fields!numberagsat = Label4(9).Caption
      Adodc3.Recordset.Fields!moneyg1 = Amin.moneyaminnojoda(Label4(4).Caption)
      Adodc3.Recordset.Fields!moneyg2 = Amin.moneyaminnojoda(Label4(10).Caption)
      Adodc3.Recordset.Fields!karmozd = Label4(5).Caption
      Adodc3.Recordset.Fields!Date = Trim(Text3.Text)
      Adodc3.Recordset.Fields!Check = Trim(Text2.Text)
      Adodc3.Recordset.Fields!tasvie = "‰‘œÂ"
      Adodc3.Recordset.Fields!user = Form2.Label2.Caption
      Adodc3.Recordset.Update
  
      Form6.Adodc3.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
      Form6.Adodc3.Recordset.Delete
      List1.RemoveItem (List1.ListIndex)
      List2.RemoveItem (List2.ListIndex)
      List3.RemoveItem (List3.ListIndex)
      z = MsgBox("«ÿ·«⁄«  ‘„« »« „Ê›ﬁÌ  œ— ”Ì” „ À»  ‘œ", vbMsgBoxRight + vbInformation, "")
    Else
      z = MsgBox("òœ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
    End If
  End If
Else
  z = MsgBox("·ÿ›« ›Ì·œ Â«Ì „—»ÊÿÂ —«  ò„Ì· ‰„«ÌÌœ", vbMsgBoxRight + vbCritical, "")
End If

akhar:
End Sub

Private Sub KewlButtons3_Click()
Form2.Show
Me.Hide
End Sub

Private Sub List1_Click()
Text3.Text = Form2.Label5.Caption
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
If Option1.Value = True Then
  If Form6.Adodc1.Recordset.RecordCount > 0 Then
    Form6.Adodc1.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
    Label4(0).Caption = Form6.Adodc1.Recordset.Fields!id1
    Label4(1).Caption = Form6.Adodc1.Recordset.Fields!idz1
    Label4(2).Caption = Form6.Adodc1.Recordset.Fields!idz2
    Label4(3).Caption = Amin.moneyaminjoda(Form6.Adodc1.Recordset.Fields!moneyvam)
    Label4(4).Caption = Amin.moneyaminjoda(Form6.Adodc1.Recordset.Fields!moneyg1)
    Label4(5).Caption = Form6.Adodc1.Recordset.Fields!karmozd
    Label4(9).Caption = Form6.Adodc1.Recordset.Fields!numberagsat
    Label4(10).Caption = Amin.moneyaminjoda(Form6.Adodc1.Recordset.Fields!moneyg2)
    Label4(11).Caption = Form6.Adodc1.Recordset.Fields!datet
    Form3.Adodc1.Recordset.Find "id='" & Label4(0).Caption & "'", , adSearchForward, 1
    Label4(6).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
    
    Form3.Adodc1.Recordset.Find "id='" & Label4(1).Caption & "'", , adSearchForward, 1
    Label4(7).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
    
    Form3.Adodc1.Recordset.Find "id='" & Label4(2).Caption & "'", , adSearchForward, 1
    Label4(8).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  End If
End If

If Option2.Value = True Then
  If Form6.Adodc2.Recordset.RecordCount > 0 Then
    Form6.Adodc2.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
    Label4(0).Caption = Form6.Adodc2.Recordset.Fields!id1
    Label4(1).Caption = Form6.Adodc2.Recordset.Fields!idz1
    Label4(2).Caption = Form6.Adodc2.Recordset.Fields!idz2
    Label4(3).Caption = Amin.moneyaminjoda(Form6.Adodc2.Recordset.Fields!moneyvam)
    Label4(4).Caption = Amin.moneyaminjoda(Form6.Adodc2.Recordset.Fields!moneyg1)
    Label4(5).Caption = Form6.Adodc2.Recordset.Fields!karmozd
    Label4(9).Caption = Form6.Adodc2.Recordset.Fields!numberagsat
    Label4(10).Caption = Amin.moneyaminjoda(Form6.Adodc2.Recordset.Fields!moneyg2)
    Label4(11).Caption = Form6.Adodc2.Recordset.Fields!datet
    Form3.Adodc1.Recordset.Find "id='" & Label4(0).Caption & "'", , adSearchForward, 1
    Label4(6).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
    
    Form3.Adodc1.Recordset.Find "id='" & Label4(1).Caption & "'", , adSearchForward, 1
    Label4(7).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
    
    Form3.Adodc1.Recordset.Find "id='" & Label4(2).Caption & "'", , adSearchForward, 1
    Label4(8).Caption = Form3.Adodc1.Recordset.Fields!Name & " " & Form3.Adodc1.Recordset.Fields!family
  End If
End If

If Option3.Value = True Then
  If Form6.Adodc3.Recordset.RecordCount > 0 Then
    Form6.Adodc3.Recordset.Find "id='" & List1.List(List1.ListIndex) & "'", , adSearchForward, 1
    Label4(0).Caption = Form6.Adodc3.Recordset.Fields!id1
    Label4(1).Caption = Form6.Adodc3.Recordset.Fields!idz1
    Label4(2).Caption = Form6.Adodc3.Recordset.Fields!idz2
    Label4(3).Caption = Amin.moneyaminjoda(Form6.Adodc3.Recordset.Fields!moneyvam)
    Label4(4).Caption = Amin.moneyaminjoda(Form6.Adodc3.Recordset.Fields!moneyg1)
    Label4(5).Caption = Form6.Adodc3.Recordset.Fields!karmozd
    Label4(9).Caption = Form6.Adodc3.Recordset.Fields!numberagsat
    Label4(10).Caption = Amin.moneyaminjoda(Form6.Adodc3.Recordset.Fields!moneyg2)
    Label4(11).Caption = Form6.Adodc3.Recordset.Fields!datet
    Form4.Adodc1.Recordset.Find "id='" & Label4(0).Caption & "'", , adSearchForward, 1
    Label4(6).Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
    
    Form4.Adodc1.Recordset.Find "id='" & Label4(1).Caption & "'", , adSearchForward, 1
    Label4(7).Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
    
    Form4.Adodc1.Recordset.Find "id='" & Label4(2).Caption & "'", , adSearchForward, 1
    Label4(8).Caption = Form4.Adodc1.Recordset.Fields!Name & " " & Form4.Adodc1.Recordset.Fields!family
  End If
End If
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
End Sub

Private Sub Option1_Click()
Dim id(1000), na(1000), da(1000), idt, nat, dat, count As String
For q = 0 To 11
 Label4(q).Caption = ""
Next q
Text2.Text = ""
Text3.Text = ""
List1.Clear
List2.Clear
List3.Clear
If Form6.Adodc1.Recordset.RecordCount > 0 Then
  Form6.Adodc1.Recordset.MoveFirst
  Do
     List1.AddItem Form6.Adodc1.Recordset.Fields!id
     List2.AddItem Form6.Adodc1.Recordset.Fields!id1
     List3.AddItem Form6.Adodc1.Recordset.Fields!datet
     Form6.Adodc1.Recordset.MoveNext
  Loop Until Form6.Adodc1.Recordset.EOF = True
End If

'
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    da(intq) = List3.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If da(intq) > da(intw) Then
         idt = id(intq)
         nat = na(intq)
         dat = da(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         da(intq) = da(intw)
         
         id(intw) = idt
         na(intw) = nat
         da(intw) = dat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem da(intq)
Next intq
'
Label12.Caption = List1.ListCount
End Sub

Private Sub Option2_Click()
Dim id(1000), na(1000), da(1000), idt, nat, dat, count As String
For q = 0 To 11
 Label4(q).Caption = ""
Next q
Text2.Text = ""
Text3.Text = ""
List1.Clear
List2.Clear
List3.Clear
If Form6.Adodc2.Recordset.RecordCount > 0 Then
  Form6.Adodc2.Recordset.MoveFirst
  Do
     List1.AddItem Form6.Adodc2.Recordset.Fields!id
     List2.AddItem Form6.Adodc2.Recordset.Fields!id1
     List3.AddItem Form6.Adodc2.Recordset.Fields!datet
     Form6.Adodc2.Recordset.MoveNext
  Loop Until Form6.Adodc2.Recordset.EOF = True
End If

'
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    da(intq) = List3.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If da(intq) > da(intw) Then
         idt = id(intq)
         nat = na(intq)
         dat = da(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         da(intq) = da(intw)
         
         id(intw) = idt
         na(intw) = nat
         da(intw) = dat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem da(intq)
Next intq
'
Label12.Caption = List1.ListCount
End Sub

Private Sub Option3_Click()
Dim id(1000), na(1000), da(1000), idt, nat, dat, count As String
For q = 0 To 11
 Label4(q).Caption = ""
Next q
Text2.Text = ""
Text3.Text = ""
List1.Clear
List2.Clear
List3.Clear
If Form6.Adodc3.Recordset.RecordCount > 0 Then
  Form6.Adodc3.Recordset.MoveFirst
  Do
     List1.AddItem Form6.Adodc3.Recordset.Fields!id
     List2.AddItem Form6.Adodc3.Recordset.Fields!id1
     List3.AddItem Form6.Adodc3.Recordset.Fields!datet
     Form6.Adodc3.Recordset.MoveNext
  Loop Until Form6.Adodc3.Recordset.EOF = True
End If

'
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    na(intq) = List2.List(intq)
    da(intq) = List3.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If da(intq) > da(intw) Then
         idt = id(intq)
         nat = na(intq)
         dat = da(intq)
         
         id(intq) = id(intw)
         na(intq) = na(intw)
         da(intq) = da(intw)
         
         id(intw) = idt
         na(intw) = nat
         da(intw) = dat
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem na(intq)
   List3.AddItem da(intq)
Next intq
'
Label12.Caption = List1.ListCount
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List3.ListCount - 1
      If List3.List(q) = Trim(Text1.Text) Then List1.ListIndex = q
   Next q
End If
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List1.ListCount - 1
      If List1.List(q) = Trim(Text15.Text) Then List1.ListIndex = q
   Next q
End If
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   For q = 0 To List2.ListCount - 1
      If List2.List(q) = Trim(Text16.Text) Then List2.ListIndex = q
   Next q
End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons2.SetFocus
End Sub
