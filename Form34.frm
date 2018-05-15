VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form34 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ”«»—”Ì"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "1230000"
      Top             =   1440
      Width           =   2295
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Index           =   10
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "„Õ«”»Â"
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
      BCOL            =   8454016
      BCOLO           =   8454016
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   8454016
      MPTR            =   1
      MICON           =   "Form34.frx":0000
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
      Height          =   330
      Left            =   120
      Top             =   120
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
      Connect         =   $"Form34.frx":001C
      OLEDBString     =   $"Form34.frx":01C2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "hesab"
      Caption         =   "Adodc2"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   5
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ò«Â‘ Õ”«» ⁄«œÌ"
      Height          =   495
      Index           =   5
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "»«‰ò"
      Height          =   495
      Index           =   49
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   50
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰ﬁÿÂ ”— »Â ”—"
      Height          =   495
      Index           =   50
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "œ— ¬„œ"
      Height          =   495
      Index           =   4
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   4
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«ﬁ”«ÿ Ê«„ Â«Ì ⁄«œÌ"
      Height          =   495
      Index           =   3
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ò«—„“œ Ê«„ Â«Ì ⁄«œÌ"
      Height          =   495
      Index           =   2
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   3
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   2
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ê«„ Â«Ì ⁄«œÌ"
      Height          =   495
      Index           =   1
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«›“«Ì‘ Õ”«» ⁄«œÌ"
      Height          =   495
      Index           =   0
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   1
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Index           =   0
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub KewlButtons1_Click(Index As Integer)
Dim db1 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset
db1.Open Adodc1.ConnectionString, "ADMIN", "pratic1"
  '«›“«Ì‘ Õ”«» ⁄«œÌ
  rs(0).Open "SELECT Sum(money) As rssum FROM MoneyAdi WHERE (Amal='«›“«Ì‘')", db1
    Label1(0).Caption = rs(0).Fields!rssum
  rs(0).Close

  'ò«Â‘ Õ”«» ⁄«œÌ
  rs(0).Open "SELECT Sum(money) As rssum FROM MoneyAdi WHERE (Amal='ò”—')", db1
    Label1(5).Caption = rs(0).Fields!rssum
  rs(0).Close

  'Ê«„ Â«Ì ⁄«œÌ
  rs(1).Open "SELECT Sum(moneyvam) As rssum FROM pvamadi", db1
    Label1(1).Caption = rs(1).Fields!rssum
  rs(1).Close

  'ò«—„“œ Ê«„ Â«Ì ⁄«œÌ
  rs(2).Open "SELECT Sum(karmozd) As rssum FROM pvamadi", db1
    Label1(2).Caption = rs(2).Fields!rssum
  rs(2).Close

  '«ﬁ”«ÿ Ê«„ Â«Ì ⁄«œÌ
  rs(3).Open "SELECT Sum(money) As rssum FROM GvamAdi", db1
    Label1(3).Caption = rs(3).Fields!rssum
  rs(3).Close

  'œ—¬„œ
  rs(4).Open "SELECT Sum(money) As rssum FROM haz WHERE (no1=1)", db1
    Label1(4).Caption = rs(4).Fields!rssum
  rs(4).Close
  
db1.Close

Label1(50).Caption = Val(((Val(Label1(0).Caption) - Val(Label1(5).Caption)) + Val(Label1(3).Caption) + Val(Label1(4).Caption) + Val(Label1(2).Caption)) - ((Val(Label1(1).Caption) - Val(Label1(2).Caption)))) - Val(Text1.Text)

End Sub
