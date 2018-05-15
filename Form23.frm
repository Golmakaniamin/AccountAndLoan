VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form23 
   BorderStyle     =   0  'None
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
   LinkTopic       =   "Form23"
   Picture         =   "Form23.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List6 
      Height          =   3165
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2040
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"Form23.frx":10378
      OLEDBString     =   $"Form23.frx":1052C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "printamal"
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
   Begin VB.ListBox List5 
      Height          =   2820
      Left            =   600
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3165
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5400
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   3165
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5400
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   3165
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ListBox List4 
      Height          =   3165
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1387/03/10"
      Top             =   3360
      Width           =   1215
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   9360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "ÈÇÒÔÊ"
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
      MICON           =   "Form23.frx":106E0
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
      Left            =   9240
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "äãÇíÔ"
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
      MICON           =   "Form23.frx":106FC
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
      Left            =   2520
      TabIndex        =   20
      Top             =   9480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Ç"
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
      MICON           =   "Form23.frx":10718
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
      Caption         =   "˜ÇÑÈÑ"
      Height          =   495
      Index           =   11
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   10
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "äÞÏíäí :"
      Height          =   495
      Index           =   9
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÌãÚ ÑÏÇÎÊí :"
      Height          =   495
      Index           =   7
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÌãÚ ÏÑíÇÝÊí :"
      Height          =   495
      Index           =   2
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÏíÝ"
      Height          =   495
      Index           =   3
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "äÇã æ äÇã ÎÇäæÇÏí"
      Height          =   495
      Index           =   4
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÊæÖíÍÇÊ"
      Height          =   495
      Index           =   5
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "˜Ï ÚÖæíÊ"
      Height          =   495
      Index           =   1
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Úãá˜ÑÏ ÑæÒÇäå"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇÑíÎ :"
      Height          =   495
      Index           =   0
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub Form_Load()
Text1.Text = Form2.Label5.Caption
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons2_Click()
Form24.Show
End Sub

Private Sub KewlButtons3_Click()
Dim z As String
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Label1(6).Caption = 0
Label1(8).Caption = 0
Label1(10).Caption = 0
'ãæÌæí ÍÓÇÈ ÚÇÏí
If Form5.Adodc1.Recordset.RecordCount > 0 Then
  Form5.Adodc1.Recordset.MoveFirst
  Do
    If Form5.Adodc1.Recordset.Fields!Date = Text1.Text Then
      List1.AddItem List1.ListCount + 1
      List2.AddItem Form5.Adodc1.Recordset.Fields!id
      Form3.Adodc1.Recordset.Find "id='" + Form5.Adodc1.Recordset.Fields!id + "'", , adSearchForward, 1
      List3.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
      z = ""
      If Form5.Adodc1.Recordset.Fields!Amal = "˜ÓÑ" Then
        z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form5.Adodc1.Recordset.Fields!Money))) + " ÑíÇá ÇÒ ÍÓÇÈ ÚÇÏí " + Form5.Adodc1.Recordset.Fields!Amal + " ÔÏå ÇÓÊ"
        List5.AddItem -1 * Form5.Adodc1.Recordset.Fields!Money
      Else
        z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form5.Adodc1.Recordset.Fields!Money))) + " ÑíÇá Èå ÍÓÇÈ ÚÇÏí " + Form5.Adodc1.Recordset.Fields!Amal + " ÏÇÏå ÔÏå ÇÓÊ"
        List5.AddItem Form5.Adodc1.Recordset.Fields!Money
      End If
      List4.AddItem z
      List6.AddItem Form5.Adodc1.Recordset.Fields!user
    End If
    Form5.Adodc1.Recordset.MoveNext
  Loop Until Form5.Adodc1.Recordset.EOF = True
End If

'ãæÌæÏí ÍÓÇÈ æíŽå
If Form5.Adodc2.Recordset.RecordCount > 0 Then
  Form5.Adodc2.Recordset.MoveFirst
  Do
    If Form5.Adodc2.Recordset.Fields!Date = Text1.Text Then
      List1.AddItem List1.ListCount + 1
      Form4.Adodc1.Recordset.Find "id='" + Form5.Adodc2.Recordset.Fields!id + "'", , adSearchForward, 1
      List3.AddItem Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
      List2.AddItem Form4.Adodc1.Recordset.Fields!idadi
      z = ""
      If Form5.Adodc2.Recordset.Fields!Amal = "˜ÓÑ" Then
        z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form5.Adodc2.Recordset.Fields!Money))) + " ÑíÇá ÇÒ ÍÓÇÈ æíŽå " + Form5.Adodc2.Recordset.Fields!Amal + " ÔÏå ÇÓÊ"
        List5.AddItem -1 * Form5.Adodc2.Recordset.Fields!Money
      Else
        z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form5.Adodc2.Recordset.Fields!Money))) + " ÑíÇá Èå ÍÓÇÈ æíŽå " + Form5.Adodc2.Recordset.Fields!Amal + " ÏÇÏå ÔÏå ÇÓÊ"
        List5.AddItem Form5.Adodc2.Recordset.Fields!Money
      End If
      List4.AddItem z
      List6.AddItem Form5.Adodc2.Recordset.Fields!user
    End If
    Form5.Adodc2.Recordset.MoveNext
  Loop Until Form5.Adodc2.Recordset.EOF = True
End If

'ÇÞÓÇØ æÇã ÚÇÏí
If Form8.Adodc1.Recordset.RecordCount > 0 Then
  Form8.Adodc1.Recordset.MoveFirst
  Do
    If Form8.Adodc1.Recordset.Fields!Date = Text1.Text Then
      
      List1.AddItem List1.ListCount + 1
      Form7.Adodc1.Recordset.Find "id='" + Form8.Adodc1.Recordset.Fields!id + "'", , adSearchForward, 1
      List2.AddItem Form7.Adodc1.Recordset.Fields!id1
      Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc1.Recordset.Fields!id1 + "'", , adSearchForward, 1
      List3.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
      z = ""
      z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form8.Adodc1.Recordset.Fields!Money))) + " ÑíÇá ÈÇÈÊ ÇÞÓÇØ æÇã ÚÇÏí ÔãÇÑå " + Form8.Adodc1.Recordset.Fields!id + " ÏÑíÇÝÊ ÔÏå ÇÓÊ"
      List4.AddItem z
      List5.AddItem Form8.Adodc1.Recordset.Fields!Money
      List6.AddItem Form8.Adodc1.Recordset.Fields!user
    End If
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
End If

'ÇÞÓÇØ æÇã ÇÖØÑÇÑí
If Form8.Adodc2.Recordset.RecordCount > 0 Then
  Form8.Adodc2.Recordset.MoveFirst
  Do
    If Form8.Adodc2.Recordset.Fields!Date = Text1.Text Then
      
      List1.AddItem List1.ListCount + 1
      Form7.Adodc2.Recordset.Find "id='" + Form8.Adodc2.Recordset.Fields!id + "'", , adSearchForward, 1
      List2.AddItem Form7.Adodc2.Recordset.Fields!id1
      Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc2.Recordset.Fields!id1 + "'", , adSearchForward, 1
      List3.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
      z = ""
      z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form8.Adodc2.Recordset.Fields!Money))) + " ÑíÇá ÈÇÈÊ ÇÞÓÇØ æÇã ÇÖØÑÇÑí ÔãÇÑå " + Form8.Adodc2.Recordset.Fields!id + " ÏÑíÇÝÊ ÔÏå ÇÓÊ"
      List4.AddItem z
      List5.AddItem Form8.Adodc2.Recordset.Fields!Money
      List6.AddItem Form8.Adodc2.Recordset.Fields!user
    End If
    Form8.Adodc2.Recordset.MoveNext
  Loop Until Form8.Adodc2.Recordset.EOF = True
End If

'ÇÞÓÇØ æÇã æíŽå
If Form8.Adodc3.Recordset.RecordCount > 0 Then
  Form8.Adodc3.Recordset.MoveFirst
  Do
    If Form8.Adodc3.Recordset.Fields!Date = Text1.Text Then
      
      List1.AddItem List1.ListCount + 1
      z = Form8.Adodc3.Recordset.Fields!id
      
      Form7.Adodc3.Recordset.Find "id='" + z + "'", , adSearchForward, 1
      z = Form7.Adodc3.Recordset.Fields!id1
      
      Form4.Adodc1.Recordset.Find "id='" + z + "'", , adSearchForward, 1
      z = Form4.Adodc1.Recordset.Fields!idadi
      
      Form3.Adodc1.Recordset.Find "id='" + z + "'", , adSearchForward, 1
      List3.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
      List2.AddItem Form3.Adodc1.Recordset.Fields!id
      
      z = ""
      z = "ãÈáÛ " + Amin.moneyaminjoda(Trim(Str(Form8.Adodc3.Recordset.Fields!Money))) + " ÑíÇá ÈÇÈÊ ÇÞÓÇØ æÇã æíŽå ÔãÇÑå " + Form8.Adodc3.Recordset.Fields!id + " ÏÑíÇÝÊ ÔÏå ÇÓÊ"
      List4.AddItem z
      List5.AddItem Form8.Adodc3.Recordset.Fields!Money
      List6.AddItem Form8.Adodc3.Recordset.Fields!user
    End If
    Form8.Adodc3.Recordset.MoveNext
  Loop Until Form8.Adodc3.Recordset.EOF = True
End If

For q = 0 To List5.ListCount - 1
  If List5.List(q) >= 0 Then
    Label1(6).Caption = Val(Label1(6).Caption) + List5.List(q)
  Else
    Label1(8).Caption = Val(Label1(8).Caption) + (-1 * List5.List(q))
  End If
Next q
Label1(10).Caption = Val(Label1(6).Caption) - Val(Label1(8).Caption)

'Ê˜ÑÇÑ


Label1(6).Caption = Amin.moneyaminjoda(Label1(6).Caption)
Label1(8).Caption = Amin.moneyaminjoda(Label1(8).Caption)
Label1(10).Caption = Amin.moneyaminjoda(Label1(10).Caption)

fso.CopyFile "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\Data\info2.mdb", "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\info2.mdb", True
Adodc1.Refresh

For q = 0 To List1.ListCount - 1
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!rad = Val(List1.List(q))
  Adodc1.Recordset.Fields!id = List2.List(q)
  Adodc1.Recordset.Fields!Name = List3.List(q)
  Adodc1.Recordset.Fields!promp = List4.List(q)
  Adodc1.Recordset.Fields!date1 = List6.List(q)
  Adodc1.Recordset.Update
Next q
Adodc1.Refresh
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List6.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List6.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
End Sub

Private Sub List6_Click()
List1.ListIndex = List6.ListIndex
List2.ListIndex = List6.ListIndex
List3.ListIndex = List6.ListIndex
List4.ListIndex = List6.ListIndex
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  KewlButtons3.SetFocus
End If
End Sub
