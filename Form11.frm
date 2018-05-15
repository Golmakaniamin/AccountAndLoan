VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
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
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2460
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6840
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2400
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Connect         =   $"Form11.frx":10378
      OLEDBString     =   $"Form11.frx":1052C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   "haz"
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
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   2535
      Left            =   2040
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6840
      Width           =   2415
   End
   Begin VB.ListBox List5 
      Height          =   2475
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   2460
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1335
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
      Height          =   2460
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1215
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
      Height          =   2460
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1335
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
      Height          =   2460
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   495
      Left            =   11760
      TabIndex        =   9
      Top             =   9480
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Form11.frx":106E0
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
      Caption         =   "ÚãáíÇÊ ÌÏíÏ"
      Height          =   3135
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   8895
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   4
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   1335
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   4455
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ËÈÊ"
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
         MICON           =   "Form11.frx":106FC
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
         Caption         =   "ÚäæÇä"
         Height          =   495
         Index           =   9
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇÑíÎ"
         Height          =   495
         Index           =   3
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÚÏÇÏ"
         Height          =   495
         Index           =   0
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   1
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊæÖíÍÇÊ"
         Height          =   495
         Index           =   2
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "äæÚ ÚãáíÇÊ"
      Height          =   1695
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÏÑÂãÏ"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "åÒíäå"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   9480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÍÐÝ"
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
      MICON           =   "Form11.frx":10718
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
      Height          =   495
      Index           =   12
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÌãÚ ˜á :"
      Height          =   495
      Index           =   7
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   495
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÚäæÇä"
      Height          =   495
      Index           =   11
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊæÖíÍÇÊ"
      Height          =   495
      Index           =   10
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÏíÝ"
      Height          =   495
      Index           =   8
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÈáÛ"
      Height          =   495
      Index           =   6
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÚÏÇÏ"
      Height          =   495
      Index           =   5
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇÑíÎ"
      Height          =   495
      Index           =   4
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "åÒíäå æ ÏÑÂãÏ"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim o As Boolean

Private Sub Form_Activate()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text2.Text = ""
End Sub

Private Sub KewlButtons1_Click()
o = False
For q = 0 To 3
  If Text1(q) = "" Then o = True
Next q
If o = True Then
  z = MsgBox("áØÝÇ ÝíáÏ åÇí ÎÇáí ÑÇ Ê˜ãíá äãÇííÏ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If
Adodc1.Recordset.AddNew
If Option1.Value = True Then Adodc1.Recordset.Fields!no = "åÒíäå": Label4.Caption = Val(List6.ListCount) + 2
If Option2.Value = True Then Adodc1.Recordset.Fields!no = "ÏÑÂãÏ": Label4.Caption = Val(List6.ListCount) + 1
Adodc1.Recordset.Fields!rad = Label4.Caption
Adodc1.Recordset.Fields!men = Text1(0).Text
Adodc1.Recordset.Fields!Money = Text1(1).Text
Adodc1.Recordset.Fields!date1 = Text1(3).Text
Adodc1.Recordset.Fields!promp = Text1(2).Text
Adodc1.Recordset.Fields!subject = Text1(4).Text
Adodc1.Recordset.Fields!user = Form2.Label2.Caption
Adodc1.Recordset.Update

List1.AddItem Label4.Caption
List2.AddItem Text1(0).Text
List3.AddItem Text1(1).Text
List4.AddItem Text1(3).Text
List5.AddItem Text1(2).Text
List6.AddItem Text1(4).Text

For q = 0 To 4
  Text1(q).Text = ""
Next q
Text1(4).SetFocus
akhar:
End Sub

Private Sub KewlButtons2_Click()
If List1.ListIndex = -1 Then
  If Option1.Value = True Then
    z = MsgBox("áØÝÇ åÒíäå ãæÑÏ äÙÑ ÑÇ ÇäÊÎÇÈ äãÇííÏ", vbCritical + vbMsgBoxRight, "")
  Else
    z = MsgBox("áØÝÇ ÏÑ ÂãÏ ãæÑÏ äÙÑ ÑÇ ÇäÊÎÇÈ äãÇííÏ", vbCritical + vbMsgBoxRight, "")
  End If
  GoTo akhar
End If

Adodc1.Recordset.MoveFirst
Do
  If Adodc1.Recordset.Fields!rad = List1.List(List1.ListIndex) Then
    If Option2.Value = True Then
      If Adodc1.Recordset.Fields!no = "ÏÑÂãÏ" Then
        Adodc1.Recordset.Delete
      End If
    Else
      If Adodc1.Recordset.Fields!no = "åÒíäå" Then
        Adodc1.Recordset.Delete
      End If
    End If
  End If
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List3.RemoveItem (List3.ListIndex)
List4.RemoveItem (List4.ListIndex)
List5.RemoveItem (List5.ListIndex)
List6.RemoveItem (List6.ListIndex)
z = MsgBox("ÇØáÇÚÇÊ ÇÒ ÓíÓÊã ÍÐÝ ÔÏ", vbInformation + vbMsgBoxRight, "")
akhar:
End Sub

Private Sub KewlButtons4_Click()
Form2.Show
Me.Hide
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
List6.ListIndex = List1.ListIndex
Text2.Text = List5.List(List5.ListIndex)
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
List6.ListIndex = List2.ListIndex
Text2.Text = List5.List(List5.ListIndex)
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
Text2.Text = List5.List(List5.ListIndex)
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
Text2.Text = List5.List(List5.ListIndex)
End Sub

Private Sub List6_Click()
List2.ListIndex = List6.ListIndex
List3.ListIndex = List6.ListIndex
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
List1.ListIndex = List6.ListIndex
Text2.Text = List5.List(List5.ListIndex)
End Sub

Private Sub Option1_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Text2.Text = ""
Label1(12).Caption = 0
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!no = "åÒíäå" Then
      List1.AddItem Adodc1.Recordset.Fields!rad
      List2.AddItem Adodc1.Recordset.Fields!men
      List3.AddItem Adodc1.Recordset.Fields!Money
      List4.AddItem Adodc1.Recordset.Fields!date1
      List5.AddItem Adodc1.Recordset.Fields!promp
      List6.AddItem Adodc1.Recordset.Fields!subject
      Label1(12).Caption = Val(Label1(12).Caption) + Val(Adodc1.Recordset.Fields!Money)
    End If
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If

'ãÑÊÈ ÓÇÒí

Dim id(1500) As Integer, n1(1500), n2(1500), n3(1500), n4(1500), n5(1500), idt, n1t, n2t, n3t, n4t, n5t, count As String
For intq = 0 To List1.ListCount - 1
    id(intq) = List1.List(intq)
    n1(intq) = List2.List(intq)
    n2(intq) = List3.List(intq)
    n3(intq) = List4.List(intq)
    n4(intq) = List5.List(intq)
    n5(intq) = List6.List(intq)
Next intq
count = List1.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If id(intq) > id(intw) Then
         idt = id(intq)
         n1t = n1(intq)
         n2t = n2(intq)
         n3t = n3(intq)
         n4t = n4(intq)
         n5t = n5(intq)
         
         id(intq) = id(intw)
         n1(intq) = n1(intw)
         n2(intq) = n2(intw)
         n3(intq) = n3(intw)
         n4(intq) = n4(intw)
         n5(intq) = n5(intw)
         
         
         id(intw) = idt
         n1(intw) = n1t
         n2(intw) = n2t
         n3(intw) = n3t
         n4(intw) = n4t
         n5(intw) = n5t
         
      End If
   Next intw
Next intq
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

For intq = 0 To count
   List1.AddItem id(intq)
   List2.AddItem n1(intq)
   List3.AddItem n2(intq)
   List4.AddItem n3(intq)
   List5.AddItem n4(intq)
   List6.AddItem n5(intq)
Next intq
Label1(12).Caption = Amin.moneyaminjoda(Label1(12).Caption)
End Sub

Private Sub Option2_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
Label1(12).Caption = 0
Text2.Text = ""
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!no = "ÏÑÂãÏ" Then
      List1.AddItem Adodc1.Recordset.Fields!rad
      List2.AddItem Adodc1.Recordset.Fields!men
      List3.AddItem Adodc1.Recordset.Fields!Money
      List4.AddItem Adodc1.Recordset.Fields!date1
      List5.AddItem Adodc1.Recordset.Fields!promp
      List6.AddItem Adodc1.Recordset.Fields!subject
      Label1(12).Caption = Val(Label1(12).Caption) + Val(Adodc1.Recordset.Fields!Money)
    End If
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
Label1(12).Caption = Amin.moneyaminjoda(Label1(12).Caption)
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
      KewlButtons1.SetFocus
    
    Case 3
      Text1(0).SetFocus
    
    Case 4
      Text1(3).SetFocus

  End Select
End If
End Sub
