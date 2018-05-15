VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "«ﬁ”«ÿ  «ŒÌ—Ì"
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
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   1
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Text            =   "1300/00/00"
      Top             =   3600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   2280
      Top             =   2640
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
      Connect         =   $"Form9.frx":10378
      OLEDBString     =   $"Form9.frx":1052C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "dir"
      Caption         =   "Adodc4"
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
      Height          =   495
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "«‰ Œ«» ‰Ê⁄ Ê«„ "
      Height          =   975
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   2895
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "⁄«œÌ"
         Height          =   495
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
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
   Begin VB.ListBox List7 
      Height          =   3165
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2055
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11400
      TabIndex        =   4
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
      MICON           =   "Form9.frx":106E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List6 
      Height          =   3165
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ListBox List5 
      Height          =   3165
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   3165
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   3165
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   3165
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3165
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1386/07/07"
      Top             =   3600
      Width           =   1215
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
      CommandType     =   8
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
      Connect         =   $"Form9.frx":106FC
      OLEDBString     =   $"Form9.frx":108B0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   ""
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
      CommandType     =   8
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
      Connect         =   $"Form9.frx":10A64
      OLEDBString     =   $"Form9.frx":10C18
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   ""
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
      CommandType     =   8
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
      Connect         =   $"Form9.frx":10DCC
      OLEDBString     =   $"Form9.frx":10F80
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   "pratic1"
      RecordSource    =   ""
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
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   495
      Left            =   2400
      TabIndex        =   23
      Top             =   9480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      MICON           =   "Form9.frx":11134
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
      Caption         =   " «"
      Height          =   495
      Index           =   2
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Õ”«»"
      Height          =   495
      Index           =   1
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ Â„—«Â"
      Height          =   495
      Index           =   8
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ¬Œ—Ì‰ Å—œ«Œ "
      Height          =   495
      Index           =   7
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   495
      Index           =   6
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      Height          =   495
      Index           =   5
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Ê«„"
      Height          =   495
      Index           =   4
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
      Height          =   495
      Index           =   3
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ :        «“"
      Height          =   495
      Index           =   0
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬁ”«ÿ  «ŒÌ—Ì"
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
      TabIndex        =   5
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim a As String, b As String

Private Sub Amin_1()
Dim stry As String, strm As String
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
strm = Val(Mid(Text1(0).Text, 6, 2)) - 1
stry = Mid(Text1(0).Text, 1, 4)
If strm = 0 Then
  strm = 12
  stry = stry - 1
End If
If Len(strm) < 2 Then strm = "0" + strm
Text2.Text = stry + "/" + strm + Mid(Text1(0).Text, 8, 3)

strm = Val(Mid(Text1(1).Text, 6, 2)) - 1
stry = Mid(Text1(1).Text, 1, 4)
If strm = 0 Then
  strm = 12
  stry = stry - 1
End If
If Len(strm) < 2 Then strm = "0" + strm
Text3.Text = stry + "/" + strm + Mid(Text1(1).Text, 8, 3)
End Sub

Private Sub Amin_2()
Dim fil1(1000), fil2(1000), fil3(1000), fil4(1000), fil5(1000), fil6(1000), count As String
Dim fil1t, fil2t, fil3t, fil4t, fil5t, fil6t As String
For intq = 0 To List2.ListCount - 1
    fil1(intq) = List2.List(intq)
    fil2(intq) = List3.List(intq)
    fil3(intq) = List4.List(intq)
    fil4(intq) = List5.List(intq)
    fil5(intq) = List6.List(intq)
    fil6(intq) = List7.List(intq)
Next intq

count = List2.ListCount - 1
For intq = 0 To count
   For intw = intq To count
      If fil4(intq) > fil4(intw) Then
         fil1t = fil1(intq)
         fil2t = fil2(intq)
         fil3t = fil3(intq)
         fil4t = fil4(intq)
         fil5t = fil5(intq)
         fil6t = fil6(intq)
         
         fil1(intq) = fil1(intw)
         fil2(intq) = fil2(intw)
         fil3(intq) = fil3(intw)
         fil4(intq) = fil4(intw)
         fil5(intq) = fil5(intw)
         fil6(intq) = fil6(intw)
         
         fil1(intw) = fil1t
         fil2(intw) = fil2t
         fil3(intw) = fil3t
         fil4(intw) = fil4t
         fil5(intw) = fil5t
         fil6(intw) = fil6t
      End If
   Next intw
Next intq
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
For intq = 0 To count
   List2.AddItem fil1(intq)
   List3.AddItem fil2(intq)
   List4.AddItem fil3(intq)
   List5.AddItem fil4(intq)
   List6.AddItem fil5(intq)
   List7.AddItem fil6(intq)
Next intq
End Sub

Private Sub Form_Activate()
Text1(0).Text = Form2.Label5.Caption
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons2_Click()
If List1.ListCount <> 0 Then
  fso.CopyFile "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\Data\info2.mdb", "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\info2.mdb", True
  Adodc4.Refresh
  
  For q = 0 To List1.ListCount - 1
    Adodc4.Refresh
    Adodc4.Recordset.AddNew
    Adodc4.Recordset.Fields!rad = List1.List(q)
    Adodc4.Recordset.Fields!idvam = List2.List(q)
    Adodc4.Recordset.Fields!shhesab = List3.List(q)
    Adodc4.Recordset.Fields!Name = List4.List(q)
    Adodc4.Recordset.Fields!family = List5.List(q)
    Adodc4.Recordset.Fields!Date = List6.List(q)
    Adodc4.Recordset.Fields!phone = List7.List(q)
    Adodc4.Recordset.Update
  Next q
  Adodc4.Refresh
  Form25.Show
End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
List6.ListIndex = List1.ListIndex
List7.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
List6.ListIndex = List2.ListIndex
List7.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
List7.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
List7.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List1.ListIndex = List5.ListIndex
List2.ListIndex = List5.ListIndex
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
List7.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
List1.ListIndex = List6.ListIndex
List2.ListIndex = List6.ListIndex
List3.ListIndex = List6.ListIndex
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
List7.ListIndex = List6.ListIndex
End Sub

Private Sub List7_Click()
List1.ListIndex = List7.ListIndex
List2.ListIndex = List7.ListIndex
List3.ListIndex = List7.ListIndex
List4.ListIndex = List7.ListIndex
List5.ListIndex = List7.ListIndex
List6.ListIndex = List7.ListIndex
End Sub

Private Sub Option1_Click()
Call Amin_1
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
If Form7.Adodc1.Recordset.RecordCount > 0 Then
  Form7.Adodc1.Recordset.MoveFirst
  Do
     If Form7.Adodc1.Recordset.Fields!tasvie = "‰‘œÂ" Then
     
       '¬Œ—Ì‰ ﬁ”ÿ
       a = "0000/00/00"
       b = "0000/00/00"
       Adodc1.RecordSource = "select * from GvamAdi where id='" + Form7.Adodc1.Recordset.Fields!id + "'"
       Adodc1.Refresh
       If Adodc1.Recordset.RecordCount > 0 Then
         Adodc1.Recordset.MoveFirst
         Do
            If Adodc1.Recordset.Fields!saragsat > a Then a = Adodc1.Recordset.Fields!saragsat: b = Adodc1.Recordset.Fields!Date
            Adodc1.Recordset.MoveNext
         Loop Until Adodc1.Recordset.EOF = True
       End If
       
       '¬Ì« ⁄ﬁ» «› «œÂ
       If (a >= Text3.Text) And (a <= Text2.Text) Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc1.Recordset.Fields!id
         List3.AddItem Form7.Adodc1.Recordset.Fields!id1
         List6.AddItem b
         Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc1.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form3.Adodc1.Recordset.Fields!Name
         List5.AddItem Form3.Adodc1.Recordset.Fields!family
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
       End If
     End If
     Form7.Adodc1.Recordset.MoveNext
  Loop Until Form7.Adodc1.Recordset.EOF = True
  Call Amin_2
End If
End Sub

Private Sub Option2_Click()
Call Amin_1
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
If Form7.Adodc2.Recordset.RecordCount > 0 Then
  Form7.Adodc2.Recordset.MoveFirst
  Do
     If Form7.Adodc2.Recordset.Fields!tasvie = "‰‘œÂ" Then
     
       '¬Œ—Ì‰ ﬁ”ÿ
       a = "0000/00/00"
       b = "0000/00/00"
       Adodc2.RecordSource = "select * from GvamAz where id='" + Form7.Adodc2.Recordset.Fields!id + "'"
       Adodc2.Refresh
       If Adodc2.Recordset.RecordCount > 0 Then
         Adodc2.Recordset.MoveFirst
         Do
            If Adodc2.Recordset.Fields!saragsat > a Then a = Adodc2.Recordset.Fields!saragsat: b = Adodc2.Recordset.Fields!Date
            Adodc2.Recordset.MoveNext
         Loop Until Adodc2.Recordset.EOF = True
       End If
       
       '¬Ì« ⁄ﬁ» «› «œÂ
       If (a >= Text3.Text) And (a <= Text2.Text) Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc2.Recordset.Fields!id
         List3.AddItem Form7.Adodc2.Recordset.Fields!id1
         List6.AddItem b
         Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc2.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form3.Adodc1.Recordset.Fields!Name
         List5.AddItem Form3.Adodc1.Recordset.Fields!family
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
       End If
     End If
     Form7.Adodc2.Recordset.MoveNext
  Loop Until Form7.Adodc2.Recordset.EOF = True
  Call Amin_2
End If
End Sub

Private Sub Option3_Click()
Call Amin_1
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
If Form7.Adodc3.Recordset.RecordCount > 0 Then
  Form7.Adodc3.Recordset.MoveFirst
  Do
     If Form7.Adodc3.Recordset.Fields!tasvie = "‰‘œÂ" Then
     
       '¬Œ—Ì‰ ﬁ”ÿ
       a = "0000/00/00"
       b = "0000/00/00"
       Adodc3.RecordSource = "select * from Gvamvig where id='" + Form7.Adodc3.Recordset.Fields!id + "'"
       Adodc3.Refresh
       If Adodc3.Recordset.RecordCount > 0 Then
         Adodc3.Recordset.MoveFirst
         Do
            If Adodc3.Recordset.Fields!saragsat > a Then a = Adodc3.Recordset.Fields!saragsat: b = Adodc3.Recordset.Fields!Date
            Adodc3.Recordset.MoveNext
         Loop Until Adodc3.Recordset.EOF = True
       End If
       
       '¬Ì« ⁄ﬁ» «› «œÂ
       If (a >= Text3.Text) And (a <= Text2.Text) Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc3.Recordset.Fields!id
         List3.AddItem Form7.Adodc3.Recordset.Fields!id1
         List6.AddItem b
         Form4.Adodc1.Recordset.Find "id='" + Form7.Adodc3.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form4.Adodc1.Recordset.Fields!Name
         List5.AddItem Form4.Adodc1.Recordset.Fields!family
         Form3.Adodc1.Recordset.Find "id='" + Form4.Adodc1.Recordset.Fields!idadi + "'", , adSearchForward, 1
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
       End If
     End If
     Form7.Adodc3.Recordset.MoveNext
  Loop Until Form7.Adodc3.Recordset.EOF = True
  Call Amin_2
End If
End Sub
