VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form13 
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
   LinkTopic       =   "Form13"
   Picture         =   "Form13.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form13.frx":10378
      Left            =   8880
      List            =   "Form13.frx":10385
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C3CEC4&
      Caption         =   "ÌÓÊÌæ"
      Height          =   4695
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   5895
      Begin VB.ListBox List7 
         Height          =   1785
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   375
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
         Height          =   3060
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   1695
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
         Height          =   2760
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
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
         Height          =   2700
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
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
         Height          =   420
         Left            =   4440
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1335
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
         Height          =   420
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin KewlButtonz.KewlButtons KewlButtons6 
         Height          =   135
         Left            =   240
         TabIndex        =   17
         Top             =   3960
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
         MICON           =   "Form13.frx":103B8
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
         Height          =   135
         Left            =   4440
         TabIndex        =   18
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "Form13.frx":103D4
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
         Height          =   135
         Left            =   2040
         TabIndex        =   19
         Top             =   3960
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
         MICON           =   "Form13.frx":103F0
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
         Left            =   3360
         TabIndex        =   24
         ToolTipText     =   "ÇäÊÎÇÈ åãå"
         Top             =   4200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÇäÊÎÇÈ åãå"
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
         MICON           =   "Form13.frx":1040C
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
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         ToolTipText     =   "ÇäÊÎÇÈ åãå"
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÚÏã ÇäÊÎÇÈ åãå"
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
         MICON           =   "Form13.frx":10428
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
         Caption         =   "ÔãÇÑå æÇã"
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "äÇã æ äÇã ÎÇäæÇÏí"
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "äÊíÌå"
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
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "1387/05/01"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "1380/01/01"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Caption         =   "ÇäÊÎÇÈ äæÚ æÇã "
      Height          =   975
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5160
      Width           =   3855
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÚÇÏí"
         Height          =   495
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "ÇÖØÑÇÑí"
         Height          =   495
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "æíŽå"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11280
      TabIndex        =   5
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
      MICON           =   "Form13.frx":10444
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
      Left            =   2640
      TabIndex        =   10
      Top             =   9240
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "Form13.frx":10460
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÇáÊ ÒÇÑÔ :"
      Height          =   495
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇ ÊÇÑíÎ :"
      Height          =   495
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÒ ÊÇÑíÎ :"
      Height          =   495
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "æÇã"
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
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Option1.Value = True
Call Option1_Click
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons3_Click()
For q = 0 To List6.ListCount - 1
  List6.Selected(q) = False
Next q
End Sub

Private Sub KewlButtons6_Click()
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

Private Sub KewlButtons7_Click()
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

Private Sub KewlButtons8_Click()
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

Private Sub KewlButtons9_Click()
For q = 0 To List6.ListCount - 1
  List6.Selected(q) = True
Next q
End Sub

Private Sub List4_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
End Sub

Private Sub List5_Click()
List6.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
List5.ListIndex = List6.ListIndex
'If List5.Selected(List5.ListIndex) = True Then
'  List5.Selected(List5.ListIndex) = False
'  List6.Selected(List6.ListIndex) = False
'Else
'  List5.Selected(List5.ListIndex) = True
'  List6.Selected(List6.ListIndex) = True
'End If
End Sub

Private Sub Option1_Click()
List6.Clear
List7.Clear
List5.Clear
List4.Clear
If Form7.Adodc1.Recordset.RecordCount > 0 Then
  If Combo1.List(Combo1.ListIndex) = "ÊãÇãí æÇã åÇ" Then
    Form7.Adodc1.Recordset.MoveFirst
    Do
      If (Form7.Adodc1.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc1.Recordset.Fields!Date < Text2.Text) Then
        List6.AddItem Form7.Adodc1.Recordset.Fields!id
        List7.AddItem Form7.Adodc1.Recordset.Fields!id1
      End If
      Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
  End If
  
  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÌÇÑí" Then
    Form7.Adodc1.Recordset.MoveFirst
    Do
      If Form7.Adodc1.Recordset.Fields!tasvie = "äÔÏå" Then
        If (Form7.Adodc1.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc1.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc1.Recordset.Fields!id
          List7.AddItem Form7.Adodc1.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
  End If

  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÊÓæíå ÔÏå" Then
    Form7.Adodc1.Recordset.MoveFirst
    Do
      If Form7.Adodc1.Recordset.Fields!tasvie = "ÔÏå" Then
        If (Form7.Adodc1.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc1.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc1.Recordset.Fields!id
          List7.AddItem Form7.Adodc1.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
  End If
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
  If Combo1.List(Combo1.ListIndex) = "ÊãÇãí æÇã åÇ" Then
    Form7.Adodc2.Recordset.MoveFirst
    Do
      If (Form7.Adodc2.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc2.Recordset.Fields!Date < Text2.Text) Then
        List6.AddItem Form7.Adodc2.Recordset.Fields!id
        List7.AddItem Form7.Adodc2.Recordset.Fields!id1
      End If
      Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
  End If
  
  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÌÇÑí" Then
    Form7.Adodc2.Recordset.MoveFirst
    Do
      If Form7.Adodc2.Recordset.Fields!tasvie = "äÔÏå" Then
        If (Form7.Adodc2.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc2.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc2.Recordset.Fields!id
          List7.AddItem Form7.Adodc2.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
  End If

  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÊÓæíå ÔÏå" Then
    Form7.Adodc2.Recordset.MoveFirst
    Do
      If Form7.Adodc2.Recordset.Fields!tasvie = "ÔÏå" Then
        If (Form7.Adodc2.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc2.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc2.Recordset.Fields!id
          List7.AddItem Form7.Adodc2.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
  End If
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
  If Combo1.List(Combo1.ListIndex) = "ÊãÇãí æÇã åÇ" Then
    Form7.Adodc3.Recordset.MoveFirst
    Do
      If (Form7.Adodc3.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc3.Recordset.Fields!Date < Text2.Text) Then
        List6.AddItem Form7.Adodc3.Recordset.Fields!id
        List7.AddItem Form7.Adodc3.Recordset.Fields!id1
      End If
      Form7.Adodc3.Recordset.MoveNext
    Loop Until Form7.Adodc3.Recordset.EOF = True
  End If
  
  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÌÇÑí" Then
    Form7.Adodc3.Recordset.MoveFirst
    Do
      If Form7.Adodc3.Recordset.Fields!tasvie = "äÔÏå" Then
        If (Form7.Adodc3.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc3.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc3.Recordset.Fields!id
          List7.AddItem Form7.Adodc3.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc3.Recordset.MoveNext
    Loop Until Form7.Adodc3.Recordset.EOF = True
  End If

  If Combo1.List(Combo1.ListIndex) = "æÇã åÇí ÊÓæíå ÔÏå" Then
    Form7.Adodc3.Recordset.MoveFirst
    Do
      If Form7.Adodc3.Recordset.Fields!tasvie = "ÔÏå" Then
        If (Form7.Adodc3.Recordset.Fields!Date > Text1.Text) And (Form7.Adodc3.Recordset.Fields!Date < Text2.Text) Then
          List6.AddItem Form7.Adodc3.Recordset.Fields!id
          List7.AddItem Form7.Adodc3.Recordset.Fields!id1
        End If
      End If
      Form7.Adodc3.Recordset.MoveNext
    Loop Until Form7.Adodc3.Recordset.EOF = True
  End If
End If
For q = 0 To List7.ListCount - 1
  Form4.Adodc1.Recordset.Find "id='" & List7.List(q) & "'", , adSearchForward, 1
  List5.AddItem Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
Next q
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
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

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub
