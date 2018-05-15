VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form14 
   BackColor       =   &H00C3CEC4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã” ÃÊ"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   3195
   StartUpPosition =   2  'CenterScreen
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3360
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
      Height          =   1860
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
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
      Height          =   1860
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
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
      Left            =   2040
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin KewlButtonz.KewlButtons KewlButtons3 
      Height          =   135
      Left            =   240
      TabIndex        =   5
      Top             =   4920
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
      MICON           =   "Form14.frx":0000
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
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
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
      MICON           =   "Form14.frx":001C
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
      Left            =   240
      TabIndex        =   7
      Top             =   2880
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
      MICON           =   "Form14.frx":0038
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   240
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   240
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
List1.Clear
List2.Clear
List3.Clear
If Form3.Adodc1.Recordset.RecordCount <> 0 Then
  Form3.Adodc1.Recordset.MoveFirst
  Do
     List1.AddItem Form3.Adodc1.Recordset.Fields!id
     List2.AddItem Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
     Form3.Adodc1.Recordset.MoveNext
  Loop Until Form3.Adodc1.Recordset.EOF = True
End If
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
End Sub

Private Sub List3_Click()
For q = 0 To List2.ListCount - 1
    If List2.List(q) = List3.List(List3.ListIndex) Then
       List2.ListIndex = q
       Exit For
    End If
Next q
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

