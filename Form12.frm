VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form12 
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
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C3CEC4&
      Height          =   5655
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3240
      Width           =   4215
      Begin VB.Frame Frame3 
         BackColor       =   &H00C3CEC4&
         Caption         =   "‰Ê⁄ ÅÌ«„ò Â«Ì «—”«·Ì"
         Height          =   2775
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   3735
         Begin VB.CheckBox Check9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   "÷«„‰"
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1560
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   "÷«„‰"
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   960
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   "÷«„‰"
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   "ò”—Ì „ÊÃÊœÌ Õ”«» Â«Ì ⁄«œÌ"
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   2160
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   " «ŒÌ— «ﬁ”«ÿ"
            Height          =   495
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   " «ŒÌ— «ﬁ”«ÿ"
            Height          =   495
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C3CEC4&
            Caption         =   " «ŒÌ— «ﬁ”«ÿ"
            Height          =   495
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ê«„ ÊÌéÂ"
            Height          =   495
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ê«„ «÷ÿ—«—Ì"
            Height          =   495
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ê«„ ⁄«œÌ"
            Height          =   495
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C3CEC4&
         Caption         =   "—Ê‘‰ / Œ«„Ê‘"
         Height          =   495
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "24:00"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "00:00"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "1"
         Top             =   960
         Width           =   1335
      End
      Begin KewlButtonz.KewlButtons KewlButtons5 
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   5040
         Width           =   3735
         _ExtentX        =   6588
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
         MICON           =   "Form12.frx":10378
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «"
         Height          =   495
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«“ ”«⁄  "
         Height          =   495
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "›«’·Â »Ì‰ —Ê“Â«Ì «—”«· ÅÌ«„ò"
         Height          =   495
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Height          =   3735
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3720
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Height          =   465
         ItemData        =   "Form12.frx":10394
         Left            =   1080
         List            =   "Form12.frx":103A1
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   6
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   6
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " «ÌÌœ"
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
         MICON           =   "Form12.frx":103CF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò·„Â ⁄»Ê— ÃœÌœ :"
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "·Ì”  ò«—»—«‰ :"
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò·„Â ⁄»Ê— ﬁ»·Ì :"
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      Top             =   9480
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Form12.frx":103EB
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
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   " €ÌÌ— —„“ ⁄»Ê—"
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
      MICON           =   "Form12.frx":10407
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
      Left            =   10440
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   " ‰ŸÌ„«  ÅÌ«„ò"
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
      MICON           =   "Form12.frx":10423
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
      Caption         =   " ‰ŸÌ„« "
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
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub KewlButtons1_Click()
Frame1.Visible = True
Frame2.Visible = False
q = 0
End Sub

Private Sub KewlButtons2_Click()
If Combo1.Text = "" Then
  z = MsgBox("·ÿ›« ‰«„ ò«—»— —« Ê«—œ ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Text1.Text = "") Or (Text2.Text = "") Then
  z = MsgBox("·ÿ›« ›Ì·œ Â« „—»ÊÿÂ —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If (Len(Text1.Text) <> 6) Or (Len(Text2.Text) <> 6) Then
  z = MsgBox("ò·„Â Â«Ì ⁄»Ê— 6 ò«—«ò — «” ", vbCritical + vbMsgBoxRight, "")
  GoTo akhar
End If

If Text1.Text <> Form1.List1.List(Combo1.ListIndex) Then
  q = q + 1
  z = MsgBox("ò·„Â ⁄»Ê— Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ" + Chr(10) + "·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ", vbMsgBoxRight + vbCritical, "")
  Call Text1_GotFocus
  If q = 3 Then End
Else
  Form1.List1.RemoveItem (Combo1.ListIndex)
  Form1.List1.AddItem Text2.Text, Combo1.ListIndex
  
  filenames$ = App.Path & "\UsePas.A@g"
  Open filenames$ For Output As #1
  For d = 0 To Form1.List1.ListCount - 1
    Print #1, Form1.List1.List(d)
  Next d
  Close #1
  z = MsgBox("ò·„Â ⁄»Ê— »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œÂ «” ", vbMsgBoxRight + vbInformation, "")
  Unload Me
  Form2.Hide
  Form1.Show
End If
akhar:
End Sub

Private Sub KewlButtons3_Click()
Frame2.Visible = True
Frame1.Visible = False
List1.Clear
filenames$ = App.Path & "\SMSUsePas.A@G"
Open filenames$ For Input As #1
Do While Not EOF(1)
  Input #1, w
  List1.AddItem w
Loop
Close #1

If List1.List(0) = 1 Then
  Check1.Value = 1
Else
  Check1.Value = 0
End If

Text3.Text = List1.List(1)
Text4.Text = Left(Right(List1.List(2), 6), 5)
Text5.Text = Left(Right(List1.List(3), 6), 5)

If List1.List(4) = 1 Then
  Check2.Value = 1
Else
  Check2.Value = 0
End If

If List1.List(5) = 1 Then
  Check7.Value = 1
Else
  Check7.Value = 0
End If

If List1.List(6) = 1 Then
  Check3.Value = 1
Else
  Check3.Value = 0
End If

If List1.List(7) = 1 Then
  Check8.Value = 1
Else
  Check8.Value = 0
End If

If List1.List(8) = 1 Then
  Check4.Value = 1
Else
  Check4.Value = 0
End If

If List1.List(9) = 1 Then
  Check9.Value = 1
Else
  Check9.Value = 0
End If

If List1.List(10) = 1 Then
  Check5.Value = 1
Else
  Check5.Value = 0
End If

End Sub

Private Sub KewlButtons4_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons5_Click()
List1.RemoveItem (0)
If Check1.Value = 1 Then
  List1.AddItem 1, 0
Else
  List1.AddItem 0, 0
End If

List1.RemoveItem (1)
List1.AddItem Text3.Text, 1

List1.RemoveItem (2)
List1.AddItem "*" + Text4.Text + "*", 2

List1.RemoveItem (3)
List1.AddItem "*" + Text5.Text + "*", 3

List1.RemoveItem (4)
If Check2.Value = 1 Then
  List1.AddItem 1, 4
Else
  List1.AddItem 0, 4
End If

List1.RemoveItem (5)
If Check7.Value = 1 Then
  List1.AddItem 1, 5
Else
  List1.AddItem 0, 5
End If

List1.RemoveItem (6)
If Check3.Value = 1 Then
  List1.AddItem 1, 6
Else
  List1.AddItem 0, 6
End If

List1.RemoveItem (7)
If Check8.Value = 1 Then
  List1.AddItem 1, 7
Else
  List1.AddItem 0, 7
End If

List1.RemoveItem (8)
If Check4.Value = 1 Then
  List1.AddItem 1, 8
Else
  List1.AddItem 0, 8
End If

List1.RemoveItem (9)
If Check9.Value = 1 Then
  List1.AddItem 1, 9
Else
  List1.AddItem 0, 9
End If

List1.RemoveItem (10)
If Check5.Value = 1 Then
  List1.AddItem 1, 10
Else
  List1.AddItem 0, 10
End If

  filenames$ = App.Path & "\SMSUsePas.A@G"
  Open filenames$ For Output As #1
  For d = 0 To List1.ListCount - 1
    Print #1, List1.List(d)
  Next d
  Close #1
  z = MsgBox("‰—„ «›“«— »« „Ê›ﬁÌ   ‰ŸÌ„ ‘œ", vbMsgBoxRight + vbInformation, "")
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 6 Then
  Text2.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) = 6 Then
  KewlButtons2.SetFocus
End If
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons2.SetFocus
End Sub

