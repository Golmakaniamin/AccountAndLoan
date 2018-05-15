VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C3CEC4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ê—Êœ »Â ”Ì” „"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3240
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
      Connect         =   $"Form1.frx":13E92
      OLEDBString     =   $"Form1.frx":14046
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "sec"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Œ—ÊÃ"
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
      MICON           =   "Form1.frx":141FA
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Ê—Êœ"
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
      MICON           =   "Form1.frx":14216
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      Height          =   2130
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   6
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form1.frx":14232
      Left            =   240
      List            =   "Form1.frx":1423C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      DataField       =   "date1"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—„“ ⁄»Ê— :"
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ò«—»—«‰ :"
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim w As String, q As Integer

Function mil2shams(miladi_mm_dd_yyyy As String) As String
Dim iran(12), CHRIS(12)
CHRIS(1) = 31: CHRIS(2) = 28: CHRIS(3) = 31
CHRIS(4) = 30: CHRIS(5) = 31: CHRIS(6) = 30
CHRIS(7) = 31: CHRIS(8) = 31: CHRIS(9) = 30
CHRIS(10) = 31: CHRIS(11) = 30: CHRIS(12) = 31
For I = 1 To 12: iran(I) = 31 - (I \ 7) - (I \ 12): Next
mo = Val(Left(miladi_mm_dd_yyyy, 2))
miladi_mm_dd_yyyyy1 = Val(Mid(miladi_mm_dd_yyyy, 4, 2))
Year1 = Val(Mid(miladi_mm_dd_yyyy, 7, 4))
leap1 = Int((Year1 - 1) / 400)
leap2 = Year1 - 1 - 400 * leap1
leap3 = leap2 \ 100
leap4 = leap2 Mod 100
leap5 = leap4 \ 4
CHRIS(2) = 28
If ((Year1 Mod 4) = 0 And (Year1 Mod 100) <> 0) Or _
(Year1 Mod 400) = 0 Then CHRIS(2) = 29
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy1
For I = 1 To mo - 1
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 + CHRIS(I)
Next I
miladi_mm_dd_yyyyy1num = 365 * (Year1 - 1) + _
miladi_mm_dd_yyyyy11 + 97 * leap1 + 24 * leap3 + leap5
miladi_mm_dd_yyyyy1num = miladi_mm_dd_yyyyy1num - 221056!
iry1 = Int(miladi_mm_dd_yyyyy1num / 12053)
iry2 = miladi_mm_dd_yyyyy1num - 12053 * iry1
iry = 33 * iry1 - 16
If iry2 > 365 Then iry = iry + 1: iry2 = iry2 - 365
iry3 = iry2 \ 1461
iry4 = iry2 Mod 1461
iry5 = iry4 \ 365
iry6 = iry4 Mod 365
iry = iry + 1 + 4 * iry3 + iry5
iran(12) = 29
esfand = (8 * iry + 22) / 33 - 0.001
esfand = esfand - Int(esfand)
If esfand > 0.77 Then iran(12) = 30
For I = 1 To 12
If iry6 > iran(I) Then iry6 = iry6 - iran(I) _
Else irm = I: miladi_mm_dd_yyyyy11 = iry6: Exit For
Next I
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 + 5
If miladi_mm_dd_yyyyy11 > iran(irm) Then
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 - iran(irm)
irm = irm + 1
If irm > 12 Then irm = 1: iry = iry + 1
End If
eirmiladi_mm_dd_yyyye = 3 * irm - 3
If irm > 7 Then eirmiladi_mm_dd_yyyye = _
eirmiladi_mm_dd_yyyye - irm + 7
girmiladi_mm_dd_yyyye = (8 * iry + 22) / 33 - 0.001
cirmiladi_mm_dd_yyyye = Int(girmiladi_mm_dd_yyyye) _
+ iry + eirmiladi_mm_dd_yyyye - miladi_mm_dd_yyyyy11 + 3
cirmiladi_mm_dd_yyyye = cirmiladi_mm_dd_yyyye Mod 7
If irm < 10 Then mo = "0" + LTrim(Str(irm)) Else _
mo = LTrim(Str(irm))
If miladi_mm_dd_yyyyy11 < 10 Then d = "0" + _
LTrim(Str(miladi_mm_dd_yyyyy11)) Else _
d = LTrim(Str(miladi_mm_dd_yyyyy11))
mil2shams = LTrim(Str(iry)) + "/" + mo + "/" + d
End Function

Private Sub Form_Activate()
Text1.Text = ""
q = 0
List1.Clear
filenames$ = App.Path & "\UsePas.A@g"
Open filenames$ For Input As #1
Do While Not EOF(1)
  Input #1, w
  List1.AddItem w
Loop
Close #1
End Sub

Private Sub Form_Load()
LoadKeyboardLayout "00000429", 1 ' 00000429 :::::> For Farsi Keyboard
If App.PrevInstance = True Then
  MsgBox "‰—„ «›“«— œ— Õ«·  «Ã—« „Ì »«‘œ", vbCritical + vbMsgBoxRight, ""
  End
End If
End Sub

Private Sub KewlButtons1_Click()
If (Combo1.Text <> "") And (Text1.Text <> "") Then
  If Text1.Text = List1.List(Combo1.ListIndex) Then
    Form2.Label2.Caption = Combo1.Text
    Form2.Label5.Caption = mil2shams(Format(Now, "mm/dd/yyyy"))
    Form2.Label7.Caption = Time$
    Adodc1.Refresh
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!date1 = Form2.Label5.Caption
    Adodc1.Recordset.Fields!time1 = Form2.Label7.Caption
    Adodc1.Recordset.Fields!user = Form2.Label2.Caption
    Adodc1.Recordset.Update
    Form2.Show
    Form1.Hide
'    Form29.Show
  Else
    q = q + 1
    z = MsgBox("—„“ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ" + Chr(10) + "·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ", vbMsgBoxRight + vbCritical, "")
    Call Text1_GotFocus
    If q = 3 Then End
  End If
Else
  z = MsgBox("·ÿ›« ò«—»— —« «‰ Œ«» Ê —„“ ⁄»Ê— ŒÊœ —« Ê«—œ ‰„«ÌÌœ", vbMsgBoxRight + vbCritical, "")
End If
End Sub

Private Sub KewlButtons2_Click()
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 Then
  If Shift = 7 Then
    Form2.Label2.Caption = "„Õ„œ œÊ·  ŒÊ«Â"
    Form2.Label5.Caption = mil2shams(Format(Now, "mm/dd/yyyy"))
    Form2.Label7.Caption = Time$
    Form2.Show
    Form1.Hide
  End If
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 6 Then
  KewlButtons1.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then KewlButtons1.SetFocus
End Sub
