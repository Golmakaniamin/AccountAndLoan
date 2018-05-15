VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form31 
   BorderStyle     =   0  'None
   Caption         =   "‰„Êœ«—"
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
   LinkTopic       =   "Form31"
   Picture         =   "Form31.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      Max             =   0
      Min             =   -14280
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   9840
      Width           =   13455
   End
   Begin VB.PictureBox Picture1 
      Height          =   7335
      Left            =   960
      RightToLeft     =   -1  'True
      ScaleHeight     =   7275
      ScaleWidth      =   13395
      TabIndex        =   5
      Top             =   2760
      Width           =   13455
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   7455
         Left            =   0
         OleObjectBlob   =   "Form31.frx":E228
         TabIndex        =   6
         Top             =   0
         Width           =   27720
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      MaxLength       =   7
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "1387/02"
      Top             =   2040
      Width           =   855
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   13200
      TabIndex        =   0
      Top             =   10200
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
      MICON           =   "Form31.frx":10C50
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Å—œ«“‘"
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
      MICON           =   "Form31.frx":10C6C
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
      Left            =   11880
      TabIndex        =   4
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "·Ì” "
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
      MICON           =   "Form31.frx":10C88
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
      Caption         =   "„Ã„Ê⁄ «ﬁ”«ÿ —Ê“«‰Â"
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
      Top             =   1680
      Width           =   4215
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qw(31) As Long

Private Sub HScroll1_Change()
MSChart1.Left = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Call HScroll1_Change
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons2_Click()
Dim a As String
Dim s As String
Dim d As String

'Label13.Caption = 0
'Label15.Caption = 0
'Label17.Caption = 0
'Label19.Caption = 0
'
'Label26.Caption = 0
'Label24.Caption = 0
'Label22.Caption = 0
'Label20.Caption = 0
'
'Label29.Caption = 0
'Label31.Caption = 0
'Label33.Caption = 0
'Label35.Caption = 0
'
'Label10.Caption = 0
'Label8.Caption = 0
'Label6.Caption = 0
'Label4.Caption = 0
'
'a = Text2.Text + "/" + "01"
'
'If Left(Form2.Label5.Caption, 7) >= a Then
'  Form8.Adodc1.Recordset.MoveFirst
'  Do
'    If Left(Form8.Adodc1.Recordset.Fields!Date, 7) = a Then
'      Label13.Caption = Val(Label13.Caption) + Val(Form8.Adodc1.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc1.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc1.Recordset.Fields!Date, 7) = a) Then
'      Label26.Caption = Val(Label26.Caption) + Val(Form8.Adodc1.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc1.Recordset.Fields!saragsat, 7) <> a) And (Left(Form8.Adodc1.Recordset.Fields!Date, 7) = a) Then
'      Label29.Caption = Val(Label29.Caption) + Val(Form8.Adodc1.Recordset.Fields!Money)
'    End If
'    Form8.Adodc1.Recordset.MoveNext
'  Loop Until Form8.Adodc1.Recordset.EOF = True
'
'  Form8.Adodc2.Recordset.MoveFirst
'  Do
'    If Left(Form8.Adodc2.Recordset.Fields!Date, 7) = a Then
'      Label15.Caption = Val(Label15.Caption) + Val(Form8.Adodc2.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc2.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc2.Recordset.Fields!Date, 7) = a) Then
'      Label24.Caption = Val(Label24.Caption) + Val(Form8.Adodc2.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc2.Recordset.Fields!saragsat, 7) <> a) And (Left(Form8.Adodc2.Recordset.Fields!Date, 7) = a) Then
'      Label31.Caption = Val(Label31.Caption) + Val(Form8.Adodc2.Recordset.Fields!Money)
'    End If
'    Form8.Adodc2.Recordset.MoveNext
'  Loop Until Form8.Adodc2.Recordset.EOF = True
'
'  Form8.Adodc3.Recordset.MoveFirst
'  Do
'    If Left(Form8.Adodc3.Recordset.Fields!Date, 7) = a Then
'      Label17.Caption = Val(Label17.Caption) + Val(Form8.Adodc3.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc3.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc3.Recordset.Fields!Date, 7) = a) Then
'      Label22.Caption = Val(Label22.Caption) + Val(Form8.Adodc3.Recordset.Fields!Money)
'    End If
'
'    If (Left(Form8.Adodc3.Recordset.Fields!saragsat, 7) <> a) And (Left(Form8.Adodc3.Recordset.Fields!Date, 7) = a) Then
'      Label33.Caption = Val(Label33.Caption) + Val(Form8.Adodc3.Recordset.Fields!Money)
'    End If
'    Form8.Adodc3.Recordset.MoveNext
'  Loop Until Form8.Adodc3.Recordset.EOF = True
'
'  Label19.Caption = Val(Label13.Caption) + Val(Label15.Caption) + Val(Label17.Caption)
'  Label20.Caption = Val(Label22.Caption) + Val(Label24.Caption) + Val(Label26.Caption)
'  Label35.Caption = Val(Label29.Caption) + Val(Label31.Caption) + Val(Label33.Caption)
'End If
'
'  Form7.Adodc1.Recordset.MoveFirst
'  Do
'    If Form7.Adodc1.Recordset.Fields!tasvie = "‰‘œÂ" Then
'      s = Left(Amin.dateaminEzafeMoon(Form7.Adodc1.Recordset.Fields!Date, Form7.Adodc1.Recordset.Fields!numberagsat), 7)
'      d = Left(Form7.Adodc1.Recordset.Fields!Date, 7)
'      If (a <= s) And (a > d) Then
'        Label10.Caption = Val(Label10.Caption) + Val(Form7.Adodc1.Recordset.Fields!moneyg1)
'      End If
'    End If
'    Form7.Adodc1.Recordset.MoveNext
'  Loop Until Form7.Adodc1.Recordset.EOF = True
'
'
'  Form7.Adodc2.Recordset.MoveFirst
'  Do
'    If Form7.Adodc2.Recordset.Fields!tasvie = "‰‘œÂ" Then
'      s = Left(Amin.dateaminEzafeMoon(Form7.Adodc2.Recordset.Fields!Date, Form7.Adodc2.Recordset.Fields!numberagsat), 7)
'      d = Left(Form7.Adodc2.Recordset.Fields!Date, 7)
'      If (a <= s) And (a > d) Then
'        Label8.Caption = Val(Label8.Caption) + Val(Form7.Adodc2.Recordset.Fields!moneyg1)
'      End If
'    End If
'    Form7.Adodc2.Recordset.MoveNext
'  Loop Until Form7.Adodc2.Recordset.EOF = True
'
'  Form7.Adodc3.Recordset.MoveFirst
'  Do
'    If Form7.Adodc3.Recordset.Fields!tasvie = "‰‘œÂ" Then
'      s = Left(Amin.dateaminEzafeMoon(Form7.Adodc3.Recordset.Fields!Date, Form7.Adodc3.Recordset.Fields!numberagsat), 7)
'      d = Left(Form7.Adodc3.Recordset.Fields!Date, 7)
'      If (a <= s) And (a > d) Then
'        Label6.Caption = Val(Label6.Caption) + Val(Form7.Adodc3.Recordset.Fields!moneyg1)
'      End If
'    End If
'    Form7.Adodc3.Recordset.MoveNext
'  Loop Until Form7.Adodc3.Recordset.EOF = True
'  Label4.Caption = Val(Label10.Caption) + Val(Label8.Caption) + Val(Label6.Caption)
'
'Label13.Caption = Amin.moneyaminjoda(Label13.Caption)
'Label15.Caption = Amin.moneyaminjoda(Label15.Caption)
'Label17.Caption = Amin.moneyaminjoda(Label17.Caption)
'Label19.Caption = Amin.moneyaminjoda(Label19.Caption)
'
'Label26.Caption = Amin.moneyaminjoda(Label26.Caption)
'Label24.Caption = Amin.moneyaminjoda(Label24.Caption)
'Label22.Caption = Amin.moneyaminjoda(Label22.Caption)
'Label20.Caption = Amin.moneyaminjoda(Label20.Caption)
'
'Label29.Caption = Amin.moneyaminjoda(Label29.Caption)
'Label31.Caption = Amin.moneyaminjoda(Label31.Caption)
'Label33.Caption = Amin.moneyaminjoda(Label33.Caption)
'Label35.Caption = Amin.moneyaminjoda(Label35.Caption)
'
'Label10.Caption = Amin.moneyaminjoda(Label10.Caption)
'Label8.Caption = Amin.moneyaminjoda(Label8.Caption)
'Label6.Caption = Amin.moneyaminjoda(Label6.Caption)
'Label4.Caption = Amin.moneyaminjoda(Label4.Caption)

'„»«·€ Ã„⁄ ¬Ê—Ì ‘œÂ œ— «Ì‰ —Ê“
MSChart1.Column = 1
For q = 1 To 31
  MSChart1.Row = q
  MSChart1.Data = 0
  qw(q) = 0
  
  If Len(Trim(Str(q))) = 1 Then
    b = "0" + Trim(Str(q))
  Else
    b = Trim(Str(q))
  End If
  a = Text2.Text
  
  Form8.Adodc1.Recordset.MoveFirst
  Do
    If Left(Form8.Adodc1.Recordset.Fields!Date, 7) = a Then
      If Right(Form8.Adodc1.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc1.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
  
  Form8.Adodc2.Recordset.MoveFirst
  Do
    If Left(Form8.Adodc2.Recordset.Fields!Date, 7) = a Then
      If Right(Form8.Adodc2.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc2.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc2.Recordset.MoveNext
  Loop Until Form8.Adodc2.Recordset.EOF = True
  
  Form8.Adodc3.Recordset.MoveFirst
  Do
    If Left(Form8.Adodc3.Recordset.Fields!Date, 7) = a Then
      If Right(Form8.Adodc3.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc3.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc3.Recordset.MoveNext
  Loop Until Form8.Adodc3.Recordset.EOF = True
  
  MSChart1.Row = q
  MSChart1.Data = qw(q)
Next q

'„»«·€ Ã„⁄ ‘œÂ „—»Êÿ »Â «Ì‰ „«Â
MSChart1.Column = 2
For q = 1 To 31
  MSChart1.Row = q
  MSChart1.Data = 0
  qw(q) = 0

  If Len(Trim(Str(q))) = 1 Then
    b = "0" + Trim(Str(q))
  Else
    b = Trim(Str(q))
  End If
  a = Text2.Text

  Form8.Adodc1.Recordset.MoveFirst
  Do
    If (Left(Form8.Adodc1.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc1.Recordset.Fields!Date, 7) = a) Then
      If Right(Form8.Adodc1.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc1.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True

  Form8.Adodc2.Recordset.MoveFirst
  Do
    If (Left(Form8.Adodc2.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc2.Recordset.Fields!Date, 7) = a) Then
      If Right(Form8.Adodc2.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc2.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc2.Recordset.MoveNext
  Loop Until Form8.Adodc2.Recordset.EOF = True

  Form8.Adodc3.Recordset.MoveFirst
  Do
    If (Left(Form8.Adodc3.Recordset.Fields!saragsat, 7) = a) And (Left(Form8.Adodc3.Recordset.Fields!Date, 7) = a) Then
      If Right(Form8.Adodc3.Recordset.Fields!Date, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc3.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc3.Recordset.MoveNext
  Loop Until Form8.Adodc3.Recordset.EOF = True

  MSChart1.Row = q
  MSChart1.Data = qw(q)
Next q

'„»«·€ ¬Ì‰œÂ
MSChart1.Column = 3

For q = 1 To 31
  MSChart1.Row = q
  MSChart1.Data = 0
  qw(q) = 0
  If Len(Trim(Str(q))) = 1 Then
    b = "0" + Trim(Str(q))
  Else
    b = Trim(Str(q))
  End If
  a = Text2.Text + "/" + b
  
  Form7.Adodc1.Recordset.MoveFirst
  Do
    s = Amin.dateaminEktelafmoon(Form7.Adodc1.Recordset.Fields!Date, a)
    If (a >= Form7.Adodc1.Recordset.Fields!Date) And (Val(s) <= Val(Form7.Adodc1.Recordset.Fields!numberagsat)) Then
      d = Amin.dateaminEzafeMoon(Form7.Adodc1.Recordset.Fields!Date, Val(s))
      If Right(d, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form7.Adodc1.Recordset.Fields!moneyg2)
      End If
    End If
    Form7.Adodc1.Recordset.MoveNext
  Loop Until Form7.Adodc1.Recordset.EOF = True


  Form7.Adodc2.Recordset.MoveFirst
  Do
    s = Amin.dateaminEktelafmoon(Form7.Adodc2.Recordset.Fields!Date, a)
    If (a >= Form7.Adodc2.Recordset.Fields!Date) And (Val(s) <= Val(Form7.Adodc2.Recordset.Fields!numberagsat)) Then
      d = Amin.dateaminEzafeMoon(Form7.Adodc2.Recordset.Fields!Date, Val(s))
      If Right(d, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form7.Adodc2.Recordset.Fields!moneyg2)
      End If
    End If
    Form7.Adodc2.Recordset.MoveNext
  Loop Until Form7.Adodc2.Recordset.EOF = True


  Form7.Adodc3.Recordset.MoveFirst
  Do
    s = Amin.dateaminEktelafmoon(Form7.Adodc3.Recordset.Fields!Date, a)
    If (a >= Form7.Adodc3.Recordset.Fields!Date) And (Val(s) <= Val(Form7.Adodc3.Recordset.Fields!numberagsat)) Then
      d = Amin.dateaminEzafeMoon(Form7.Adodc3.Recordset.Fields!Date, Val(s))
      If Right(d, 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form7.Adodc3.Recordset.Fields!moneyg2)
      End If
    End If
    Form7.Adodc3.Recordset.MoveNext
  Loop Until Form7.Adodc3.Recordset.EOF = True

  MSChart1.Row = q
  MSChart1.Data = qw(q)
Next q
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

