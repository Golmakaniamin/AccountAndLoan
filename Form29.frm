VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form29 
   BackColor       =   &H00B0C4B1&
   BorderStyle     =   0  'None
   Caption         =   "«—”«· ÅÌ«„ò"
   ClientHeight    =   2235
   ClientLeft      =   210
   ClientTop       =   8535
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form29.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   2895
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   480
      Top             =   0
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6840
      Top             =   9360
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
      Connect         =   $"Form29.frx":13E92
      OLEDBString     =   $"Form29.frx":14046
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "datesms"
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
   Begin VB.Timer Timer1 
      Interval        =   9000
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List8 
      Height          =   3165
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "1386/07/07"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3165
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   3165
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   3165
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List5 
      Height          =   3165
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List6 
      Height          =   3165
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List7 
      Height          =   3165
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3960
      Top             =   9480
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
      Connect         =   $"Form29.frx":141FA
      OLEDBString     =   $"Form29.frx":143AE
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
      Left            =   5280
      Top             =   9480
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
      Connect         =   $"Form29.frx":14562
      OLEDBString     =   $"Form29.frx":14716
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
      Left            =   120
      Top             =   9480
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
      Connect         =   $"Form29.frx":148CA
      OLEDBString     =   $"Form29.frx":14A7E
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   1440
      Top             =   9480
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
      Connect         =   $"Form29.frx":14C32
      OLEDBString     =   $"Form29.frx":14DE6
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "sendsms"
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
   Begin MSCommLib.MSComm Comm1 
      Left            =   5040
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      ParityReplace   =   32
      RThreshold      =   1
      RTSEnable       =   -1  'True
   End
   Begin VB.Label lblErrors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Error && general Messages"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   2910
   End
   Begin VB.Label lblGSMStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Good"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   20
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSM Status :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      DataField       =   "date1"
      DataSource      =   "Adodc4"
      Height          =   495
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      DataField       =   "number"
      DataSource      =   "Adodc5"
      Height          =   375
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   11160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ Ê«„"
      Height          =   495
      Index           =   2
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   10200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
      Height          =   495
      Index           =   3
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Ê«„"
      Height          =   495
      Index           =   4
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      Height          =   495
      Index           =   5
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   10200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Œ«‰Ê«œêÌ"
      Height          =   495
      Index           =   6
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ¬Œ—Ì‰ Å—œ«Œ "
      Height          =   495
      Index           =   7
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   10200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰ Â„—«Â"
      Height          =   495
      Index           =   8
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   10200
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Boolean
Public Ring As Boolean
Public Error As Boolean
Public Greater_Sign As Boolean
Public Message_Store As Boolean
Public Message_Buffer As String
Public SMS_TelNumber As String
Public SMS_MsgNumber As String
Public SMS_Message As String
Dim SMS_Break() As String
Dim SMS_Header() As String
Dim I As Integer
Dim p1 As Boolean
Dim str1 As String
Dim str2 As String
Dim fso As New FileSystemObject
Dim a As String, b As String, e As String, f As String
Dim strtext3 As String, p As Boolean
Dim smscount As Long

Private Sub Amin_1()
Dim stry As String, strm As String
List1.Clear
List2.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
strm = Val(Mid(Text1(0).Text, 6, 2)) - 1
stry = Mid(Text1(0).Text, 1, 4)
If strm = 0 Then
  strm = 12
  stry = stry - 1
End If
If Len(strm) < 2 Then strm = "0" + strm
Text2.Text = stry + "/" + strm + Mid(Text1(0).Text, 8, 3)
End Sub

Private Sub Amin_2()
Text1(0).Text = Form2.Label5.Caption
Call Amin_1
List1.Clear
List2.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
If Form2.List1.List(4) = 1 Then
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
       If a < Text2.Text Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc1.Recordset.Fields!id
         List6.AddItem b
         Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc1.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form3.Adodc1.Recordset.Fields!Name
         List5.AddItem Form3.Adodc1.Recordset.Fields!family
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
         List8.AddItem "Adi"
       End If
     End If
     Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
  End If
End If

If Form2.List1.List(6) = 1 Then
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
       If a < Text2.Text Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc2.Recordset.Fields!id
         List6.AddItem b
         Form3.Adodc1.Recordset.Find "id='" + Form7.Adodc2.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form3.Adodc1.Recordset.Fields!Name
         List5.AddItem Form3.Adodc1.Recordset.Fields!family
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
         List8.AddItem "Ezterari"
       End If
     End If
     Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
  End If
End If

If Form2.List1.List(8) = 1 Then
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
       If a < Text2.Text Then
         List1.AddItem List1.ListCount + 1
         List2.AddItem Form7.Adodc3.Recordset.Fields!id
         List6.AddItem b
         Form4.Adodc1.Recordset.Find "id='" + Form7.Adodc3.Recordset.Fields!id1 + "'", , adSearchForward, 1
         List4.AddItem Form4.Adodc1.Recordset.Fields!Name
         List5.AddItem Form4.Adodc1.Recordset.Fields!family
         Form3.Adodc1.Recordset.Find "id='" + Form4.Adodc1.Recordset.Fields!idadi + "'", , adSearchForward, 1
         List7.AddItem Form3.Adodc1.Recordset.Fields!mobile
         List8.AddItem "Vige"
       End If
     End If
     Form7.Adodc3.Recordset.MoveNext
    Loop Until Form7.Adodc3.Recordset.EOF = True
  End If
End If

If Form2.List1.List(0) = 1 Then
  For q = 0 To List1.ListCount - 1
    If Len(List7.List(q)) = 11 Then
      If List6.List(q) <> "0000/00/00" Then
        strtext3 = "Ba salam" + Chr(13) + Chr(10) + "Aghsat Vam " + List8.List(q) + " shomare " + List2.List(q) + " Takhir darad Lotfan Nesbat be pardakht on eghdam farmaeed" + Chr(13) + Chr(10) + "Sandog Karim Ahl Beyt"
      Else
        strtext3 = "Ba salam" + Chr(13) + Chr(10) + "Az aghsat Vam " + List8.List(q) + " shomare " + List2.List(q) + " Hanoz moredi sabt Nashode ast Nesbat be pardakht on eghdam farmaeed" + Chr(13) + Chr(10) + "Sandog Karim Ahl Beyt"
      End If
      Adodc5.Recordset.AddNew
      Adodc5.Recordset.Fields!number = List7.List(q)
      Adodc5.Recordset.Fields!promp = strtext3
      Adodc5.Recordset.Fields!no = "œÌ— ò—œ «ﬁ”«ÿ"
      Adodc5.Recordset.Fields!no1 = "1"
      Adodc5.Recordset.Fields!user = Form2.Label2.Caption
      Adodc5.Recordset.Fields!Time = Form2.Label7.Caption
      Adodc5.Recordset.Fields!Date = Form2.Label5.Caption
      Adodc5.Recordset.Fields!send = "0"
      Adodc5.Recordset.Fields!delivery = "0"
      Adodc5.Recordset.Update
    End If
  Next q
End If

If Form2.List1.List(10) = 1 Then
  If Form3.Adodc1.Recordset.RecordCount > 0 Then
    Form3.Adodc1.Recordset.MoveFirst
    Do
      If (Len(Form3.Adodc1.Recordset.Fields!mobile) = 11) And Form3.Adodc1.Recordset.Fields!Delete = "‰‘œÂ" Then
        e = Amin.dateaminEktelafmoon(Form3.Adodc1.Recordset.Fields!edate, Text1(0).Text)
        If Val(Form3.Adodc1.Recordset.Fields!Money) < (Val(e) * 20000) Then
          f = (Val(e) * 20000) - Val(Form3.Adodc1.Recordset.Fields!Money)
          strtext3 = "Ba salam" + Chr(13) + Chr(10) + "Hesab adi shoma mablagh " + f + " Rial kasr darad nesbat be takmil on eghdam farmaeed" + Chr(13) + Chr(10) + "Sandog Karim Ahl Beyt"
          Adodc5.Recordset.AddNew
          Adodc5.Recordset.Fields!number = Form3.Adodc1.Recordset.Fields!mobile
          Adodc5.Recordset.Fields!promp = strtext3
          Adodc5.Recordset.Fields!no = "ò”—Ì Õﬁ ⁄÷ÊÌ "
          Adodc5.Recordset.Fields!no1 = "2"
          Adodc5.Recordset.Fields!user = Form2.Label2.Caption
          Adodc5.Recordset.Fields!Time = Form2.Label7.Caption
          Adodc5.Recordset.Fields!Date = Form2.Label5.Caption
          Adodc5.Recordset.Fields!send = "0"
          Adodc5.Recordset.Fields!delivery = "0"
          Adodc5.Recordset.Update
          q = q + 1
        End If
      End If
      Form3.Adodc1.Recordset.MoveNext
    Loop Until Form3.Adodc1.Recordset.EOF = True
  End If
End If
End Sub

Private Sub Comm1_OnComm()
Static CEvent As String
Dim CChar As String * 1
Select Case Comm1.CommEvent

    Case comEvReceive
    
        Do
            CChar = Comm1.Input
            If Message_Store Then
                Message_Buffer = Message_Buffer & CChar
            End If
            Select Case CChar
                Case ">"
                     Greater_Sign = True
                     List3.AddItem CChar
                Case vbLf
                
                Case vbCr
                    If Len(CEvent) > 0 Then
                      Handle_Rec_Event CEvent
                      CEvent = ""
                    End If
                Case Else
                    CEvent = CEvent + CChar
            End Select
        
        Loop While Comm1.InBufferCount
End Select

End Sub

Private Sub cmdReadParticularSMS_Click()
Call Get_Message(InputBox("Msg No:"))
    MSF1.Row = MSF1.Rows - 1
    MSF1.Col = 0
    MSF1.Text = MSF1.Row
    MSF1.Col = 1
    MSF1.Text = SMS_TelNumber
    MSF1.Col = 2
    MSF1.Text = SMS_Message
    MSF1.Rows = MSF1.Rows + 1

End Sub
Private Sub Initialise_Modem()
On Error GoTo X
' small Suggestion , please check your own GSM settings
Comm1.CommPort = 2
Comm1.Settings = "9600,n,8,1"
Comm1.Handshaking = comNone
' Other wise it will make you MAD

If Comm1.PortOpen = False Then
    Comm1.PortOpen = True
    Comm1.DTREnable = True
    Comm1.RTSEnable = True
    Comm1.RThreshold = 1
    Comm1.InputLen = 1
    OK = False
    Error = False
    Comm1.Output = "AT" & vbCrLf
    Wait_For_Response
    If Not OK Then
        Comm1.PortOpen = False
        lblGSMStatus.Caption = "Modem is not responding, check the connection"
        Exit Sub
    End If
    Comm1.Output = "ATE0" & vbCrLf
    Wait_For_Response
    Else
        lblGSMStatus.Caption = "Port Already Open"
    End If
    lblGSMStatus.Caption = "Connected"
    Exit Sub
X:
    If Err.number = 8005 Then
        lblGSMStatus.Caption = "Port is already opened by another programme, please check"
    ElseIf Err.number = 8002 Then
        lblGSMStatus.Caption = "Invalid Port"
    Exit Sub
    End If
End Sub

Private Sub Wait_For_Response()
Dim Start

   Start = Timer
   Do While Timer < Start + 8
      DoEvents
      If OK Then
        Exit Sub
      End If
      If Error Then
        Exit Sub
      End If
   Loop
End Sub


Private Sub Handle_Rec_Event(CEvent As String)
Dim Temp As Variant
List3.AddItem CEvent
If Mid(CEvent, 1, 5) = "+CMTI" Then
    lblErrors.Caption = "SMS Received"
    Temp = Split(CEvent, ",")
    SMS_MsgNumber = Temp(1)
    Debug.Print Temp(1)
    SMS_TelNumber = ""
    SMS_Message = ""
    Message_Buffer = ""
    CEvent = ""
    Get_Message SMS_MsgNumber
    OK = False
    Error = False
    'Delete_Message SMS_MsgNumber
    'Store_Message_In_Database SMS_TelNumber, SMS_Message
    
    MSF1.Row = MSF1.Rows - 1
    MSF1.Col = 0
    MSF1.Text = SMS_MsgNumber
    MSF1.Col = 1
    MSF1.Text = SMS_TelNumber
    MSF1.Col = 2
    MSF1.Text = SMS_Message
    MSF1.Rows = MSF1.Rows + 1

    Exit Sub
End If
Select Case CEvent
    Case "OK"
        OK = True
    Case "ERROR"
        Error = True
    Case "RING"
        If Ring = False Then
            Ring = True
        End If
    Case Else
        ' catch the ERRORS
        ' like Memory errors
        '
End Select
End Sub


Private Sub Get_Message(ByVal MsgNo As String)
SMS_TelNumber = ""
SMS_Message = ""
Message_Buffer = ""
OK = False
Error = False
Comm1.Output = "AT+CMGR=" & MsgNo & vbCrLf
While Not OK Or Error
  Message_Store = True
  DoEvents
  Wait_For_Response
Wend
If OK Then
   Read_Message
   lblErrors.Caption = "SMS Read"
End If
If Error Then
   lblErrors.Caption = "Bad Format, could not read "
End If
End Sub

Private Sub Read_Message()
If Parse_SMS Then
   SMS_Break = Split(Message_Buffer, vbCrLf, , vbTextCompare)
   SMS_Header = Split(SMS_Break(0), ",", , vbTextCompare)
   SMS_TelNumber = Mid(Right(SMS_Header(1), 11), 1, 10)
   Message_Buffer = ""
   For I = 1 To UBound(SMS_Break(), 1)
       Message_Buffer = Message_Buffer & SMS_Break(I) & vbCrLf
   Next I
   SMS_Message = Message_Buffer
   lblErrors.Caption = "Decoding"
Else
   lblErrors.Caption = "Unable to decode Message"
End If
End Sub

Public Function Parse_SMS() As Boolean
Dim StartPoint As Long
Dim EndPoint As Long
Dim Buffer1 As String
Dim Buffer2 As String
Buffer1 = Message_Buffer
StartPoint = InStr(1, Buffer1, "+CMGR:", vbTextCompare)
EndPoint = InStr(1, Buffer1, vbCrLf & "OK", vbTextCompare)
If StartPoint <> 0 And EndPoint > StartPoint Then
   I = StartPoint
   While I < EndPoint
    Buffer2 = Buffer2 & Mid(Buffer1, I, 1)
    I = I + 1
   Wend
   Parse_SMS = True
   Message_Buffer = Buffer2
   Exit Function
End If
Parse_SMS = False
End Function

Private Sub Delete_Message(ByVal MsgNo As String)
Comm1.Output = "AT+CMGD=" & MsgNo & vbCrLf
  While Not Greater_Sign
    DoEvents
    Wait_For_Response
  Wend
If Error Then
   lblErrors.Caption = "Could not delete"
ElseIf OK Then
   lblErrors.Caption = "Message Deleted"
End If
End Sub


Public Function Send_Message(ByVal MobNumber As String, ByVal MsgText As String) As Boolean
Greater_Sign = False
Comm1.Output = "AT+CMGF=1" & vbCrLf
Comm1.Output = "AT+CSCA=" & Chr(34) & "+9891100500" & Chr(34) & vbCrLf
Comm1.Output = "AT+CMGS=" & Chr(34) & Trim(MobNumber) & Chr(34) & vbCrLf
While Not Greater_Sign
    DoEvents
    Wait_For_Response
Wend
If Greater_Sign Then
    Comm1.Output = Trim(MsgText) & Chr(26) & vbCrLf
    OK = False
    Error = False
    While Not OK Or Error
        DoEvents
        Wait_For_Response
    Wend
    If OK Then
        lblErrors.Caption = "Message Sent"
        Send_Message = True
    Else
        lblErrors.Caption = "Message Not Sent"
        Send_Message = False
    End If
Else
    lblErrors.Caption = "Message cannot be sent"
    Send_Message = False
End If

End Function

Private Sub Command1_Click()
'For q = 1 To 50
'  Call Get_Message(q)
'  Comm1.Output = "AT+CMGF=1" & vbCrLf
'  If SMS_Message <> "" Then
'    Form27.Adodc2.Recordset.AddNew
'    Form27.Adodc2.Recordset.Fields!number = SMS_TelNumber
'    Form27.Adodc2.Recordset.Fields!promp = SMS_Message
'    Form27.Adodc2.Recordset.Fields!user = Form2.Label2.Caption
'    Form27.Adodc2.Recordset.Fields!Time = Form2.Label7.Caption
'    Form27.Adodc2.Recordset.Fields!Date = Form2.Label5.Caption
'    Form27.Adodc2.Recordset.Update
''    Call Delete_Message(2)
'  End If
'Next q
End Sub

Private Sub Form_Activate()
Text1(0).Text = Form2.Label5.Caption
'Me.Hide
End Sub

Private Sub Form_Load()
  Message_Store = False
  Call Initialise_Modem
'  Dim nid As NOTIFYICONDATA
'  With nid
'    .cbSize = Len(nid) ' size of this structure
'    .hWnd = Me.hWnd ' memory location(handle) for the processor of its message and icon
'    .uID = 0 ' the unique ID for the icon. must be different from the other sys tray icon
'    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP ' notify of message, display icon, display tooltip
'    .uCallbackMessage = 1400 ' message used to be notified when there's event. any number greater than 1300 will do
'    .hIcon = Form1.Icon      ' assign the icon to the form's icon
'    .szTip = "Tooltip here" & vbNullChar ' terminate the string with vbNullChar or Chr(0)
'  End With
'  ' Shell_NotifyIconA ID_OF_ICON, NOTIFYICONDATA
'  Shell_NotifyIconA NIM_ADD, nid
'  ' oldproc is the address(memory location) of the original window procedure
'  oldproc = SetWindowLongA(Me.hWnd, -4, AddressOf proc)
'  formloaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  Dim nid As NOTIFYICONDATA
'  With nid
'    .hWnd = Me.hWnd
'    .cbSize = Len(nid) ' size of structure
'    .uID = 0 ' the id we set earlier
'  End With
'  ' delete the icon!
'  Shell_NotifyIconA NIM_DELETE, nid
'  SetWindowLongA Me.hWnd, -4, oldproc
End Sub


Private Sub Timer1_Timer()
Text1(0).Text = Form2.Label5.Caption
If Adodc4.Recordset.RecordCount > 0 Then
  Adodc4.Recordset.Sort = "date1"
  Adodc4.Recordset.MoveLast
  If Amin.dateaminEzafeday(Adodc4.Recordset.Fields!date1, Val(Form2.List1.List(1))) <= Text1(0).Text Then
    Adodc4.Recordset.AddNew
    Adodc4.Recordset.Fields!date1 = Text1(0).Text
    Adodc4.Recordset.Update
    Call Amin_2
  End If
Else
  Adodc4.Recordset.AddNew
  Adodc4.Recordset.Fields!date1 = Text1(0).Text
  Adodc4.Recordset.Update
  Call Amin_2
End If

If InStr(List3.List(List3.ListCount - 1), "ERROR") <> 0 Then
  Timer2.Enabled = False
End If

If (List3.List(List3.ListCount - 1) = "OK") And (Timer2.Enabled = False) Then
  Adodc5.CommandType = adCmdUnknown
  Adodc5.RecordSource = "select * from sendsms where send='0'"
  Adodc5.Refresh
  If (Adodc5.Recordset.RecordCount > 0) Then
    Adodc5.Recordset.Sort = "q"
    Adodc5.Recordset.MoveFirst
    smscount = Adodc5.Recordset.Fields!q
    Timer2.Enabled = True
  End If
End If
If (List3.List(List3.ListCount - 1) <> "OK") And (Timer2.Enabled = False) Then
  Unload Me
  Form29.Show
End If
End Sub

Private Sub Timer2_Timer()
p1 = False
If Adodc5.Recordset.RecordCount > 0 Then
  Adodc5.Recordset.MoveFirst
  Do
    If Adodc5.Recordset.Fields!q = smscount Then
      p1 = True
      str1 = Adodc5.Recordset.Fields!number
      str2 = Adodc5.Recordset.Fields!promp
      Call Send_Message(str1, str2)
      If List3.List(List3.ListCount - 1) <> "ERROR" Then
        Adodc5.Recordset.Fields!send = 1
      End If
      smscount = smscount + 1
      Exit Do
    End If
    Adodc5.Recordset.MoveNext
  Loop Until Adodc5.Recordset.EOF = True
End If
If p1 = False Then
  Timer2.Enabled = False
End If
End Sub
