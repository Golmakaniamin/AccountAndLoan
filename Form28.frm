VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form28 
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
   LinkTopic       =   "Form28"
   Picture         =   "Form28.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   5235
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4440
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   5640
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   9
      Left            =   7560
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form28.frx":10378
      Left            =   10080
      List            =   "Form28.frx":10385
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   495
      Left            =   11040
      TabIndex        =   0
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
      MICON           =   "Form28.frx":103B3
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
      Height          =   375
      Left            =   2280
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   $"Form28.frx":103CF
      OLEDBString     =   $"Form28.frx":10583
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   "pratic1"
      RecordSource    =   "sec"
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form28.frx":10737
      Height          =   5295
      Left            =   8760
      TabIndex        =   5
      Top             =   3960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   28
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÊÇÑíÎ æ ÒãÇä åÇí æÑæÏ Èå ÓíÓÊã"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "date1"
         Caption         =   "ÊÇÑíÎ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "user"
         Caption         =   "˜ÇÑÈÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "time1"
         Caption         =   "ÒãÇä"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1874.835
         EndProperty
      EndProperty
   End
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3240
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
      MICON           =   "Form28.frx":1074C
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
      Caption         =   "ÚãáíÇÊ ÇäÌÇã ÔÏå ÇÒ Óæí ÔÎÕ ÏÑ ÊÇÑíÎ ÇäÊÎÇÈ ÔÏå"
      Height          =   375
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇ"
      Height          =   495
      Index           =   1
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇÑíÎ : ÇÒ"
      Height          =   495
      Index           =   9
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "áíÓÊ ˜ÇÑÈÑÇä :"
      Height          =   495
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      DataField       =   "date1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Úãá˜ÑÏ ˜ÇÑÈÑÇä"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1(9).SetFocus
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
List1.Clear
If LastRow > 0 Then
  DataGrid1.Col = 0
  'ÍÓÇÈ ÚÇÏí
  If Form3.Adodc1.Recordset.RecordCount > 0 Then
    Form3.Adodc1.Recordset.MoveFirst
    Do
      If (Form3.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form3.Adodc1.Recordset.Fields!edate = DataGrid1.Text) Then
        List1.AddItem "ÇÝÊÊÇÍ ÍÓÇÈ ÚÇÏí ÂÞÇí " + Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family + " Èå ÔãÇÑå " + Form3.Adodc1.Recordset.Fields!id
      End If
      Form3.Adodc1.Recordset.MoveNext
    Loop Until Form3.Adodc1.Recordset.EOF = True
  End If
  'ÍÓÇÈ æíŽå
  If Form4.Adodc1.Recordset.RecordCount > 0 Then
    Form4.Adodc1.Recordset.MoveFirst
    Do
      If (Form4.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form4.Adodc1.Recordset.Fields!edate = DataGrid1.Text) Then
        List1.AddItem "ÇÝÊÊÇÍ ÍÓÇÈ æíŽå ÂÞÇí " + Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family + " Èå ÔãÇÑå " + Form4.Adodc1.Recordset.Fields!id
      End If
      Form4.Adodc1.Recordset.MoveNext
    Loop Until Form4.Adodc1.Recordset.EOF = True
  End If
  'ãæÌæÏí ÍÓÇÈ ÚÇÏí
  If Form5.Adodc1.Recordset.RecordCount > 0 Then
    Form5.Adodc1.Recordset.MoveFirst
    Do
      If (Form5.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form5.Adodc1.Recordset.Fields!Date = DataGrid1.Text) Then
        If Form5.Adodc1.Recordset.Fields!Amal = "ÇÝÒÇíÔ" Then
          Form3.Adodc1.Recordset.Find "id='" + Form5.Adodc1.Recordset.Fields!id + "'", , adSearchForward, 1
          q = Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
          List1.AddItem "ÇÝÒÇíÔ ãÈáÛ " + Trim(Str(Form5.Adodc1.Recordset.Fields!Money)) + " ÑíÇá Èå ÍÓÇÈ ÚÇÏí ÂÞÇí " + q
        Else
          Form3.Adodc1.Recordset.Find "id='" + Form5.Adodc1.Recordset.Fields!id + "'", , adSearchForward, 1
          q = Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
          List1.AddItem "˜ÓÑ ãÈáÛ " + Trim(Str(Form5.Adodc1.Recordset.Fields!Money)) + " ÑíÇá ÇÒ ÍÓÇÈ ÚÇÏí ÂÞÇí " + q
        End If
      End If
      Form5.Adodc1.Recordset.MoveNext
    Loop Until Form5.Adodc1.Recordset.EOF = True
  End If
  'ãæÌæÏí ÍÓÇÈ æíŽå
  If Form5.Adodc2.Recordset.RecordCount > 0 Then
    Form5.Adodc2.Recordset.MoveFirst
    Do
      If (Form5.Adodc2.Recordset.Fields!user = Combo1.Text) And (Form5.Adodc2.Recordset.Fields!Date = DataGrid1.Text) Then
        If Form5.Adodc2.Recordset.Fields!Amal = "ÇÝÒÇíÔ" Then
          Form4.Adodc1.Recordset.Find "id='" + Form5.Adodc2.Recordset.Fields!id + "'", , adSearchForward, 1
          q = Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
          List1.AddItem "ÇÝÒÇíÔ ãÈáÛ " + Trim(Str(Form5.Adodc2.Recordset.Fields!Money)) + " ÑíÇá Èå ÍÓÇÈ æíŽå ÂÞÇí " + q
        Else
          Form4.Adodc1.Recordset.Find "id='" + Form5.Adodc2.Recordset.Fields!id + "'", , adSearchForward, 1
          q = Form4.Adodc1.Recordset.Fields!Name + " " + Form4.Adodc1.Recordset.Fields!family
          List1.AddItem "˜ÓÑ ãÈáÛ " + Trim(Str(Form5.Adodc2.Recordset.Fields!Money)) + " ÑíÇá ÇÒ ÍÓÇÈ æíŽå ÂÞÇí " + q
        End If
      End If
      Form5.Adodc2.Recordset.MoveNext
    Loop Until Form5.Adodc2.Recordset.EOF = True
  End If
  'ÏÑÎæÇÓÊ æÇã ÚÇÏí
  If Form6.Adodc1.Recordset.RecordCount > 0 Then
    Form6.Adodc1.Recordset.MoveFirst
    Do
      If (Form6.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form6.Adodc1.Recordset.Fields!datet = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÏÑÎæÇÓÊ æÇã ÚÇÏí ÔãÇÑå " + Form6.Adodc1.Recordset.Fields!id
      End If
      Form6.Adodc1.Recordset.MoveNext
    Loop Until Form6.Adodc1.Recordset.EOF = True
  End If
  'ÏÑÎæÇÓÊ æÇã ÇÖØÑÇÑí
  If Form6.Adodc2.Recordset.RecordCount > 0 Then
    Form6.Adodc2.Recordset.MoveFirst
    Do
      If (Form6.Adodc2.Recordset.Fields!user = Combo1.Text) And (Form6.Adodc2.Recordset.Fields!datet = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÏÑÎæÇÓÊ æÇã ÇÖØÑÇÑí ÔãÇÑå " + Form6.Adodc2.Recordset.Fields!id
      End If
      Form6.Adodc2.Recordset.MoveNext
    Loop Until Form6.Adodc2.Recordset.EOF = True
  End If
  'ÏÑÎæÇÓÊ æÇã æíŽå
  If Form6.Adodc3.Recordset.RecordCount > 0 Then
    Form6.Adodc3.Recordset.MoveFirst
    Do
      If (Form6.Adodc3.Recordset.Fields!user = Combo1.Text) And (Form6.Adodc3.Recordset.Fields!datet = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÏÑÎæÇÓÊ æÇã æíŽå ÔãÇÑå " + Form6.Adodc3.Recordset.Fields!id
      End If
      Form6.Adodc3.Recordset.MoveNext
    Loop Until Form6.Adodc3.Recordset.EOF = True
  End If
  'ÑÏÇÎÊ æÇã ÚÇÏí
  If Form7.Adodc1.Recordset.RecordCount > 0 Then
    Form7.Adodc1.Recordset.MoveFirst
    Do
      If (Form7.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form7.Adodc1.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ÑÏÇÎÊ æÇã ÚÇÏí ÔãÇÑå " + Form7.Adodc1.Recordset.Fields!id + ""
      End If
      Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
  End If
  'ÑÏÇÎÊ æÇã ÇÖØÑÇÑí
  If Form7.Adodc2.Recordset.RecordCount > 0 Then
    Form7.Adodc2.Recordset.MoveFirst
    Do
      If (Form7.Adodc2.Recordset.Fields!user = Combo1.Text) And (Form7.Adodc2.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ÑÏÇÎÊ æÇã ÇÖØÑÇÑí ÔãÇÑå " + Form7.Adodc2.Recordset.Fields!id + ""
      End If
      Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
  End If
  'ÑÏÇÎÊ æÇã æíŽå
  If Form7.Adodc3.Recordset.RecordCount > 0 Then
    Form7.Adodc3.Recordset.MoveFirst
    Do
      If (Form7.Adodc3.Recordset.Fields!user = Combo1.Text) And (Form7.Adodc3.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ÑÏÇÎÊ æÇã æíŽå ÔãÇÑå " + Form7.Adodc3.Recordset.Fields!id + ""
      End If
      Form7.Adodc3.Recordset.MoveNext
    Loop Until Form7.Adodc3.Recordset.EOF = True
  End If
  'ÏÑíÇÝÊ ÇÞÓÇØ æÇã ÚÇÏí
  If Form8.Adodc1.Recordset.RecordCount > 0 Then
    Form8.Adodc1.Recordset.MoveFirst
    Do
      If (Form8.Adodc1.Recordset.Fields!user = Combo1.Text) And (Form8.Adodc1.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÞÓØ ÔãÇÑå " + Trim(Str(Form8.Adodc1.Recordset.Fields!rad)) + " æÇã ÚÇÏí ÔãÇÑå " + Form8.Adodc1.Recordset.Fields!id
      End If
      Form8.Adodc1.Recordset.MoveNext
    Loop Until Form8.Adodc1.Recordset.EOF = True
  End If
  'ÏÑíÇÝÊ ÇÞÓÇØ æÇã ÇÖØÑÇÑí
  If Form8.Adodc2.Recordset.RecordCount > 0 Then
    Form8.Adodc2.Recordset.MoveFirst
    Do
      If (Form8.Adodc2.Recordset.Fields!user = Combo1.Text) And (Form8.Adodc2.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÞÓØ ÔãÇÑå " + Trim(Str(Form8.Adodc2.Recordset.Fields!rad)) + " æÇã ÇÖØÑÇÑí ÔãÇÑå " + Form8.Adodc2.Recordset.Fields!id
      End If
      Form8.Adodc2.Recordset.MoveNext
    Loop Until Form8.Adodc2.Recordset.EOF = True
  End If
  'ÏÑíÇÝÊ ÇÞÓÇØ æÇã æíŽå
  If Form8.Adodc3.Recordset.RecordCount > 0 Then
    Form8.Adodc3.Recordset.MoveFirst
    Do
      If (Form8.Adodc3.Recordset.Fields!user = Combo1.Text) And (Form8.Adodc3.Recordset.Fields!Date = DataGrid1.Text) Then
        List1.AddItem "ËÈÊ ÞÓØ ÔãÇÑå " + Trim(Str(Form8.Adodc3.Recordset.Fields!rad)) + " æÇã æíŽå ÔãÇÑå " + Form8.Adodc3.Recordset.Fields!id
      End If
      Form8.Adodc3.Recordset.MoveNext
    Loop Until Form8.Adodc3.Recordset.EOF = True
  End If

End If
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Form28.Hide
End Sub

Private Sub KewlButtons2_Click()
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from sec where (date1>='" + Text1(9).Text + "' and date1<='" + Text1(0).Text + "') and user='" + Combo1.Text + "'"
Adodc1.Refresh
Adodc1.Recordset.Sort = "date1,time1"
DataGrid1.Refresh
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 9 Then If KeyCode = 13 Then Text1(0).SetFocus
If Index = 0 Then If KeyCode = 13 Then KewlButtons2.SetFocus
End Sub
