VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form16 
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
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons KewlButtons2 
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1800
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
      MICON           =   "Form16.frx":EBD5
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
      Left            =   13440
      TabIndex        =   3
      Top             =   10440
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
      MICON           =   "Form16.frx":EBF1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2520
      MaxLength       =   4
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1387"
      Top             =   1800
      Width           =   615
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7935
      Left            =   960
      OleObjectBlob   =   "Form16.frx":EC0D
      TabIndex        =   38
      Top             =   2400
      Width           =   13695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   12600
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "«ﬁ”«ÿ „—»Êÿ »Â „«Â œÌê— òÂ œ— «Ì‰ „«Â Ã„⁄ ‘œÂ «‰œ"
      TabPicture(0)   =   "Form16.frx":110B9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "«ﬁ”«ÿÌ òÂ ÅÌ‘ »Ì‰Ì „Ì ‘Êœ œ—Ì«›  ‘Êœ"
      TabPicture(1)   =   "Form16.frx":110D5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(7)=   "Label11"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "«ﬁ”«ÿ „—»Êÿ »Â «Ì‰ „«Â òÂ œ— «Ì‰ „«Â Ã„⁄ ‘œÂ «‰œ"
      TabPicture(2)   =   "Form16.frx":110F1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label27"
      Tab(2).Control(1)=   "Label26"
      Tab(2).Control(2)=   "Label25"
      Tab(2).Control(3)=   "Label24"
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(5)=   "Label22"
      Tab(2).Control(6)=   "Label21"
      Tab(2).Control(7)=   "Label20"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "«ﬁ”«ÿ „—»Êÿ »Â «Ì‰ „«Â òÂ œ— „«Â œÌê— Ã„⁄ ‘œÂ «‰œ"
      TabPicture(3)   =   "Form16.frx":1110D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label35"
      Tab(3).Control(1)=   "Label34"
      Tab(3).Control(2)=   "Label33"
      Tab(3).Control(3)=   "Label32"
      Tab(3).Control(4)=   "Label31"
      Tab(3).Control(5)=   "Label30"
      Tab(3).Control(6)=   "Label29"
      Tab(3).Control(7)=   "Label28"
      Tab(3).ControlCount=   8
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄  „«„Ì Ê«„ Â« :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ÊÌéÂ :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì «÷ÿ—«—Ì :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ⁄«œÌ :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ⁄«œÌ :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì «÷ÿ—«—Ì :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ÊÌéÂ :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄  „«„Ì Ê«„ Â« :"
         Height          =   495
         Left            =   -72240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1860
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄  „«„Ì Ê«„ Â« :"
         Height          =   495
         Left            =   -72120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ÊÌéÂ :"
         Height          =   495
         Left            =   -72120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -69360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1860
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì «÷ÿ—«—Ì :"
         Height          =   495
         Left            =   -66840
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   -69360
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ⁄«œÌ :"
         Height          =   495
         Left            =   -66840
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ⁄«œÌ :"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì «÷ÿ—«—Ì :"
         Height          =   495
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄ Ê«„ Â«Ì ÊÌéÂ :"
         Height          =   495
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "„Ã„Ê⁄  „«„Ì Ê«„ Â« :"
         Height          =   495
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2100
         Width           =   2415
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”«·"
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «ﬁ”«ÿ „«Â«‰Â"
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
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qw(12) As String

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

'„»«·€ Ã„⁄ ¬Ê—Ì ‘œÂ œ— «Ì‰ „«Â
MSChart1.Column = 1
For q = 1 To 12
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
    If Left(Form8.Adodc1.Recordset.Fields!Date, 4) = a Then
      If Right(Left(Form8.Adodc1.Recordset.Fields!Date, 7), 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc1.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
  
  Form8.Adodc2.Recordset.MoveFirst
  Do
    If Left(Form8.Adodc2.Recordset.Fields!Date, 4) = a Then
      If Right(Left(Form8.Adodc2.Recordset.Fields!Date, 7), 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc2.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc2.Recordset.MoveNext
  Loop Until Form8.Adodc2.Recordset.EOF = True
  
  Form8.Adodc3.Recordset.MoveFirst
  Do
    If Left(Form8.Adodc3.Recordset.Fields!Date, 4) = a Then
      If Right(Left(Form8.Adodc3.Recordset.Fields!Date, 7), 2) = b Then
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
For q = 1 To 12
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
    If (Left(Form8.Adodc1.Recordset.Fields!saragsat, 4) = a) And (Left(Form8.Adodc1.Recordset.Fields!Date, 4) = a) Then
      If Right(Left(Form8.Adodc1.Recordset.Fields!Date, 7), 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc1.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
  
  Form8.Adodc2.Recordset.MoveFirst
  Do
    If (Left(Form8.Adodc2.Recordset.Fields!saragsat, 4) = a) And (Left(Form8.Adodc2.Recordset.Fields!Date, 4) = a) Then
      If Right(Left(Form8.Adodc2.Recordset.Fields!Date, 7), 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form8.Adodc2.Recordset.Fields!Money)
      End If
    End If
    Form8.Adodc2.Recordset.MoveNext
  Loop Until Form8.Adodc2.Recordset.EOF = True

  Form8.Adodc3.Recordset.MoveFirst
  Do
    If (Left(Form8.Adodc3.Recordset.Fields!saragsat, 4) = a) And (Left(Form8.Adodc3.Recordset.Fields!Date, 4) = a) Then
      If Right(Left(Form8.Adodc3.Recordset.Fields!Date, 7), 2) = b Then
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

For q = 1 To 12
  MSChart1.Row = q
  MSChart1.Data = 0
  qw(q) = 0
  If Len(Trim(Str(q))) = 1 Then
    b = "0" + Trim(Str(q))
  Else
    b = Trim(Str(q))
  End If
  a = Text2.Text + "/" + b + "/01"
  Form7.Adodc1.Recordset.MoveFirst
  Do
    s = Amin.dateaminEktelafmoon(Form7.Adodc1.Recordset.Fields!Date, a)
    If (a >= Form7.Adodc1.Recordset.Fields!Date) And (Val(s) <= Val(Form7.Adodc1.Recordset.Fields!numberagsat)) Then
      d = Amin.dateaminEzafeMoon(Form7.Adodc1.Recordset.Fields!Date, Val(s))
      If Right(Left(d, 7), 2) = b Then
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
      If Right(Left(d, 7), 2) = b Then
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
      If Right(Left(d, 7), 2) = b Then
        qw(q) = Val(qw(q)) + Val(Form7.Adodc3.Recordset.Fields!moneyg2)
      End If
    End If
    Form7.Adodc3.Recordset.MoveNext
  Loop Until Form7.Adodc3.Recordset.EOF = True
  
  MSChart1.Row = q
  MSChart1.Data = qw(q)
Next q
End Sub

