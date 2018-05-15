VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form Form15 
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   Picture         =   "Form15.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List12 
      Height          =   1095
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C3CEC4&
      Caption         =   "⁄«œÌ Ê «÷ÿ—«—Ì"
      Height          =   2535
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   8295
      Begin VB.ListBox List10 
         Height          =   1785
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox List8 
         Height          =   1785
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List9 
         Height          =   1785
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox List7 
         Height          =   1785
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ Ê«„ :"
         Height          =   375
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄"
         Height          =   495
         Index           =   9
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ⁄÷ÊÌ  :"
         Height          =   375
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÊÌÂ"
         Height          =   495
         Index           =   8
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ Ê«„"
         Height          =   495
         Index           =   5
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Ê«„"
         Height          =   495
         Index           =   4
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò”—Ì :"
         Height          =   375
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊÃÊœÌ :"
         Height          =   375
         Index           =   0
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C3CEC4&
      Caption         =   "ÊÌéÂ"
      Height          =   4575
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5880
      Width           =   10935
      Begin VB.ListBox List14 
         Height          =   1095
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox List13 
         Height          =   1095
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox List11 
         Height          =   1095
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   1095
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox List6 
         Height          =   2130
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.ListBox List5 
         Height          =   2130
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.ListBox List4 
         Height          =   2130
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List3 
         Height          =   2130
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.ListBox List2 
         Height          =   2130
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»” «‰ò«—Ì »Â ’‰œÊﬁ :"
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ  „«„Ì Õ”«» Â« :"
         Height          =   375
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ Õ”«» Â«Ì œ«—«Ì Ê«„ :"
         Height          =   375
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ Õ”«» Â«Ì »œÊ‰ Ê«„ :"
         Height          =   375
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊÃÊœÌ ò· Õ”«» Â« :"
         Height          =   375
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊÃÊœÌ Õ”«» Â«Ì œ«—«Ì Ê«„ :"
         Height          =   375
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊÃÊœÌ Õ”«» Â«Ì »œÊ‰ Ê«„ :"
         Height          =   375
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄ „»«·€ ò· Ê«„ Â« :"
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄ „»«·€ Ê«„  ”ÊÌÂ ‘œÂ :"
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄ „»«·€ Ê«„ Ã«—Ì :"
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ò· Ê«„ Â« :"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ Ê«„  ”ÊÌÂ ‘œÂ :"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ Ê«„ Ã«—Ì :"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÊÌÂ"
         Height          =   495
         Index           =   2
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ Ê«„"
         Height          =   495
         Index           =   1
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Ê«„"
         Height          =   495
         Index           =   3
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊÃÊœÌ"
         Height          =   495
         Index           =   7
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â Õ”«»"
         Height          =   495
         Index           =   6
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C3CEC4&
      Caption         =   "„‘Œ’« "
      Height          =   2055
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   120
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ã” ÃÊ"
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
         MICON           =   "Form15.frx":10378
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
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "‰„«Ì‘"
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
         MICON           =   "Form15.frx":10394
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "òœ ⁄÷ÊÌ "
         Height          =   495
         Index           =   0
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      MICON           =   "Form15.frx":103B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons4 
      Height          =   495
      Left            =   120
      TabIndex        =   65
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "ÅÌ«„ò Â«Ì «—”«·Ì"
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
      MICON           =   "Form15.frx":103CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons5 
      Height          =   495
      Left            =   120
      TabIndex        =   67
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "÷«„‰ Â«"
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
      MICON           =   "Form15.frx":103E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      Caption         =   "Label35"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ :"
      Height          =   495
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ê÷⁄Ì  Õ”«»"
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
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Text1(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If
End Sub

Private Sub KewlButtons1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub KewlButtons2_Click()
Form14.Show
End Sub

Private Sub KewlButtons3_Click()
Dim p As Boolean
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
List11.Clear
List12.Clear
List14.Clear
p = False
Form3.Adodc1.Recordset.MoveFirst
Do
  If Form3.Adodc1.Recordset.Fields!id = Trim(Text1(0).Text) Then
    Label12.Caption = Form3.Adodc1.Recordset.Fields!Name + " " + Form3.Adodc1.Recordset.Fields!family
    Label4.Caption = Form3.Adodc1.Recordset.Fields!Money
    Label7.Caption = Form3.Adodc1.Recordset.Fields!edate
    Label35.Caption = Form3.Adodc1.Recordset.Fields!mobile
    q = Val(Amin.dateaminEktelafmoon(Label7.Caption, Form2.Label5.Caption))
    Label6.Caption = 0
    Label6.Caption = Val(Label4.Caption) - (q * 20000)
    If Label6.Caption >= 0 Then
      Label5.Caption = "„«“«œ :"
    Else
      Label6.Caption = -1 * Val(Label6.Caption)
      Label5.Caption = "ò”—Ì :"
    End If
    '
    Form7.Adodc1.Recordset.MoveFirst
    Do
      If Form7.Adodc1.Recordset.Fields!id1 = Text1(0).Text Then
        List7.AddItem Form7.Adodc1.Recordset.Fields!id
        List8.AddItem Form7.Adodc1.Recordset.Fields!moneyvam
        List9.AddItem Form7.Adodc1.Recordset.Fields!tasvie
        List10.AddItem "⁄«œÌ"
      End If
      Form7.Adodc1.Recordset.MoveNext
    Loop Until Form7.Adodc1.Recordset.EOF = True
    
    Form7.Adodc2.Recordset.MoveFirst
    Do
      If Form7.Adodc2.Recordset.Fields!id1 = Text1(0).Text Then
        List7.AddItem Form7.Adodc2.Recordset.Fields!id
        List8.AddItem Form7.Adodc2.Recordset.Fields!moneyvam
        List9.AddItem Form7.Adodc2.Recordset.Fields!tasvie
        List10.AddItem "«÷ÿ—«—Ì"
      End If
      Form7.Adodc2.Recordset.MoveNext
    Loop Until Form7.Adodc2.Recordset.EOF = True
    Label10.Caption = List7.ListCount
    
    Form4.Adodc1.Recordset.MoveFirst
    Do
      If Form4.Adodc1.Recordset.Fields!idadi = Text1(0).Text Then
        List2.AddItem Form4.Adodc1.Recordset.Fields!id
        List3.AddItem Form4.Adodc1.Recordset.Fields!Money
      End If
      Form4.Adodc1.Recordset.MoveNext
    Loop Until Form4.Adodc1.Recordset.EOF = True
    
    For q = 0 To List2.ListCount - 1
      Form7.Adodc3.Recordset.Find "id1='" & List2.List(q) & "'", , adSearchForward, 1
      If Form7.Adodc3.Recordset.EOF = False Then
        List1.AddItem "+"
        Form7.Adodc3.Recordset.MoveFirst
        Do
          If Form7.Adodc3.Recordset.Fields!id1 = List2.List(q) Then
            List14.AddItem Form7.Adodc3.Recordset.Fields!id
            List11.AddItem Form7.Adodc3.Recordset.Fields!moneyvam
            List12.AddItem Form7.Adodc3.Recordset.Fields!tasvie
          End If
          Form7.Adodc3.Recordset.MoveNext
        Loop Until Form7.Adodc3.Recordset.EOF = True
      Else
        List1.AddItem "-"
      End If
    Next q
    
    Label26.Caption = 0
    Label28.Caption = 0
    Label30.Caption = 0
    
    Label20.Caption = 0
    Label22.Caption = 0
    Label24.Caption = 0

    Label16.Caption = List11.ListCount
    Label11.Caption = List11.ListCount
    Label18.Caption = List11.ListCount
    
    Label33.Caption = 0
    Label32.Caption = 0
    Label37.Caption = List2.ListCount
    
    Label36.Caption = 0

    For q = 0 To List1.ListCount - 1
      If List1.List(q) = "-" Then
        Label32.Caption = Val(Label32.Caption) + 1
        Label26.Caption = Val(Label26.Caption) + Val(List3.List(q))
      Else
        Label33.Caption = Val(Label33.Caption) + 1
        Label28.Caption = Val(Label28.Caption) + Val(List3.List(q))
      End If
    Next q
    Label30.Caption = Val(Label28.Caption) + Val(Label26.Caption)
    
    For q = 0 To List11.ListCount - 1
      If List12.List(q) = "‘œÂ" Then
        Label22.Caption = Val(Label22.Caption) + Val(List11.List(q))
        Label16.Caption = Val(Label16.Caption) - 1
      Else
        Label20.Caption = Val(Label20.Caption) + Val(List11.List(q))
        Label11.Caption = Val(Label11.Caption) - 1
        Form8.Adodc3.CommandType = adCmdUnknown
        Form8.Adodc3.RecordSource = "SELECT * FROM GvamVig WHERE id='" + List14.List(q) + "'"
        Form8.Adodc3.Refresh
        tmpas = Val(List11.List(q))
        If Form8.Adodc3.Recordset.RecordCount > 0 Then
          Form8.Adodc3.Recordset.MoveFirst
          Do
            tmpas = Val(tmpas) - Val(Form8.Adodc3.Recordset.Fields!Money)
            Form8.Adodc3.Recordset.MoveNext
          Loop Until Form8.Adodc3.Recordset.EOF = True
        End If
        Label36.Caption = Val(Label36.Caption) + Val(tmpas)
      End If
      Label24.Caption = Val(Label20.Caption) + Val(Label22.Caption)
    Next q
    p = True
    Exit Do
  End If
  Form3.Adodc1.Recordset.MoveNext
Loop Until Form3.Adodc1.Recordset.EOF = True

Form8.Adodc3.CommandType = adCmdUnknown
Form8.Adodc3.RecordSource = "SELECT * FROM GvamVig"
Form8.Adodc3.Refresh

Label4.Caption = Amin.moneyaminjoda(Label4.Caption)
Label6.Caption = Amin.moneyaminjoda(Label6.Caption)

Label26.Caption = Amin.moneyaminjoda(Label26.Caption)
Label28.Caption = Amin.moneyaminjoda(Label28.Caption)
Label30.Caption = Amin.moneyaminjoda(Label30.Caption)

Label20.Caption = Amin.moneyaminjoda(Label20.Caption)
Label22.Caption = Amin.moneyaminjoda(Label22.Caption)
Label24.Caption = Amin.moneyaminjoda(Label24.Caption)

Label36.Caption = Amin.moneyaminjoda(Label36.Caption)
List13.Clear
For q = 0 To List3.ListCount - 1
  List13.AddItem Amin.moneyaminjoda(List3.List(q))
Next q
List3.Clear
For q = 0 To List13.ListCount - 1
  List3.AddItem List13.List(q)
Next q

List13.Clear
For q = 0 To List8.ListCount - 1
  List13.AddItem Amin.moneyaminjoda(List8.List(q))
Next q
List8.Clear
For q = 0 To List13.ListCount - 1
  List8.AddItem List13.List(q)
Next q

If p = False Then
  z = MsgBox("òœ Õ”«» ⁄«œÌ Ê«—œ ‘œÂ œ— ”Ì” „ ÊÃÊœ ‰œ«—œ", vbMsgBoxRight + vbCritical, "")
  Label12.Caption = "-"
End If
End Sub

Private Sub KewlButtons4_Click()
If Label12.Caption <> "-" Then
  Form30.Adodc1.CommandType = 8
  Form30.Adodc1.RecordSource = "select * from sendsms"
  Form30.Adodc1.Refresh
  Form30.Adodc1.RecordSource = "select * from sendsms where number='" + Label35.Caption + "'"
  Form30.Adodc1.Refresh
  Form30.DataGrid1.Caption = "·Ì”  ÅÌ«„ò Â«Ì «—”«· ‘œÂ »Â ¬ﬁ«Ì " + Label12.Caption
  Form30.DataGrid1.Refresh
  Form30.Show
End If
End Sub

Private Sub List2_Click()
List3.ListIndex = List2.ListIndex

List4.Clear
List5.Clear
List6.Clear
Form7.Adodc3.Recordset.MoveFirst
Do
  If Form7.Adodc3.Recordset.Fields!id1 = List2.List(List2.ListIndex) Then
    List4.AddItem Form7.Adodc3.Recordset.Fields!id
    List5.AddItem Form7.Adodc3.Recordset.Fields!moneyvam
    List6.AddItem Form7.Adodc3.Recordset.Fields!tasvie
  End If
  Form7.Adodc3.Recordset.MoveNext
Loop Until Form7.Adodc3.Recordset.EOF = True

If List5.ListCount <> 0 Then
  List13.Clear
  For q = 0 To List5.ListCount - 1
    List13.AddItem Amin.moneyaminjoda(List5.List(q))
  Next q
  List5.Clear
  For q = 0 To List13.ListCount - 1
    List5.AddItem List13.List(q)
  Next q
End If
End Sub


Private Sub List3_Click()
List2.ListIndex = List3.ListIndex
End Sub


Private Sub List4_Click()
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List4.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
End Sub

Private Sub List6_Click()
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
End Sub

Private Sub List7_Click()
List8.ListIndex = List7.ListIndex
List9.ListIndex = List7.ListIndex
List10.ListIndex = List7.ListIndex
End Sub

Private Sub List8_Click()
List7.ListIndex = List8.ListIndex
List9.ListIndex = List8.ListIndex
List10.ListIndex = List8ListIndex
End Sub

Private Sub List9_Click()
List7.ListIndex = List9.ListIndex
List8.ListIndex = List9.ListIndex
List10.ListIndex = List9.ListIndex
End Sub

Private Sub List10_Click()
List7.ListIndex = List10.ListIndex
List9.ListIndex = List10.ListIndex
List8.ListIndex = List10.ListIndex
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  KewlButtons3.SetFocus
End If
End Sub
