VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form21"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      lastProp        =   500
      _cx             =   14896
      _cy             =   12356
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   0   'False
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport5

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Report.Text3.SetText Form10.List6.List(Form10.List6.ListIndex)
If Form10.Option1.Value = True Then Report.Text13.SetText "ÚÇÏí"
If Form10.Option2.Value = True Then Report.Text13.SetText "ÇÖØÑÇÑí"
If Form10.Option3.Value = True Then Report.Text13.SetText "æíŽå"
Report.Text15.SetText Form10.Label4(6).Caption
Report.Text14.SetText Form10.Label4(7).Caption

Report.Text11.SetText ""

Report.Text12.SetText Form10.Label4(2).Caption

Report.Text18.SetText Form10.Label4(0).Caption
Report.Text10.SetText Form10.Label4(1).Caption
Report.Text4.SetText Form10.Label4(11).Caption
Report.Text16.SetText Form10.Label4(10).Caption

CRViewer91.ReportSource = Report
CRViewer91.ViewReport
Screen.MousePointer = vbDefault
Report.PrintOutEx
End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub
