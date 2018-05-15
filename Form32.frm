VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form32 
   Caption         =   "Form32"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form32"
   ScaleHeight     =   5985
   ScaleWidth      =   9375
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
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim Report As New CrystalReport9
If Form10.Option1.Value = True Then q = "ÚÇÏí"
If Form10.Option2.Value = True Then q = "ÇÖØÑÇÑí"
If Form10.Option3.Value = True Then q = "æíŽå"
q = "   " + q
Report.Text17.SetText Form10.List6.List(Form10.List6.ListIndex) + q
Report.Text1.SetText Form10.Label4(6).Caption
Report.Text21.SetText Form10.Label4(1).Caption
Report.Text32.SetText Form10.Label4(2).Caption
Report.Text28.SetText Form10.Label4(9).Caption
Report.Text3.SetText Form10.Label4(7).Caption
Report.Text7.SetText Form10.Label4(12).Caption
Report.Text16.SetText Form10.Label4(13).Caption
Report.Text24.SetText Form10.Label4(8).Caption
Report.Text12.SetText Form10.Label4(3).Caption
Report.Text9.SetText Form10.Label4(0).Caption
Report.Text14.SetText Form10.Label4(11).Caption
Report.Text11.SetText Form10.Label4(10).Caption
Report.Text30.SetText Form10.Label4(4).Caption
Report.Text29.SetText Form10.Label4(5).Caption
Report.Text39.SetText Form2.Label2.Caption
Report.Text46.SetText Form2.Label5.Caption
Screen.MousePointer = vbHourglass
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
