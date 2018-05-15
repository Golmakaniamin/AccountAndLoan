VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form17"
   ScaleHeight     =   7455
   ScaleWidth      =   9120
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
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub Form_Load()
Dim q As String
Screen.MousePointer = vbHourglass

If Form8.Option1.Value = True Then q = "ÚÇÏí"
If Form8.Option2.Value = True Then q = "ÇÖØÑÇÑí"
If Form8.Option3.Value = True Then q = "æíŽå"
q = "   " + q
Report.Text9.SetText Form8.List6.List(Form8.List6.ListIndex) + q
Report.Text25.SetText Form8.List6.List(Form8.List6.ListIndex) + q

Report.Text24.SetText Form8.List2.List(Form8.List2.ListIndex)
Report.Text29.SetText Form8.List2.List(Form8.List2.ListIndex)

Report.Text21.SetText Form8.Label4(7).Caption
Report.Text26.SetText Form8.Label4(7).Caption

Report.Text22.SetText Form8.List1.List(Form8.List1.ListIndex)
Report.Text27.SetText Form8.List1.List(Form8.List1.ListIndex)

Report.Text23.SetText Form8.List3.List(Form8.List3.ListIndex)
Report.Text28.SetText Form8.List3.List(Form8.List3.ListIndex)

Report.Text12.SetText Form8.List9.List(Form8.List9.ListIndex)
Report.Text33.SetText Form8.List9.List(Form8.List9.ListIndex)

Report.Text8.SetText Form8.Label4(4).Caption
Report.Text31.SetText Form8.Label4(4).Caption

Report.Text13.SetText Form8.Label4(5).Caption
Report.Text35.SetText Form8.Label4(5).Caption

Report.Text10.SetText Form2.Label2.Caption
Report.Text30.SetText Form2.Label2.Caption

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
