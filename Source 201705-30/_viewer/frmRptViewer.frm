VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRptViewer 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12840
   ClipControls    =   0   'False
   Icon            =   "frmRptViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer crViewer 
      Height          =   4245
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   4860
      lastProp        =   600
      _cx             =   8572
      _cy             =   7488
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4845
      Top             =   1515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptViewer.frx":08CA
            Key             =   "setup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptViewer.frx":225C
            Key             =   "print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptViewer.frx":3BEE
            Key             =   "export"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptViewer.frx":5580
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   1058
      ButtonWidth     =   2805
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Printer Setup"
            Object.ToolTipText     =   "Setup Printer Settings"
            ImageKey        =   "setup"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print Report"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Object.ToolTipText     =   "Export Report"
            ImageKey        =   "export"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close this Preview"
            ImageKey        =   "close"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Cost Accounting System
' module        :   frmRptViewer
' programmer    :   _-=[ srm ]=-_
' date          :   25 feb 2006

Option Explicit
Dim cRptName As String, cRptCaption As String, _
    objCrystal As New CRAXDRT.Application, _
    oReport As New CRAXDRT.Report

Sub SetFilter(ByVal cReportName As String, ByVal cReportTitle As String)
    cRptName = cReportName
    cRptCaption = cReportTitle
End Sub

Function IsFormula(oRpt As CRAXDRT.Report, cFormulaName As String) As Boolean
    Dim nCtr As Integer
    DoEvents
    For nCtr = 1 To oRpt.FormulaFields.Count
        If oRpt.FormulaFields(nCtr).FormulaFieldName = cFormulaName Then
            IsFormula = True
            Exit Function
        End If
    Next nCtr
End Function

Sub SetFormula(oRpt As CRAXDRT.Report, ByVal cFormulaName As String, ByVal cValue As String)
    If IsFormula(oRpt, cFormulaName) Then oRpt.FormulaFields.GetItemByName(cFormulaName).Text = cQuote & cValue & cQuote
End Sub

Private Sub Form_Load()
    DoEvents
    With CRViewer
        Set oReport = objCrystal.OpenReport(cReportPath & cRptName)
        
        oReport.DiscardSavedData
        
        SetFormula oReport, "cCompanyName", cCompany
        SetFormula oReport, "cReportTitle", cRptCaption
        SetFormula oReport, "cAddress", gAddress
        
        .ReportSource = oReport
        .ViewReport
        
        Do While .IsBusy
            DoEvents
        Loop
        
        .Zoom 94
    End With
End Sub

Private Sub Form_Resize()
    With CRViewer
        .top = Toolbar1.Height
        .left = 0
        If ScaleHeight - Toolbar1.Height >= 0 Then .Height = ScaleHeight - Toolbar1.Height
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Terminate()
    Set objCrystal = Nothing
    Set oReport = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    With oReport
        Select Case Button.Index
            Case 1      ' --> printer setup
                .PrinterSetup Me.hwnd
    
            Case 3      ' --> print
                .PrintOut True
                
            Case 5
                ' --> export
    '            oReport.ExportOptions.FormatType = crEFTExcel97
    '            oReport.ExportOptions.DiskFileName = "test.pdf"
    '            oReport.ExportOptions.DestinationType = crEDTDiskFile
    '            oReport.Export True
                .Export True
    
            Case 7      ' --> close
                Unload Me
        End Select
    End With
End Sub
