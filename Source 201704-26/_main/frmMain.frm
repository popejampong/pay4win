VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F1C2FD4A-613D-432B-A4E4-0076F9414952}#1.1#0"; "ciaXPSideBarMenu.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll & Time Management System for Dong-in Entech"
   ClientHeight    =   5400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2790
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XPSideBarMenu.XPSideMenu XPSideMenu1 
      Align           =   3  'Align Left
      Height          =   5025
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8864
      LicValid        =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   60
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
            Picture         =   "frmMain.frx":5762
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CFF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5025
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "3/13/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "2:47 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      HelpContextID   =   1000
      Begin VB.Menu mnuFM 
         Caption         =   "File Maintenance"
         HelpContextID   =   1100
         Begin VB.Menu mnuAdmin 
            Caption         =   "Administrator Entry"
            HelpContextID   =   1101
         End
         Begin VB.Menu fm1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmployee 
            Caption         =   "Employee Entry"
            HelpContextID   =   1103
         End
         Begin VB.Menu mnuBlocklist 
            Caption         =   "Blocklist Entry"
            HelpContextID   =   1104
         End
         Begin VB.Menu mnuLevel 
            Caption         =   "Level Entry"
            HelpContextID   =   1109
         End
         Begin VB.Menu mnuDept 
            Caption         =   "Department Entry"
            HelpContextID   =   1102
         End
         Begin VB.Menu mnuPosition 
            Caption         =   "Position Entry"
            HelpContextID   =   1104
         End
         Begin VB.Menu mnuShift 
            Caption         =   "Shift Entry"
            HelpContextID   =   1107
         End
         Begin VB.Menu fm2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHoliday 
            Caption         =   "Holiday Entry"
            HelpContextID   =   1105
         End
         Begin VB.Menu mnuPayPeriod 
            Caption         =   "Payroll Period Entry"
            HelpContextID   =   1106
         End
         Begin VB.Menu mnuLeave 
            Caption         =   "Incentive Leave"
            HelpContextID   =   1108
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "Swap Date Entry"
            HelpContextID   =   1110
         End
      End
      Begin VB.Menu mnuFMDed 
         Caption         =   "Deduction"
         HelpContextID   =   1200
         Begin VB.Menu mnuDeduction 
            Caption         =   "Deduction Entry"
            HelpContextID   =   1201
         End
         Begin VB.Menu mnuEmpLoan 
            Caption         =   "Employee Loan Entry"
            HelpContextID   =   1205
         End
         Begin VB.Menu d1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSSS 
            Caption         =   "SSS Table Entry"
            HelpContextID   =   1202
         End
         Begin VB.Menu mnuPhilHealth 
            Caption         =   "Philhealth Table Entry"
            HelpContextID   =   1203
         End
         Begin VB.Menu d2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTax 
            Caption         =   "Withholding Tax Entry"
            HelpContextID   =   1204
         End
         Begin VB.Menu mnuTax2 
            Caption         =   "Annual Tax Entry"
            HelpContextID   =   1206
         End
         Begin VB.Menu mnuTax2Old 
            Caption         =   "Annual Tax Entry OLD"
            HelpContextID   =   1207
         End
      End
      Begin VB.Menu mnuFMRptLst 
         Caption         =   "Report Listing"
         HelpContextID   =   1400
         Begin VB.Menu mnuRptEmp 
            Caption         =   "Employee Listing"
            HelpContextID   =   1401
         End
         Begin VB.Menu r1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptDep 
            Caption         =   "Department Listing"
            HelpContextID   =   1403
         End
         Begin VB.Menu mnuRptPos 
            Caption         =   "Position Listing"
            HelpContextID   =   1404
         End
         Begin VB.Menu r2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptHoliday 
            Caption         =   "Holiday Listing"
            HelpContextID   =   1406
         End
         Begin VB.Menu mnuRptPeriod 
            Caption         =   "Period Listing"
            HelpContextID   =   1407
         End
         Begin VB.Menu mnuRptShift 
            Caption         =   "Shift Listing"
            HelpContextID   =   1405
         End
         Begin VB.Menu r3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptDed 
            Caption         =   "Deduction Listing"
            HelpContextID   =   1408
         End
         Begin VB.Menu mnuRptWTax 
            Caption         =   "Withholding Tax Report"
            HelpContextID   =   1409
         End
      End
      Begin VB.Menu f1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "Custom"
         HelpContextID   =   1600
         Begin VB.Menu mnuEmpSSRep 
            Caption         =   "Employment Report (SS R-1A)"
            HelpContextID   =   1601
         End
         Begin VB.Menu mnuEmpPHRep 
            Caption         =   "PHILHEALTH ER2"
            HelpContextID   =   1602
         End
         Begin VB.Menu mnuEmpOt 
            Caption         =   "Employee OT listing"
            HelpContextID   =   1604
         End
         Begin VB.Menu mnuEmpLoanRpt 
            Caption         =   "Employee Loan Report"
            HelpContextID   =   1603
         End
         Begin VB.Menu mnuWLCostRpt 
            Caption         =   "Actual Weekly Labor Cost"
            HelpContextID   =   1605
         End
         Begin VB.Menu c1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDLCostRpt 
            Caption         =   "Daily Labor Cost"
            HelpContextID   =   1606
         End
         Begin VB.Menu mnuMasterData 
            Caption         =   "Employee Master Data Generation"
            HelpContextID   =   1607
         End
         Begin VB.Menu c2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSal_Increase 
            Caption         =   "Salary Increase Update"
            HelpContextID   =   1608
         End
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log-Off"
         HelpContextID   =   1500
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   1300
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTime 
      Caption         =   "Time Management"
      HelpContextID   =   2000
      Begin VB.Menu mnuShiftSched 
         Caption         =   "Shift Schedule"
         HelpContextID   =   2100
         Begin VB.Menu mnuDepShiftSched 
            Caption         =   "Shifting Schedule (by Department)"
            HelpContextID   =   2101
         End
         Begin VB.Menu mnuEmpShiftSched 
            Caption         =   "Shifting Schedule"
            HelpContextID   =   2102
         End
      End
      Begin VB.Menu mnuEmpLeave 
         Caption         =   "Leave Entry"
         HelpContextID   =   2200
      End
      Begin VB.Menu mnuEmpInc 
         Caption         =   "Incentive"
         HelpContextID   =   2800
      End
      Begin VB.Menu mnuSSCheck 
         Caption         =   "Shift Schedule Checking"
         HelpContextID   =   2900
      End
      Begin VB.Menu t1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFinger 
         Caption         =   "Fingerprint Registration"
         HelpContextID   =   2300
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEmpLvl 
         Caption         =   "Employee Level Entry"
         HelpContextID   =   2400
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "Upload Data"
         HelpContextID   =   2500
         Shortcut        =   {F3}
      End
      Begin VB.Menu t2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMS 
         Caption         =   "Time Management"
         HelpContextID   =   2600
      End
      Begin VB.Menu t3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepTMS 
         Caption         =   "Report Listing"
         HelpContextID   =   2700
         Begin VB.Menu mnuTMSDailyRep 
            Caption         =   "Daily Attendance Report"
            HelpContextID   =   2701
         End
         Begin VB.Menu mnuTMSLA 
            Caption         =   "Late/Absent Report"
            HelpContextID   =   2702
         End
         Begin VB.Menu mnuWklyCon 
            Caption         =   "Weekly Consumption Report"
            HelpContextID   =   2709
         End
         Begin VB.Menu t4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMonthlyFC 
            Caption         =   "Finish Contract Report (Monthly)"
            HelpContextID   =   2703
         End
         Begin VB.Menu mnuPeriodFC 
            Caption         =   "Finish Contract Report (by Period)"
            HelpContextID   =   2704
         End
         Begin VB.Menu t5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptLeave 
            Caption         =   "Leave Report"
            HelpContextID   =   2705
         End
         Begin VB.Menu t6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRptDTCons 
            Caption         =   "Daily Time Consumption Report"
            HelpContextID   =   2706
         End
         Begin VB.Menu mnuRptDTRRep 
            Caption         =   "Daily Time Attendance Report"
            HelpContextID   =   2707
         End
         Begin VB.Menu mnuRptActMan 
            Caption         =   "Daily Actual Manpower Report"
            HelpContextID   =   2709
         End
         Begin VB.Menu mnuRptDTRPAna 
            Caption         =   "Payroll Analysis"
            HelpContextID   =   2708
         End
      End
      Begin VB.Menu mnuTMSList 
         Caption         =   "TMS Report Listing"
         HelpContextID   =   2800
         Begin VB.Menu mnuTMSRpt 
            Caption         =   "TMS Report"
            HelpContextID   =   2801
         End
         Begin VB.Menu mnuDTRSum 
            Caption         =   "DTR (Summary)"
            HelpContextID   =   2802
         End
         Begin VB.Menu mnnuDTRDet 
            Caption         =   "DTR (Detail)"
            HelpContextID   =   2803
         End
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "Payroll"
      HelpContextID   =   4000
      Begin VB.Menu mnuTransaction 
         Caption         =   "Process Transaction"
         HelpContextID   =   4100
      End
      Begin VB.Menu p1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayrollPrint 
         Caption         =   "Print Utility"
         HelpContextID   =   4200
         Begin VB.Menu mnuPaySlip 
            Caption         =   "Pay Slip"
            HelpContextID   =   4201
         End
         Begin VB.Menu mnuPayrollSheet 
            Caption         =   "Payroll Sheet"
            HelpContextID   =   4202
         End
         Begin VB.Menu mnuSalDiv 
            Caption         =   "Salary Division"
            HelpContextID   =   4206
         End
         Begin VB.Menu mnuAcknowledgeSheet 
            Caption         =   "Acknowledgement Sheet"
            HelpContextID   =   4203
         End
         Begin VB.Menu mnuDenomination 
            Caption         =   "Denomination Report"
            HelpContextID   =   4204
         End
         Begin VB.Menu mnuDen_res 
            Caption         =   "Denomination for Emergency & Resign Report"
            HelpContextID   =   4207
         End
      End
      Begin VB.Menu mnuPPRESFC 
         Caption         =   "Print Utility (Resigned\FC)"
         HelpContextID   =   4800
         Begin VB.Menu mnuPPRESFCPaySlip 
            Caption         =   "Pay Slip"
            HelpContextID   =   4801
         End
         Begin VB.Menu mnuPPRESFCPayShit 
            Caption         =   "Payroll Sheet"
            HelpContextID   =   4802
         End
         Begin VB.Menu mnuPPRESFCSalDiv 
            Caption         =   "Salary Division"
            HelpContextID   =   4803
         End
         Begin VB.Menu mnuPPRESFCAckShit 
            Caption         =   "Acknowledgement Sheet"
            HelpContextID   =   4804
         End
      End
      Begin VB.Menu mnuRemit 
         Caption         =   "Remittances"
         HelpContextID   =   4300
         Begin VB.Menu mnuRemitSSS 
            Caption         =   "SSS"
            HelpContextID   =   4301
         End
         Begin VB.Menu mnuSSSR3 
            Caption         =   "SSS R3"
            HelpContextID   =   4305
         End
         Begin VB.Menu mnuRemitMedicare 
            Caption         =   "Medicare"
            HelpContextID   =   4302
         End
         Begin VB.Menu mnuRemitPagIbig 
            Caption         =   "Pag-Ibig"
            HelpContextID   =   4303
         End
         Begin VB.Menu mnuRemitTax 
            Caption         =   "Withholding Tax"
            HelpContextID   =   4304
         End
      End
      Begin VB.Menu mnuPEZA 
         Caption         =   "PEZA Report"
         HelpContextID   =   4500
      End
      Begin VB.Menu p2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Payroll"
         HelpContextID   =   4600
      End
      Begin VB.Menu mnuClosePeriod 
         Caption         =   "Close Period"
         HelpContextID   =   4400
      End
      Begin VB.Menu mnu13mo 
         Caption         =   "13 Month Pay"
         HelpContextID   =   4700
         Begin VB.Menu mnu13moGen 
            Caption         =   "Process 13th Month Pay"
            HelpContextID   =   4701
         End
         Begin VB.Menu p3 
            Caption         =   "-"
         End
         Begin VB.Menu mnu13moslip 
            Caption         =   "Payslip"
            HelpContextID   =   4702
         End
         Begin VB.Menu mnu13mosheet 
            Caption         =   "Payroll Sheet"
            HelpContextID   =   4703
         End
         Begin VB.Menu mnu13moAcknowledgement 
            Caption         =   "Acknowledgment Sheet"
            HelpContextID   =   4705
         End
         Begin VB.Menu mnu13moDenom 
            Caption         =   "Denomination Report"
            HelpContextID   =   4706
         End
         Begin VB.Menu p4 
            Caption         =   "-"
         End
         Begin VB.Menu mnu13mobackup 
            Caption         =   "Backup 13th Month"
            HelpContextID   =   4704
         End
         Begin VB.Menu p7 
            Caption         =   "-"
         End
         Begin VB.Menu mnu13moslip_SLVL 
            Caption         =   "SLVL Payslip"
            HelpContextID   =   4707
         End
         Begin VB.Menu mnu13mosheet_SLVL 
            Caption         =   "SLVL Payroll Sheet"
            HelpContextID   =   4708
         End
         Begin VB.Menu mnu13moAck_SLVL 
            Caption         =   "SLVL Acknowledgment Sheet"
            HelpContextID   =   4709
         End
         Begin VB.Menu mnu13moDenom_SLVL 
            Caption         =   "SLVL Denomination Report"
            HelpContextID   =   4710
         End
      End
      Begin VB.Menu mnuPayInc 
         Caption         =   "Incentive"
         HelpContextID   =   4900
         Begin VB.Menu mnuPayIncSlip 
            Caption         =   "Pay Slip"
            HelpContextID   =   4901
         End
         Begin VB.Menu mnuPayrollIncSheet 
            Caption         =   "Payroll Sheet"
            HelpContextID   =   4902
         End
         Begin VB.Menu mnuPayIncSht2 
            Caption         =   "Payroll Sheet (Resigned/FC)"
            HelpContextID   =   4903
         End
         Begin VB.Menu mnuAcknowledgeSheetInc 
            Caption         =   "Acknowledgement Sheet"
            HelpContextID   =   4904
         End
      End
   End
   Begin VB.Menu mnuERP 
      Caption         =   "ERP"
      HelpContextID   =   7000
      Begin VB.Menu mnuERP_FMain 
         Caption         =   "File Maintenance"
         HelpContextID   =   7100
         Begin VB.Menu mnuERP_Cost_Center 
            Caption         =   "Cost Center Entry"
            HelpContextID   =   7101
         End
         Begin VB.Menu mnuERP_Work_Center 
            Caption         =   "Work Center Entry"
            HelpContextID   =   7102
         End
         Begin VB.Menu mnuERP_COMP 
            Caption         =   "Company"
            HelpContextID   =   7103
         End
      End
      Begin VB.Menu mnuERP_RPTG 
         Caption         =   "Report Generation"
         HelpContextID   =   7200
         Begin VB.Menu mnuERP_SALG 
            Caption         =   "Finance report"
            HelpContextID   =   7201
         End
         Begin VB.Menu mnuERP_EMPL 
            Caption         =   "Employee Report"
            HelpContextID   =   7203
         End
         Begin VB.Menu mnuERP_TMS 
            Caption         =   "TMS Report"
            HelpContextID   =   7204
         End
      End
   End
   Begin VB.Menu mnuAlphaList 
      Caption         =   "Alphalist"
      HelpContextID   =   6000
      Begin VB.Menu mnuAlpha 
         Caption         =   "Alphalist"
         HelpContextID   =   6100
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpAlpha 
         Caption         =   "Upload Utility"
         HelpContextID   =   6500
      End
   End
   Begin VB.Menu mnuRCBC 
      Caption         =   "RCBC"
      HelpContextID   =   5000
      Begin VB.Menu mnuRCBCRPTTrans 
         Caption         =   "Transmittal Sheet"
         HelpContextID   =   5201
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utility"
      HelpContextID   =   3000
      Begin VB.Menu mnuSerCon 
         Caption         =   "Server Config"
         HelpContextID   =   3800
      End
      Begin VB.Menu mnuGenBCID 
         Caption         =   "Bio-Clock ID Generator"
         HelpContextID   =   3500
      End
      Begin VB.Menu mnuXPSideMenu 
         Caption         =   "Display Side Menu"
         HelpContextID   =   3400
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuUtilMenu 
         Caption         =   "Menu"
         HelpContextID   =   3100
         Begin VB.Menu mnuSysMenu 
            Caption         =   "System Menu"
            HelpContextID   =   3101
         End
         Begin VB.Menu mnuAccRight 
            Caption         =   "Access Right"
            HelpContextID   =   3102
         End
         Begin VB.Menu mnuResetMnu 
            Caption         =   "Reset Menu"
            HelpContextID   =   3103
         End
      End
      Begin VB.Menu u1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRebuild 
         Caption         =   "Rebuild"
         HelpContextID   =   3300
         Begin VB.Menu mnuRebuildEmp 
            Caption         =   "Employee Entry"
            HelpContextID   =   3302
         End
         Begin VB.Menu mnuRebuildDed 
            Caption         =   "Deduction Entry"
            HelpContextID   =   3301
         End
         Begin VB.Menu mnuRebuildYTD 
            Caption         =   "YTD Info"
            HelpContextID   =   3303
         End
         Begin VB.Menu mnuWAP 
            Caption         =   "WAP Sequence Number"
            HelpContextID   =   3304
         End
      End
      Begin VB.Menu u2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAudit 
         Caption         =   "Audit Trail"
         HelpContextID   =   3200
      End
      Begin VB.Menu mnuUploadOld 
         Caption         =   "Upload Old Payroll Data"
         HelpContextID   =   3600
      End
      Begin VB.Menu mnuDTROld 
         Caption         =   "Old DTR Data"
         HelpContextID   =   3700
         Begin VB.Menu mnuDTROldRep 
            Caption         =   "Old DTR Report"
            HelpContextID   =   3701
         End
         Begin VB.Menu mnuDTROldBackup 
            Caption         =   "Backup"
            HelpContextID   =   3702
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                  $`````$
'                $( o  o )$
'    >------oOO--(_)--OOo------------------------------------------------------------------------------<
'    "Intelligent people can be bored. They just know a lot. But Smart people
'    are never bored, because they're always looking for something to engage
'    their minds."
'    >------oooo(O) (0)oooo----------------------------------------------------------------------------<

' project name  :   Dong-in Payroll & Time Management System
' module        :   frmMain
' description   :   Main Form
' programmer    :   _-=[ srm ]=-_
' date started  :   7 oct 2005

Option Explicit
    Dim oTempADO As New ADODB.Recordset
    Dim lSuccess As Boolean, _
        nTimeCnt As Long
        
Sub CheckFin_Active()
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String
    
    cSqlStmt = "select empid, datereg, date_hire, date_fin ,date_res," & _
               " depid,posid, rate_amt, active, emp_stat, cmpid " & _
               " from di3670 where (active=0)" & _
               " and (emp_stat<>2)" & _
               " and (date_fin<=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & ")"
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0, , 1
        
        While Not oRecordSet.EOF
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Changing status of Emp Id#" & oRecordSet("empid")
            
            OpenQueryDNS "update di3670 set active = 2 where empid =" & cQuote & oRecordSet("empid") & cQuote, objdbRs, True
            
            cSqlStmt = " INSERT INTO PA3674(empid, time_history, date_history, datereg, date_hire, date_fin ,date_res," & _
                       " depid,posid, rate_amt, active, emp_stat, cmpid)VALUES(" & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & Format(Time, "hh:ss:mm") & cQuote & "," & _
                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("datereg"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_res"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & oRecordSet("depid") & cQuote & "," & _
                       cQuote & oRecordSet("posid") & cQuote & "," & _
                       oRecordSet("rate_amt") & "," & _
                       2 & "," & _
                       oRecordSet("emp_stat") & "," & _
                       cQuote & oRecordSet("cmpid") & cQuote & ")"
            
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        ShowProgress 4
    End If
End Sub
        
' --> report listing
Sub CreateTemp(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTableName As String, _
        cParam As String
    
    Select Case nMode
        Case 1      ' --> Holiday
            cTableName = "tmpHoliday"
            cParam = " [HOLIDAYID] char(3),        [DATE_HOL] date, " & _
                     " [DESCRIPTION] char(100),    [FIX_DAY] integer, "
                     
        Case 2      ' --> Position
            cTableName = "tmppos"
            cParam = " [POSID] char(3),        [POSNAME] char(100), " & _
                     " [STAFF] integer,        [ALLOWANCE] double, "
                     
        Case 3      ' --> Shift
            cTableName = "tmpShift"
            cParam = " [SHIFTID] char(5),      [DESCRIPTION] char(100), " & _
                     " [TIME1] char(10),       [TIME2] char(10), " & _
                     " [REMARK] char(50),      [NDIFF] integer, " & _
                     " [ALLOWANCE] double,     [REG_HR] double, " & _
                     " [BTIME] double,         [DEFAULT] integer, "
                     
        Case 4      ' --> Withholding Tax
            cTableName = "tmpWTax"
            cParam = " [TAXID] char(3),       [TAXCODE] char(5), " & _
               " [TAXNAME] char(100),   " & _
               " [DEDPCT0] double,      [DEDPCT1] double, " & _
               " [DEDPCT2] double,      [DEDPCT3] double, " & _
               " [DEDPCT4] double,      [DEDPCT5] double, " & _
               " [DEDPCT6] double," & _
               " [DEDAMT_0] double,     [DEDAMT_1] double, " & _
               " [DEDAMT_2] double,     [DEDAMT_3] double, " & _
               " [DEDAMT_4] double,     [DEDAMT_5] double, " & _
               " [DEDAMT_6] double, " & _
               " [DEDAMT1_0] double,     [DEDAMT1_1] double, " & _
               " [DEDAMT1_2] double,     [DEDAMT1_3] double, " & _
               " [DEDAMT1_4] double,     [DEDAMT1_5] double, " & _
               " [DEDAMT1_6] double, "

        Case 5      ' --> Deduction
            cTableName = "tmpDed"
            cParam = " [DEDID] char(3),         [DEDNAME] char(50)," & _
                     " [DEF_AMT] double,        [CUT_OFF_AMT] double," & _
                     " [FIX_DED] integer,       [PERIOD1] integer," & _
                     " [PERIOD2] integer,       [AUTO_DED] integer,"
    
        Case 6      ' --> Department
            cTableName = "tmpDep"
            cParam = " [LineID] char(3),        [LineName] char(100), " & _
                     " [production] integer,"
                     
        Case 7      ' --> Period
            cTableName = "tmpPPeriod"
            cParam = " [PERIODID] char(5),     [DATE_START] date, " & _
                     " [DATE_END] date,        [DURATION] char(50), " & _
                     " [PCLOSE] integer, " & _
                     " [PCLOSENAME] char(100), [DATE_CLOSE] date, " & _
                     " [WORKINDAYS] integer,   [HOLIDAYS] integer, " & _
                     " [STATUS] integer,"
    
    End Select
    
    If Trim(cParam) <> "" Then
        cSqlStmt = "CREATE TABLE " & cTableName & "(" & cParam & " [CMPName] char(50))"
        oTempConn.Execute cSqlStmt
        While oTempConn.State = adStateExecuting
            DoEvents
        Wend
    End If

ErrCreate:
    ' in case table is already existing, let's clear it...
    If Trim(cTableName) <> "" Then QueryTemp "DELETE FROM " & cTableName, objdbRs, True
End Sub
        
Sub showLogin()
    Dim nCtr As Integer, _
        cSqlStmt As String, _
        oADORS As New ADODB.Recordset
    
    lSuccess = False
    gUserID = ""
    gUserName = ""
    gUserLevel = 0
    gUserPW = ""
    
    gWSID = PadStr(right(Winsock1.LocalIP, Len(Winsock1.LocalIP) - InStrRev(Winsock1.LocalIP, ".")), "0", 3)
    
    For nCtr = 1 To 3
    
        lSuperUser = False
        
        frmLogin.Show 1
        If ModalResult = mrOk Then
        
            If lSuperUser Then
                gUserLevel = 1
                lSuccess = True
                Exit For
            Else
                OpenQueryDNS "SELECT * FROM PA2360 WHERE USERID=" & cQuote & gUserID & cQuote & " AND AES_DECRYPT(PASSWORD,UCASE(USERID))=" & cQuote & EncodeStr(gUserPW) & cQuote, oADORS, False
'                Script2File "SELECT * FROM PA2360 WHERE USERID=" & cQuote & gUserID & cQuote & " AND AES_DECRYPT(PASSWORD,UCASE(USERID))=" & cQuote & EncodeStr(gUserPW) & cQuote
                
                If Not oADORS.EOF Then
'-------------->Renz logout no more active
                    If gUserID = "renz" Then

                        gUserLevel = oADORS("USERLEVEL")
                        gUserName = oADORS("USERID") & " - " & oADORS("LASTNAME") & ", " & oADORS("FIRSTNAME")
                        gUserGroup = oADORS("groupid")
                        OpenQueryDNS "UPDATE PA2360 SET STATUS=0, WSID=" & cQuote & gWSID & cQuote & ", PA2360.TIME=" & cQuote & Format(Now, "hh:mm:ss") & cQuote & ", PA2360.DATE_LOG=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & " WHERE USERID=" & cQuote & gUserID & cQuote, oADORS, True
                        Script2File "UPDATE PA2360 SET STATUS=0, WSID=" & cQuote & gWSID & cQuote & ", PA2360.TIME=" & cQuote & Format(Now, "hh:mm:ss") & cQuote & ", PA2360.DATE_LOG=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & " WHERE USERID=" & cQuote & gUserID & cQuote
                        lSuccess = True
                           Exit For
                    End If
'----------->End
                    If oADORS("STATUS") = 0 Then
                    
                        gUserLevel = oADORS("USERLEVEL")
                        gUserName = oADORS("USERID") & " - " & oADORS("LASTNAME") & ", " & oADORS("FIRSTNAME")
                        gUserGroup = oADORS("groupid")
                        
                        OpenQueryDNS "UPDATE PA2360 SET STATUS=1, WSID=" & cQuote & gWSID & cQuote & ", PA2360.TIME=" & cQuote & Format(Now, "hh:mm:ss") & cQuote & ", PA2360.DATE_LOG=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & " WHERE USERID=" & cQuote & gUserID & cQuote, oADORS, True
                        Script2File "UPDATE PA2360 SET STATUS=1, WSID=" & cQuote & gWSID & cQuote & ", PA2360.TIME=" & cQuote & Format(Now, "hh:mm:ss") & cQuote & ", PA2360.DATE_LOG=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & " WHERE USERID=" & cQuote & gUserID & cQuote
                        
                        lSuccess = True
                        
                        Exit For
                    Else
                        Write2File gUserID & " is trying to logged-in at terminal " & gWSID & " at " & Now
                        Log2Audit "frmMain", gUserID & " is trying to logged-in at terminal " & gWSID
                    
                    
                        MsgBox "Please logged-off first on workstation #" & oADORS("WSID") & Chr$(13) & Chr$(10) & _
                               "Or you may call your System Administrator for assistance...", vbCritical, App.Title
                    End If
                Else
                    Log2Audit "frmMain", gUserID & " is trying to logged-in using an invalid user name or password"
                    MsgBox "Invalid User Name or Password.", vbError
                End If
            End If
            
        Else
            Hide
            Exit For
            
        End If
        
    Next nCtr
    
    Set oADORS = Nothing
        
    If lSuccess Then
    
        XPSideMenu1.Visible = lSuperUser
        If lSuperUser Then
            OpenQueryDNS "DELETE FROM DI82250 WHERE SYSNAME='PAYROLL'", objdbRs, True
            OpenQueryDNS "DELETE FROM DI82253 WHERE SYSNAME='PAYROLL'", objdbRs, True
            
            ChkTable
        End If
        
        If Dir(cImgFile) <> "" Then Me.Picture = LoadPicture(cImgFile)
        
        Write2File gUserID & " logged-in at terminal " & gWSID & " at " & Now
        Log2Audit "frmMain", "User logged-in."
    
        GetMenu Me
        GetUserMenu Me, gUserID
        StatusBar1.Panels(1).Text = gUserName
        StatusBar1.Panels(7).Text = "v" & PadStr(App.Major, "0", 2) & "." & PadStr(App.Minor, "0", 3) & "." & PadStr(App.Revision, "0", 4)
        
'        Timer1.Enabled = True
        
        frmSplash.Hide
        
        ' --> for deduction...
        If Not IfExists("PA3330", "DEDID='001'") Then       ' --> SSS Premium
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('001','SSS Premium',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        
        If Not IfExists("PA3330", "DEDID='002'") Then       ' --> SSS Loan
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('002','SSS Loan',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        
        If Not IfExists("PA3330", "DEDID='003'") Then       ' --> Pag-Ibig Premium
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('003','PAG-IBIG Premium',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        
        If Not IfExists("PA3330", "DEDID='004'") Then       ' --> Pag-Ibig Loan
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('004','PAG-IBIG Loan',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        
        If Not IfExists("PA3330", "DEDID='005'") Then       ' --> Medicare
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('005','Medicare',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        
        If Not IfExists("PA3330", "DEDID='006'") Then       ' --> Withholding Tax
            cSqlStmt = "INSERT INTO PA3330(DEDID,DEDNAME,FIX_DED)VALUES('006','W/holding Tax',1)"
            OpenQueryDNS cSqlStmt, objdbRs, True
        End If
        ' --> end deduction
        
        Show
        
        If lCheckFC Then
            ' --> check FC here...
            OpenQueryDNS "select * from di3670 where (active=0) and (emp_stat<>2) and (date_fin<=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & ")", objdbRs, False
            If objdbRs.RecordCount > 0 Then
                If MsgBox("Automate Employee to be finished?", vbYesNo + vbCritical, "Automate Finish Contract...") = vbYes Then
                    CheckFin_Active
                Else
                    frmManualFin.Show
                End If
            End If
        End If
        
        OpenQueryDNS "select * from PA73887", oTempADO, False
        If oTempADO.RecordCount = 0 Then
'        chkPA73887 = Array(Array("UL_AVAIL", "int(3)", 0, "'0'"), _
'                           Array("UL_USE", "int(3)", 0, "'0'"), _
'                           Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), _
'                           Array("COLA_AMT", "decimal(18,4)", 0, "0.0000"), _
'                           Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
            
            cSqlStmt = "insert into pa73887(ul_avail, ul_use, rate_amt, cola_amt)values(175, 0," & _
                       gBasicRate & "," & gColaAmt & ")"
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
        Else
            gBasicRate = oTempADO("rate_amt")
            gColaAmt = oTempADO("cola_amt")
        End If
        
    Else
        Write2File gUserID & "End Session"
        Log2Audit "frmMain", "Login failed."
        Add2List "Login failed!"
        Add2List "Exiting the system..."
        End
    End If
End Sub

Private Sub MDIForm_Load()
    Caption = App.Title
    
    XPSideMenu1.AddFrame "User", vbMain, vbOpen, True  ', ImageList1.ListImages(1).Picture
    XPSideMenu1.AddButton 0, "Log Off", 2400, xpHyperlink, True, True, , "Log Off as " & gUserName
    XPSideMenu1.AddButton 0, "Exit Application", 2400, xpHyperlink, True, True, , "End Application"
    XPSideMenu1.AddFrame "File Maintenance", vbMain, vbOpen, True  ', ImageList1.ListImages(1).Picture
    XPSideMenu1.AddButton 1, "Department Entry", 2400, xpHyperlink, True, True, , "Maintenance for Department Entry"
    XPSideMenu1.AddButton 1, "Employee Entry", 2400, xpHyperlink, True, True, , "Maintenance for Employee Entry"
    XPSideMenu1.AddFrame "Time Management System", vbMain, vbOpen, True   ', ImageList1.ListImages(2).Picture
    XPSideMenu1.AddButton 2, "Shifting Schedule", 2400, xpHyperlink, True, True, , "Create/Update Shifting Schedule by Line"
    XPSideMenu1.AddButton 2, "Time Management", 2400, xpHyperlink, True, True, , "Create/Update/Generate Daily Attendance Report"
    XPSideMenu1.AddButton 2, "Upload from Bio-clock", 2400, xpHyperlink, True, True, , "Uploading of data from stand-alone bio-clock"
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim nCtr, nCount As Integer
    For nCtr = 0 To Forms.Count - 1
        nCount = nCount + IIf(Forms(nCtr).Name <> "frmSplash", 1, 0)
        If Forms(nCtr).Visible Then Forms(nCtr).Show
    Next nCtr
    
    If nCount > 1 Then
        MsgBox "Please close all the other module to terminate...", vbOKOnly, App.Title
        Cancel = 1
    Else
        'Timer1.Enabled = False
        frmDialog.Show 1
        If ModalResult = mrOk Then
            If nAccess_Tag = 0 Then
                XPSideMenu1_Action 0, 0
                Cancel = IIf(lSuccess, 1, 0)
            End If
            If (nAccess_Tag <> 0) Or (Not lSuccess) Then
                XPSideMenu1_Action 0, 1
            End If
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub MDIForm_Terminate()
    oTempConn.Close
    objdbConn.Close
    Set oTempConn = Nothing
    Set objdbConn = Nothing
    Set objdbRs = Nothing
    Set oTempADO = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnnuDTRDet_Click()
     With frmSel
        .SetSelection 49
        .Caption = "Detailed Report"
        .Show
    End With
End Sub

Private Sub mnu13moAck_SLVL_Click()
    With frmSel
        .SetSelection 34
        .Caption = "SLVL Acknowledgement Report"
        .Show
    End With
End Sub

Private Sub mnu13moAcknowledgement_Click()
    With frmSel
        .SetSelection 22
        .Caption = "13th Month Acknowledgement Report"
        .Show
    End With
End Sub

Private Sub mnu13mobackup_Click()
    With frmSel
        ' update - exclude close 13th month period here, 20071206
        OpenQueryDNS "select periodid from pa7730 where 13month=1 and pclose=0", objdbRs, False
        .Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
        .SetSelection 21
        .Caption = "Backup 13th Month Pay Transaction"
        .Show
    End With
End Sub

Private Sub mnu13moDenom_Click()
    With frmSel
        .SetSelection 23
        .Caption = "13th Month Denomination Report"
        .Show
    End With
End Sub

Private Sub mnu13moDenom_SLVL_Click()
    With frmSel
        .SetSelection 35
        .Caption = "SLVL Denomination Report"
        .Show
    End With
End Sub

Private Sub mnu13moGen_Click()
    frmGen13mo.Show
End Sub

Private Sub mnu13mosheet_Click()
    With frmSel
        ' update - exclude close 13th month period here, 20071206
        OpenQueryDNS "select periodid from pa7730 where 13month=1 and pclose=0", objdbRs, False
        .Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
        .SetSelection 20
        .Caption = "13th Month Payroll Sheet Print Utility"
        .Show
    End With
End Sub

Private Sub mnu13mosheet_SLVL_Click()
    With frmSel
        ' update - exclude close 13th month period here, 20071206
        OpenQueryDNS "select periodid from pa7730 where 13month=1 and pclose=0", objdbRs, False
        .Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
        .SetSelection 33
        .Caption = "SLVL Payroll Sheet Print Utility"
        .Show
    End With
End Sub

Private Sub mnu13moslip_Click()
    With frmSel
        ' update - exclude close 13th month period here, 20071206
        OpenQueryDNS "select periodid from pa7730 where 13month=1 and pclose=0", objdbRs, False
        .Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
        .SetSelection 19
        .Caption = "13th Month Payslip Print Utility"
        .Show
    End With
End Sub

Private Sub mnu13moslip_SLVL_Click()
    With frmSel
        ' update - exclude close 13th month period here, 20071206
        OpenQueryDNS "select periodid from pa7730 where 13month=1 and pclose=0", objdbRs, False
        .Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
        .SetSelection 32
        .Caption = "SLVL Payslip Print Utility"
        .Show
    End With
    
'    Dim cSqlStmt, _
'        cParam, _
'        cPeriodName, _
'        cPeriodID, _
'        oRset1 As New ADODB.Recordset, _
'        nActive, _
'        nCtr, nFilter As Integer, _
'        aOtherInfo As Variant, _
'        aDedName As Variant, aDedAmt As Variant, _
'        aMonthName As Variant
'
'    nFilter = 0
'    ' --> process active employee first here...
'    nActive = 0
'
'    aMonthName = Array("Enero", "Pebrero", "Marso", "Abril", "Mayo", "Hunyo", "Hulyo", "Agosto", "Setyembre", "Oktubre", "Nobyembre", "Disyembre")
'
'    If Trim(cParam) <> "" Then
'        cParam = "a.depid IN " & cParam
'    End If
'
'    CreateTmpPaySlip 0  ' --> header
'    CreateTmpPaySlip 1 ' --> detail
'
'loopd2:
'
'    aOtherInfo = Array("", "", "", "")
'
'    OpenQueryDNS "select * from pa7730 where 13month=1 and pclose=0", objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        cPeriodID = objdbRs("periodid")
'        cPeriodName = "For the Payroll period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
'    End If
'    cSqlStmt = ""
'
'    cSqlStmt = " select a.periodid, a.seq_no, " & _
'               " a.empid, a.firstname, a.lastname, a.mname, a.emp_stat,  a.fullname, " & _
'               " a.depid, a.rate_amt, " & _
'               " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as LEAVE_PAY, " & _
'               " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as gross_pay, " & _
'               " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as net_pay, " & _
'               " a.active " & _
'               " from pa87260 a " & _
'               " left join di3670 b on a.empid=b.empid"
'
'    cSqlStmt = cSqlStmt & " where (a.active" & IIf(nActive = 0, "=0", "<>0") & ")" & IIf(Trim(cParam) = "", "", " and (" & cParam & ")") & _
'               IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
'               " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'               " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")"
''    MsgBox cSqlStmt
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oTempADO, False
'    If oTempADO.RecordCount > 0 Then
'
''        OpenQueryDNS "select posid, posname from DI7670 order by posid", oRSet1, False
'        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRset1, False
'
'        ShowProgress 0
'
'        While Not oTempADO.EOF
'
'            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
'
'            '---> need to kasama to
'            If oRset1.RecordCount > 0 Then
'                oRset1.Requery adAsyncFetch
'                oRset1.Find "lineid='" & oTempADO("depid") & "'"
'                aOtherInfo(1) = IIf(oRset1.EOF, "", oRset1("linename"))
'            End If
'
''            cSqlStmt = "insert into tmp7297655(periodname, seq_no, depid, deptname," & _
''                       IIf(nActive = 1, " depid2, deptname2,", "") & _
''                       " empid, emp_stat, [active], fullname, fname, mname, lname, " & _
''                       " rate_amt, c13mo_pay, gross_pay, net_pay)values(" & _
''                       cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("seq_no") & "," & _
''                       cQuote & IIf(nActive = 1, "999", cQuote & oTempADO("depid") & cQuote & "," & _
''                       cQuote & IIf(nActive = 1, "Resigned/FC", cQuote & DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
''                       cQuote & IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & ",", "") & _
''                       cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
''                       cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("mname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
''                       oTempADO("rate_amt") & "," & oTempADO("M13PAY") & "," & oTempADO("gross_pay") & "," & _
''                       oTempADO("net_pay") & ")"
'
'            cSqlStmt = "insert into tmp7297655(periodname, seq_no, depid, deptname," & _
'                       IIf(nActive = 1, " depid2, deptname2,", "") & _
'                       " empid, emp_stat, [active], fullname, fname, mname, lname, " & _
'                       " rate_amt, LEAVE_PAY, gross_pay, net_pay)values(" & _
'                       cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("seq_no") & "," & _
'                       IIf(nActive = 1, "999", cQuote & oTempADO("depid") & cQuote) & "," & _
'                       IIf(nActive = 1, "Resigned/FC", cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote) & "," & _
'                       IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
'                       cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
'                       cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("mname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
'                       oTempADO("rate_amt") & "," & oTempADO("LEAVE_PAY") & "," & oTempADO("gross_pay") & "," & _
'                       oTempADO("net_pay") & ")"
''            MsgBox cSqlStmt
'            QueryTemp cSqlStmt, objdbRs, True
'
'            oTempADO.MoveNext
'        Wend
'
'        ShowProgress 4
'    End If
'
'    ' --> process resigned/fc employee next...
'    If nActive = 0 Then
'        nActive = 1
'        GoTo loopd2
'    End If
'
'    ShowProgress 0
'
'    QueryTemp "select * from tmp7297655 ", objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'
'        ShowProgress 3
'        GenerateReport "SLVL PAYSLIP", "rpt7297547.rpt"
'
'    Else
'        ShowProgress 3
'        MsgBox "No report to generate!", vbCritical, "System Advisory"
'    End If
'
'    ShowProgress 4
'
'    Set oRset1 = Nothing
End Sub

Private Sub mnuAccRight_Click()
    frmAccess.Show
End Sub

Private Sub mnuAcknowledgeSheet_Click()
    With frmSel
        .SetSelection 5
        .Caption = "Acknowledgement Report"
        .Show
    End With
End Sub

Private Sub mnuAcknowledgeSheetInc_Click()
    With frmSel
        .SetSelection 30
        .Caption = "Acknowledgement Report"
        .Show
    End With
End Sub

Private Sub mnuAdmin_Click()
    GetUserRights PadStr(mnuAdmin.Name, " ", 100, PadRight), gUserID
    frmAdmin.Show
End Sub

Private Sub mnuAlpha_Click()
    With frmSel
        .SetSelection 24
        .Caption = "Generate Alphalist"
        .Show
    End With
'    Dim cSqlStmt As String, _
'        cPeriodID As String, _
'        oRecordSet As New ADODB.Recordset
'
'    If MsgBox("Generate Alphalist?", vbYesNo, "System Advisory!!!") = vbYes Then
'        OpenQueryDNS "select periodid from pa7730 where pclose=0 and wtax=1", objdbRs, False
'        cPeriodID = IIf(objdbRs.RecordCount > 0, objdbRs("periodid"), "")
'
'        cSqlStmt = "sele"
'    End If
'
'    Set oRecordSet = Nothing
End Sub



'Private Sub mnuAudit_Click()
'    Form1.Show
'End Sub

Private Sub mnuBackup_Click()
    With frmSel
        .SetSelection 17
        .Caption = "Backup Payroll Transaction"
        .Show
    End With
End Sub


Private Sub mnuBlocklist_Click()
 GetUserRights PadStr(mnuBlocklist.Name, " ", 100, PadRight), gUserID
    frmBlockList.Show
End Sub

'Private Sub mnuBioClock_Click()
'    GetUserRights PadStr(mnuBioClock.Name, " ", 100, PadRight), gUserID
'    frmBioClock.Show
'End Sub

Private Sub mnuClosePeriod_Click()
    With frmSel
        .SetSelection 8
        .Caption = "Close Entry"
        .Show
    End With
End Sub

Private Sub mnuDeduction_Click()
    GetUserRights PadStr(mnuDeduction.Name, " ", 100, PadRight), gUserID
    frmDeduction.Show
End Sub

Private Sub mnuDen_res_Click()
    With frmSel
        .SetSelection 41
        .Caption = "Denomination For Emergency and Resign Report"
        .Show
    End With
End Sub

Private Sub mnuDenomination_Click()
    With frmSel
        .SetSelection 7
        .Caption = "Denomination Report"
        .Show
    End With
End Sub

Private Sub mnuDepShiftSched_Click()
    GetUserRights PadStr(mnuDepShiftSched.Name, " ", 100, PadRight), gUserID
    frmShiftSched2.Show
End Sub

Private Sub mnuDept_Click()
    GetUserRights PadStr(mnuDept.Name, " ", 100, PadRight), gUserID
    frmline.Show
End Sub

Private Sub mnuDLCostRpt_Click()
      With frmReport
        .GetRpt 9
        .Show
    End With

End Sub

Private Sub mnuDTROldBackup_Click()
Dim cSqlStmt As String, _
    oRecordSet As New ADODB.Recordset
    
    ShowProgress 0
    
    'att2000h-att2000hh
    cSqlStmt = " insert into att2000hh (empid, BCID, TCID, TRANSDATE, TRANTIME, trantype, logid, TAG, id) " & _
               " select empid, BCID, TCID, TRANSDATE, TRANTIME, trantype, logid, TAG, id from att2000h where year(transdate) < year(current_date)-1"
    OpenQueryDNS cSqlStmt, oRecordSet, True
    
    cSqlStmt = " delete FROM att2000h where year(transdate) < year(current_date)-1 "
    OpenQueryDNS cSqlStmt, oRecordSet, True

    'dih36770-dihh36770
    cSqlStmt = " insert into dihh36770(EMPID, PERIODID, DATE, SHIFTID, DESCRIPTION, TIME1, TIME2, reg_hr, reg_ot_hr, sa_reg_ot, nd_hr, nd_ot_hr, sa_nd_ot, sun_hr, sun_ot_hr, REMARK, CMPID, allowance, TAG, sun_nd, sun_nd_ot, Inc_hr, tot_ot, nd_tot_ot) " & _
               " SELECT EMPID, PERIODID, DATE, SHIFTID, DESCRIPTION, TIME1, TIME2, reg_hr, reg_ot_hr, sa_reg_ot, nd_hr, nd_ot_hr, sa_nd_ot, sun_hr, sun_ot_hr, REMARK, CMPID, allowance, TAG, sun_nd, sun_nd_ot, Inc_hr, tot_ot, nd_tot_ot FROM dih36770 where year(date) < year(current_date)-1 "
    OpenQueryDNS cSqlStmt, oRecordSet, True
    
    cSqlStmt = " delete FROM dih36770 where year(date) < year(current_date)-1 "
    OpenQueryDNS cSqlStmt, oRecordSet, True

    'pah84650-pahh84650
    cSqlStmt = " insert into pahh84650(TRAN_NO, BCID, TCID, EMPID, SHIFTID, LOGDATE, TRANSDATE, TRANTIME, TRANTYPE, TAG, SWAPDATE, CMPID, ID) " & _
               " select TRAN_NO, BCID, TCID, EMPID, SHIFTID, LOGDATE, TRANSDATE, TRANTIME, TRANTYPE, TAG, SWAPDATE, CMPID, ID from pah84650 where year(logdate) < year(current_date)-1 "
    OpenQueryDNS cSqlStmt, oRecordSet, True
    
    cSqlStmt = " delete from pah84650 where year(logdate) < year(current_date)-1 "
    OpenQueryDNS cSqlStmt, oRecordSet, True
    
    ShowProgress 4

Set oRecordSet = Nothing

End Sub

Private Sub mnuDTROldRep_Click()
' With frmSel
'        .SetSelection 43
'        .Caption = "Old DTR Report Listing"
'        .Show
'    End With

'    GetUserRights PadStr(mnuDTROldRep.Name, " ", 100, PadRight), gUserID
'    frmTMS.Show
'    frmTMSOLD.Show
End Sub

Private Sub mnuDTRSum_Click()
     With frmSel
        .SetSelection 48
        .Caption = "Summary Report"
        .Show
    End With
End Sub

Private Sub mnuEmpInc_Click()
    GetUserRights PadStr(mnuEmpInc.Name, " ", 100, PadRight), gUserID
    frmEmpIncentive.Show
End Sub

Private Sub mnuEmpLeave_Click()
    GetUserRights PadStr(mnuEmpLeave.Name, " ", 100, PadRight), gUserID
    frmEmpSLVL.Show
End Sub

Private Sub mnuEmpLoan_Click()
    GetUserRights PadStr(mnuEmpLoan.Name, " ", 100, PadRight), gUserID
    frmLoan.Show
End Sub

Private Sub mnuEmpLoanRpt_Click()
    With frmSel
        .SetSelection 31
        .Caption = "Employee Loan Report"
        .Show
    End With
End Sub

Private Sub mnuEmployee_Click()
    GetUserRights PadStr(mnuEmployee.Name, " ", 100, PadRight), gUserID
    frmEmployee.Show
End Sub

Private Sub mnuEmpLvl_Click()
    GetUserRights PadStr(mnuEmpLvl.Name, " ", 100, PadRight), gUserID
    frmEmpLevel.Show
End Sub

Private Sub mnuEmpOt_Click()
    With frmReport
        .GetRpt 3
        .Show
    End With
End Sub

Private Sub mnuEmpPHRep_Click()
    With frmSel
        .SetSelection 42
        .Caption = "PHILHEALTH Er2 REPORT"
        .Show
    End With
End Sub

Private Sub mnuEmpShiftSched_Click()
    GetUserRights PadStr(mnuEmpShiftSched.Name, " ", 100, PadRight), gUserID
    frmShiftSchedEmp.Show
End Sub

Private Sub mnuEmpSSRep_Click()
    With frmSel
        .SetSelection 26
        .Caption = "SSS R-1A"
        .Show
    End With
End Sub

Private Sub mnuERP_COMP_Click()
    GetUserRights PadStr(mnuERP_COMP.Name, " ", 100, PadRight), gUserID
    frmCompany.Show
End Sub

Private Sub mnuERP_Cost_Center_Click()
    GetUserRights PadStr(mnuERP_Cost_Center.Name, " ", 100, PadRight), gUserID
    frmCostCenter.Show
End Sub

Private Sub mnuERP_EMPL_Click()
    With frmReport
        .GetRpt 7
        .Show
    End With
End Sub

Private Sub mnuERP_SALG_Click()

    '----> ERP Salary Report Gen 2012-09-10
    With frmSel
        .SetSelection 50
        .Caption = "Salaray Generation Report"
        .Show
    End With
End Sub

Private Sub mnuERP_TMS_Click()
    With frmReport
        .GetRpt 8
        .Show
    End With
End Sub

Private Sub mnuERP_Work_Center_Click()
    GetUserRights PadStr(mnuERP_Work_Center.Name, " ", 100, PadRight), gUserID
    frmWorkCenter.Show
End Sub

Private Sub mnuExit_Click()
    XPSideMenu1_Action 0, 1
End Sub

Private Sub mnuGenBCID_Click()
'    frmConnect.Show
    frmGenBCID.Show
End Sub

'Private Sub mnuGroup_Click()
'    frmAutomate.Show
'End Sub

Private Sub mnuHoliday_Click()
    GetUserRights PadStr(mnuHoliday.Name, " ", 100, PadRight), gUserID
    frmHoliday.Show
End Sub

Private Sub mnuLeave_Click()
    GetUserRights PadStr(mnuLeave.Name, " ", 100, PadRight), gUserID
    frmLeave.Show
End Sub

Private Sub mnuLevel_Click()
    GetUserRights PadStr(mnuLevel.Name, " ", 100, PadRight), gUserID
    frmlevel.Show
End Sub

Private Sub mnuLogOff_Click()
    XPSideMenu1_Action 0, 0
End Sub

Private Sub mnuMasterData_Click()
'      With frmReport
'        .GetRpt 10
'        .Show
'   End With
   
     With frmSel
        .SetSelection 51
        .Caption = "Employee Master Data Generation"
        .Show
    End With
End Sub

Private Sub mnuMonthlyFC_Click()
    With frmSel
        .SetSelection 9
        .Caption = "Monthly Finish Contract "
        .Show
    End With
End Sub

Private Sub mnuPayIncSht2_Click()
    With frmSel
        .SetSelection 29
        .Caption = "Resigned/FC Incentive Payroll Sheet Report"
        .Show
    End With
End Sub

Private Sub mnuPayIncSlip_Click()
    With frmSel
        .SetSelection 27
        .Caption = "Incetive Pay Slip Print Utility"
        .Show
    End With
End Sub

Private Sub mnuPayPeriod_Click()
    GetUserRights PadStr(mnuPayPeriod.Name, " ", 100, PadRight), gUserID
    frmPeriod.Show
End Sub

Private Sub mnuPayrollIncSheet_Click()
    With frmSel
        .SetSelection 28
        .Caption = "Incentive Payroll Sheet Report"
        .Show
    End With
End Sub

Private Sub mnuPayrollSheet_Click()
    With frmSel
        .SetSelection 6
        .Caption = "Payroll Sheet Report"
        .Show
    End With
End Sub

Private Sub mnuPaySlip_Click()
    With frmSel
        .SetSelection 4
        .Caption = "Pay Slip Print Utility"
        .Show
    End With
End Sub

Private Sub mnuPeriodFC_Click()
    With frmSel
        .SetSelection 10
        .Caption = "Finish Contract by Period"
        .Show
    End With
End Sub

Private Sub mnuPEZA_Click()
    frmSel.SetSelection 2
    frmSel.Caption = "PEZA Report Utility"
    frmSel.Show
End Sub

Private Sub mnuPhilHealth_Click()
    GetUserRights PadStr(mnuPhilHealth.Name, " ", 100, PadRight), gUserID
    frmPhilhealth.Show
End Sub

Private Sub mnuPosition_Click()
    GetUserRights PadStr(mnuPosition.Name, " ", 100, PadRight), gUserID
    frmPosition.Show
End Sub

Private Sub mnuRCBCGDExcell_Click()
'    With frmSel
'        .SetSelection 37
'        .Caption = "RCBC-Genarate EXcell"
'        .Show
'    End With
End Sub

Private Sub mnuRCBCGDIX1_Click()
'    With frmSel
'        .SetSelection 38
'        .Caption = "RCBC-Genarate Data"
'        .Show
'    End With
End Sub

Private Sub mnuPPRESFCAckShit_Click()
    With frmSel
        .SetSelection 45
        .Caption = "Payroll Sheet Report"
        .Show
    End With
End Sub

Private Sub mnuPPRESFCPayShit_Click()
    With frmSel
        .SetSelection 15
        .Caption = "Resigned/FC Payroll Sheet Report"
        .Show
    End With
End Sub

Private Sub mnuPPRESFCPaySlip_Click()
    With frmSel
        .SetSelection 44
        .Caption = "Pay Slip Print Utility"
        .Show
    End With
End Sub

Private Sub mnuPPRESFCSalDiv_Click()
    With frmSel
        .SetSelection 46
        .Caption = "Salary Division Report for Resigned/FC"
        .Show
    End With
End Sub

Private Sub mnuRCBCRPTTrans_Click()
    With frmSel
        .SetSelection 39
        .Caption = "Transmittal Sheet"
        .Show
    End With
End Sub

Private Sub mnuRebuildDed_Click()
    Dim cSqlStmt As String
    If MsgBox("Would you like to rebuild Employee's deduction entry?", vbYesNo + vbCritical, "System Advisory") = vbYes Then
        OpenQueryDNS "SELECT EMPID, CONCAT(LASTNAME,' ',FIRSTNAME) AS FULLNAME FROM DI3670", oTempADO, False
        If oTempADO.RecordCount > 0 Then
            OpenQueryDNS "DELETE FROM DI3673", objdbRs, True
            Script2File "DELETE FROM DI3673"
            
            ShowProgress 0, , 100
            While Not oTempADO.EOF
                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , , "Building Deduction entry for " & oTempADO("FULLNAME")
'            OpenQueryDNS "SELECT DEDID, DEF_AMT, CUT_OFF_AMT, PERIOD1, PERIOD2 FROM PA3330", oTempADO, False
                cSqlStmt = "INSERT INTO DI3673(EMPID,DEDID,DEF_AMT,CUT_OFF_AMT,PERIOD1,PERIOD2)" & _
                           " SELECT " & cQuote & oTempADO("EMPID") & cQuote & ",DEDID, DEF_AMT, CUT_OFF_AMT, PERIOD1, PERIOD2 FROM PA3330"
                OpenQueryDNS cSqlStmt, objdbRs, True
                oTempADO.MoveNext
            Wend
            ShowProgress 4
        End If
    End If
End Sub


Sub CreateTmpDed()
    On Error GoTo ErrCreate
    Dim cParam, _
        cSqlStmt As String, _
        nFieldCnt As Integer

    cParam = ""
    For nFieldCnt = 1 To 24
        cParam = cParam & "[ded_amt" & nFieldCnt & "] double,"
    Next nFieldCnt
    
    cSqlStmt = "create table tmpDeduction(" & _
               "    [empid] char(6),        [fullname] char(100)," & _
               "    [tot_ded] double," & cParam & _
               "    [leave_pay] double,     [m13pay] double)"
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpDeduction"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Private Sub mnuRebuildEmp_Click()
    Dim cSqlStmt As String, _
        cDedID As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset
    
'    CreateTmpDed
'
'    For nCtr = 0 To UBound(aTaxExempt)
'        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
'        cDedID = cDedID & aTaxExempt(nCtr) & ","
'    Next nCtr
'    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
'
'    cSqlStmt = "select periodid from pa7730 where year(date_end)=2006 and 13month=0"
'    OpenQueryDNS cSqlStmt, oTempADO, False
'    If oTempADO.RecordCount > 0 Then
'        While Not oTempADO.EOF
'            oTempADO.MoveNext
'        Wend
'    End If
'
'    Set oRecordSet = Nothing
    
    If MsgBox("Would you like to rebuild Employee master file?", vbYesNo + vbCritical, "System Advisory") = vbYes Then
        ShowProgress 0
        cSqlStmt = "select a.empno, a.lastname, a.firstname, ifnull(a.minitial,'') as minitial, a.birthday, if(a.sex='M',0,1) as sex, ifnull(a.pagibigno,'') as pagibigno," & _
                   " lpad(if(a.resigned is null,a.department,left(a.empno,2)),3,'0') as department, lpad(a.position,3,'0') as `position`, a.rate, a.posallow," & _
                   " if(a.resigned is null,0,1) as active, if(a.empstatus='W',0,if(a.empstatus='C',1,2)) as empstat, if(a.paystatus='D',0,1) as paystat," & _
                   " ifnull(a.sssnumber,'') as ssnum,if(a.union='Y',0,1) as `union`, ifnull(a.tin,'') as tin, a.taxcode," & _
                   " ifnull(a.datehired,curdate()) as datehired, ifnull(a.hiredate,curdate()) as hiredate," & _
                   " if(a.resigned is null,curdate(),a.resigned) as date_res " & _
                   "from masterk1 a"
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then
            OpenQueryDNS "delete from di3670", objdbRs, True
            While Not oTempADO.EOF
                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
                cSqlStmt = "insert into di3670(empid, lastname, firstname, mname, birthday, sex, pagibigno, " & _
                           " depid, posid, rate_amt, pos_allow, active, emp_stat, paystatus," & _
                           " ssnum, isunion, tin, taxcode, date_hire, date_res, date_fin)values(" & _
                           cQuote & oTempADO("empno") & cQuote & "," & _
                           cQuote & EncodeStr(oTempADO("lastname")) & cQuote & "," & cQuote & EncodeStr(oTempADO("firstname")) & cQuote & "," & cQuote & EncodeStr(oTempADO("minitial")) & cQuote & "," & _
                           cQuote & Format(oTempADO("birthday"), "yyyy-mm-dd") & cQuote & "," & oTempADO("sex") & "," & _
                           cQuote & oTempADO("pagibigno") & cQuote & "," & cQuote & oTempADO("department") & cQuote & "," & _
                           cQuote & oTempADO("position") & cQuote & "," & oTempADO("rate") & "," & oTempADO("posallow") & "," & _
                           oTempADO("active") & "," & oTempADO("empstat") & "," & oTempADO("paystat") & "," & _
                           cQuote & oTempADO("ssnum") & cQuote & "," & oTempADO("union") & "," & _
                           cQuote & oTempADO("tin") & cQuote & "," & cQuote & oTempADO("taxcode") & cQuote & "," & _
                           cQuote & Format(oTempADO("datehired"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_res"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(DateAdd("m", IIf(oTempADO("empstat") = 0, 3, IIf(oTempADO("empstat") = 1, 5, 0)), oTempADO("datehired")), "yyyy-mm-dd") & cQuote & ")"

'                   " date_add(ifnull(a.datehired,curdate()), INTERVAL if(a.empstatus='W',3,if(a.empstatus='C',5,0)) MONTH) as date_fin," & _

'                MsgBox cSqlStmt
                OpenQueryDNS cSqlStmt, objdbRs, True
                oTempADO.MoveNext
            Wend
        End If
        ShowProgress 4
    End If
End Sub

Private Sub mnuRebuildYTD_Click()
    Dim cSqlStmt As String
    
    If MsgBox("Rebuild YTD Info?", vbYesNo, "System Advisory!!!") = vbYes Then
        cSqlStmt = "update di3670 set ytd_gross = 0, " & _
                   "  ytd_cola=0, " & _
                   "  ytd_basic=0, " & _
                   "  ytd_wtax=0, " & _
                   "  ytd_gross_sa=0"
        OpenQueryDNS cSqlStmt, objdbRs, True
        Script2File cSqlStmt
        
        cSqlStmt = "select a.empid, " & _
                   "  a.gross_pay-a.m13pay as gross_pay, " & _
                   "  (a.cola_amt*(a.reg_day+a.ndiff_day)) as cola_amt, " & _
                   "  a.reg_pay+a.ndiff_pay as basic_pay, " & _
                   "  a.wtax, " & _
                   "  a.sa_net_pay " & _
                   "from pah87260 a where a.periodid in (select periodid from pa7730 where year(date_start)=2007)"
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then
            ShowProgress 0
            While Not oTempADO.EOF
                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
                cSqlStmt = "update di3670 set " & _
                           "  ytd_gross=ytd_gross+" & oTempADO("gross_pay") & ", " & _
                           "  ytd_cola=ytd_cola+" & oTempADO("cola_amt") & ", " & _
                           "  ytd_basic=ytd_basic+" & oTempADO("basic_pay") & ", " & _
                           "  ytd_wtax=ytd_wtax+" & oTempADO("wtax") & ", " & _
                           "  ytd_gross_sa=ytd_gross_sa+" & oTempADO("sa_net_pay") & _
                           " where empid=" & cQuote & oTempADO("empid") & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                oTempADO.MoveNext
            Wend
            ShowProgress 4
            MsgBox "Done"
        End If
    End If
End Sub

Private Sub mnuRemitMedicare_Click()
    frmSel.SetSelection 12
    frmSel.Caption = "Philhealth/Medicare Monthly Remittance Report"
    frmSel.Show
End Sub

Private Sub mnuRemitPagIbig_Click()
    frmSel.SetSelection 14
    frmSel.Caption = "Pag-Ibig Remittance Report"
    frmSel.Show
End Sub

Private Sub mnuRemitSSS_Click()
    frmSel.SetSelection 16
    frmSel.Caption = "SSS Premium Remittance Report"
    frmSel.Show
End Sub

Private Sub mnuRemitTax_Click()
    frmSel.SetSelection 13
    frmSel.Caption = "Withholding Tax Report"
    frmSel.Show
End Sub

Private Sub mnuResetMnu_Click()
    If MsgBox("Are you sure you want to reset system menu?", vbYesNo, App.Title) = vbYes Then GetMenu Me, True
End Sub

Private Sub mnuRptActMan_Click()
    With frmReport
        .GetRpt 5
        .Show
    End With
End Sub

Private Sub mnuRptDed_Click()
    On Error GoTo ErrDedRpt
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String

    CreateTemp 5
    
    OpenQueryDNS "select a.dedid,a.dedname,a.def_amt,a.cut_off_amt,a.fix_ded,a.period1,a.period2,a.auto_ded,b.cmpname FROM pa3330 a left join di2660 b on a.cmpid=b.cmpid ", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0, , oRecordSet.RecordCount
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " INSERT INTO tmpDed(DEDID,DEDNAME,DEF_AMT,CUT_OFF_AMT,FIX_DED,PERIOD1,PERIOD2,AUTO_DED,CMPName)VALUES(" & _
                        cQuote & oRecordSet("DEDID") & cQuote & "," & cQuote & oRecordSet("DEDNAME") & cQuote & "," & _
                        oRecordSet("DEF_AMT") & "," & oRecordSet("CUT_OFF_AMT") & "," & oRecordSet("FIX_DED") & "," & _
                        oRecordSet("PERIOD1") & "," & oRecordSet("PERIOD2") & "," & oRecordSet("AUTO_DED") & "," & _
                        cQuote & oRecordSet("CMPNAME") & cQuote & ")"

'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        
        GenerateReport "Deduction Listing Report", "LST3330.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndDedRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrDedRpt:
    ErrorMsg Err.Number, Err.Description, "Deduction Listing", Name
    Resume EndDedRpt
End Sub

Private Sub mnuRptDep_Click()
    On Error GoTo ErrDepRpt
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String

    CreateTemp 6
    
    OpenQueryDNS "select a.lineid,ifnull(a.linename,'') as linename,ifnull(a.production,'') as production, b.cmpname from di5463 a left join di2660 b on a.cmpid=b.cmpid ", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " INSERT INTO tmpDep(LineID,LineName,production,CMPName)VALUES(" & _
                       cQuote & oRecordSet("LINEID") & cQuote & "," & cQuote & oRecordSet("LINENAME") & cQuote & "," & _
                       oRecordSet("production") & "," & cQuote & oRecordSet("CMPNAME") & cQuote & ")"

'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        
        GenerateReport "Department Report Listing", "LST5463.RPT", , True
        
        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndDepRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrDepRpt:
    ErrorMsg Err.Number, Err.Description, "Department Listing", Name
    Resume EndDepRpt
End Sub

Private Sub mnuRptDTCons_Click()
    With frmReport
        .GetRpt 4
        .Show
    End With
End Sub

Private Sub mnuRptDTRPAna_Click()
 With frmSel
        .SetSelection 40
        .Caption = "Payroll Analysis Report"
        .Show
    End With
End Sub

Private Sub mnuRptDTRRep_Click()
    frmDTRReport.Show
End Sub

Private Sub mnuRptEmp_Click()
    With frmSel
        .SetSelection 3
        .Caption = "Employee Report Listing"
        
        .Label2.Caption = "Select filter:"
        .Combo1.Clear
        .Combo1.AddItem "Active Employee only"
        .Combo1.AddItem "Resigned/Finished Employee only"
        .Combo1.ListIndex = 0
        .Show
    End With
End Sub

Private Sub mnuRptHoliday_Click()
    On Error GoTo ErrHolRpt
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String

    CreateTemp 1
    
    OpenQueryDNS "select a.holidayid,a.date,a.description,a.fix_day,b.cmpname FROM pa4329 a left join di2660 b on a.cmpid=b.cmpid ", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " INSERT INTO tmpHoliday(HOLIDAYID,DATE_HOL,DESCRIPTION,FIX_DAY,CMPName)VALUES(" & _
                        cQuote & oRecordSet("HOLIDAYID") & cQuote & "," & cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                        cQuote & oRecordSet("DESCRIPTION") & cQuote & "," & oRecordSet("FIX_DAY") & "," & cQuote & oRecordSet("CMPNAME") & cQuote & ")"

'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        
        GenerateReport "Holiday Listing Report", "LST4329.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndHolRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrHolRpt:
    ErrorMsg Err.Number, Err.Description, "Holiday Listing", Name
    Resume EndHolRpt
End Sub

Private Sub mnuRptLeave_Click()
    frmSel.SetSelection 11
    frmSel.Caption = "Leave Report"
    frmSel.Show
End Sub

Private Sub mnuRptPeriod_Click()
    On Error GoTo ErrPPeriodRpt
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        aOtherInfo As Variant
        
    aOtherInfo = Array("", "")

    CreateTemp 7
        
    cSqlStmt = "SELECT periodid,date_start,date_end,Duration,pclose,date_close,workindays,holidays,status,cmpid FROM pa7730 "
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
                   
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving Period ID" & oRecordSet("periodid")
            
            aOtherInfo(0) = IIf(oRecordSet("pclose") = 0, "", "Close")
            
            OpenQueryDNS "select * from di2660 where cmpid = " & gCompanyID, objdbRs, False
            aOtherInfo(1) = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")

            cSqlStmt = " INSERT INTO tmpPPeriod(PERIODID,DATE_START,DATE_END,DURATION,PCLOSE,PCLOSENAME,DATE_CLOSE,WORKINDAYS,HOLIDAYS,STATUS,CMPNAME)VALUES(" & _
                       cQuote & oRecordSet("periodid") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_start"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_end"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & oRecordSet("duration") & cQuote & "," & _
                       oRecordSet("pclose") & "," & _
                       cQuote & aOtherInfo(0) & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_close"), "yyyy-mm-dd") & cQuote & "," & _
                       oRecordSet("workindays") & "," & _
                       oRecordSet("holidays") & "," & _
                       oRecordSet("status") & "," & _
                       cQuote & aOtherInfo(1) & cQuote & ")"
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport "Payroll Period Report Listing", "LST7730.rpt", , True
        
        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndPPeriodRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrPPeriodRpt:
    ErrorMsg Err.Number, Err.Description, "Payroll Period Listing", Name
    Resume EndPPeriodRpt
End Sub

Private Sub mnuRptPos_Click()
    On Error GoTo ErrPosRpt
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String

    CreateTemp 2
    
    OpenQueryDNS "select a.posid,ifnull(a.posname,'') as posname, a.staff, a.allowance, b.cmpname  FROM di7670 a left join di2660 b on a.cmpid=b.cmpid ", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " INSERT INTO tmppos(POSID,POSNAME,STAFF,ALLOWANCE,CMPName)VALUES(" & _
                        cQuote & oRecordSet("POSID") & cQuote & "," & cQuote & oRecordSet("POSNAME") & cQuote & "," & _
                        oRecordSet("STAFF") & "," & oRecordSet("ALLOWANCE") & "," & cQuote & oRecordSet("CMPNAME") & cQuote & ")"
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        
        GenerateReport "Position Listing Report", "LST7670.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndPosRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrPosRpt:
    ErrorMsg Err.Number, Err.Description, "Position Listing", Name
    Resume EndPosRpt
End Sub

Private Sub mnuRptShift_Click()
    On Error GoTo ErrShiftRpt
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String

    CreateTemp 3
    
    OpenQueryDNS "select a.shiftid,a.description,a.time1,a.time2,a.remark,a.ndiff,a.allowance,a.reg_hr,a.btime,a.default,b.cmpname FROM pa74380 a left join di2660 b on a.cmpid=b.cmpid", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " INSERT INTO tmpShift(SHIFTID,DESCRIPTION,[TIME1],[TIME2],REMARK,NDIFF,[ALLOWANCE],REG_HR,BTIME,[DEFAULT],CMPName)VALUES(" & _
                        cQuote & oRecordSet("SHIFTID") & cQuote & "," & cQuote & oRecordSet("DESCRIPTION") & cQuote & "," & _
                        cQuote & Format(oRecordSet("TIME1"), "hh:mm AMPM") & cQuote & "," & _
                        cQuote & Format(oRecordSet("TIME2"), "hh:mm AMPM") & cQuote & "," & _
                        cQuote & oRecordSet("REMARK") & cQuote & "," & oRecordSet("NDIFF") & "," & _
                        oRecordSet("ALLOWANCE") & "," & oRecordSet("REG_HR") & "," & oRecordSet("BTIME") & "," & _
                        oRecordSet("DEFAULT") & "," & cQuote & oRecordSet("CMPNAME") & cQuote & ")"

'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        
        GenerateReport "Position Listing Report", "LST74380.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndShiftRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrShiftRpt:
    ErrorMsg Err.Number, Err.Description, "Shift Listing", Name
    Resume EndShiftRpt
End Sub

Private Sub mnuRptWTax_Click()
    On Error GoTo ErrWTaxRpt
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        nCtr As Integer, _
        cCmpname As String, _
        aOtherInfo As Variant, aOtherInfo1 As Variant, _
        aOtherinfo2 As Variant
    
    aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aOtherInfo1 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aOtherinfo2 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
                
    CreateTemp 4
        
    'cSqlStmt = "SELECT a.taxid, a.seq_no, a.ded_pct, a.ded_amt, a.ded_amt2, b.taxcode, ifnull(b.taxname,'') as taxname  FROM PA8293 a left join pa8290 b on a.taxid=a.taxid order by b.taxid,a.seq_no,ded_amt "
    cSqlStmt = "select taxid,taxcode,taxname from PA8290 order by taxid "
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            nCtr = 0
            OpenQueryDNS "select seq_no,ded_pct,ded_amt,ded_amt2 from PA8293 WHERE TAXID=" & cQuote & oRecordSet("TAXID") & cQuote & " ORDER BY SEQ_NO", oRSet, False
            If oRSet.RecordCount > 0 Then
                While Not oRSet.EOF
                    aOtherInfo(nCtr) = oRSet("ded_pct")
                    aOtherInfo1(nCtr) = oRSet("ded_amt")
                    aOtherinfo2(nCtr) = oRSet("ded_amt2")
                                       
                    nCtr = nCtr + 1
                    oRSet.MoveNext
                    
                Wend
            Else
                aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#)
                aOtherInfo1 = Array(0#, 0#, 0#, 0#, 0#, 0#)
                aOtherinfo2 = Array(0#, 0#, 0#, 0#, 0#, 0#)
            End If
            
            OpenQueryDNS "select * from di2660 where cmpid =" & gCompanyID, objdbRs, False
            cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
                       
            cSqlStmt = " INSERT INTO tmpWTax(TAXID,TAXCODE,TAXNAME,DEDPCT0,DEDPCT1,DEDPCT2,DEDPCT3,DEDPCT4,DEDPCT5,DEDPCT6,DEDAMT_0, " & _
                       " DEDAMT_1,DEDAMT_2,DEDAMT_3,DEDAMT_4,DEDAMT_5,DEDAMT_6,DEDAMT1_0,DEDAMT1_1,DEDAMT1_2,DEDAMT1_3,DEDAMT1_4,DEDAMT1_5,DEDAMT1_6,cmpname)VALUES(" & _
                       cQuote & oRecordSet("TAXID") & cQuote & "," & cQuote & oRecordSet("TAXCODE") & cQuote & "," & _
                       cQuote & oRecordSet("TAXNAME") & cQuote & "," & _
                       aOtherInfo(0) & "," & aOtherInfo(1) & "," & aOtherInfo(2) & "," & aOtherInfo(3) & "," & aOtherInfo(4) & "," & aOtherInfo(5) & "," & aOtherInfo(6) & "," & _
                       aOtherInfo1(0) & "," & aOtherInfo1(1) & "," & aOtherInfo1(2) & "," & aOtherInfo1(3) & "," & aOtherInfo1(4) & "," & aOtherInfo1(5) & "," & aOtherInfo1(6) & "," & _
                       aOtherinfo2(0) & "," & aOtherinfo2(1) & "," & aOtherinfo2(2) & "," & aOtherinfo2(3) & "," & aOtherinfo2(4) & "," & aOtherinfo2(5) & "," & aOtherinfo2(6) & "," & _
                       cQuote & cCmpname & cQuote & ")"
            QueryTemp cSqlStmt, objdbRs, True
                       
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport "Witholding Tax Report Listing", "LST8290.rpt", , True
        
        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
EndWTaxRpt:
    Set oRecordSet = Nothing
    Exit Sub
    
ErrWTaxRpt:
    ErrorMsg Err.Number, Err.Description, "Witholding Tax Listing", Name
    Resume EndWTaxRpt
End Sub

Private Sub mnuSal_Increase_Click()
    GetUserRights PadStr(mnuSal_Increase.Name, " ", 100, PadRight), gUserID
    frmSalIncrease.Show
End Sub

Private Sub mnuSalDiv_Click()
    With frmSel
        .SetSelection 18
        .Caption = "Salary Division Report"
        .Show
    End With
End Sub

Private Sub mnuSerCon_Click()
    GetUserRights PadStr(mnuSerCon.Name, " ", 100, PadRight), gUserID
    frmConfig.Show
End Sub

Private Sub mnuShift_Click()
    GetUserRights PadStr(mnuShift.Name, " ", 100, PadRight), gUserID
    frmShift.Show
End Sub

Private Sub mnuSSCheck_Click()
    GetUserRights PadStr(mnuSSCheck.Name, " ", 100, PadRight), gUserID
    frmCheckSched.Show
End Sub

Private Sub mnuSSS_Click()
    GetUserRights PadStr(mnuSSS.Name, " ", 100, PadRight), gUserID
    frmsss.Show
End Sub

Private Sub mnuSSSR3_Click()
    GetUserRights PadStr(mnuSSSR3.Name, " ", 100, PadRight), gUserID
    frmR3sss.Show
End Sub

Private Sub mnuSwap_Click()
    GetUserRights PadStr(mnuSwap.Name, " ", 100, PadRight), gUserID
    frmSwap.Show
End Sub

Private Sub mnuSysMenu_Click()
    frmSysMenu.Show
End Sub

Private Sub mnuTax_Click()
    GetUserRights PadStr(mnuTax.Name, " ", 100, PadRight), gUserID
    frmTax.Show
End Sub

Private Sub mnuTax2_Click()
    GetUserRights PadStr(mnuTax2.Name, " ", 100, PadRight), gUserID
    frmTOI.Show
End Sub

Private Sub mnuTax2Old_Click()
    GetUserRights PadStr(mnuTax2Old.Name, " ", 100, PadRight), gUserID
    frmTOI_OLD.Show
End Sub

Private Sub mnuTMS_Click()
    GetUserRights PadStr(mnuTMS.Name, " ", 100, PadRight), gUserID
'    frmTMS.Show
    frmTMS2.Show
End Sub

Private Sub mnuTMSDailyRep_Click()
    With frmReport
        .GetRpt 1
        .Show
    End With
End Sub

Private Sub mnuTMSLA_Click()
    With frmReport
        .GetRpt 2
        .Show
    End With
End Sub

Private Sub mnuTMSRpt_Click()
    With frmSel
        .SetSelection 47
        .Caption = "TMS Report"
        .Show
    End With
End Sub

Private Sub mnuTransaction_Click()
    frmSel.SetSelection 1
    frmSel.Caption = "Transaction Entry"
    frmSel.Show
End Sub

Private Sub mnuUpAlpha_Click()
    frmAlphalist.Show
End Sub

Private Sub mnuUploadOld_Click()
    frmUpload.Show
End Sub

Sub RebuildHoliday()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        nDedAmt, _
        nDedAmt2, _
        nDedAmt3, _
        nTotDed, _
        n13mopay, _
        nSAAmt, _
        nGrossAmt As Double, _
        aDedAmt As Variant, _
        oRecordSet As New ADODB.Recordset, _
        lFound As Boolean
        
    Dim oRset1 As New ADODB.Recordset
    
    aDedAmt = Array(0#, 0#, 0#, 0#, 0#, 0#, "")     ' --> for deduction purposes
    ' (0)   -   SSS ER
    ' (1)   -   SSS Premium
    ' (2)   -   Withholding Tax
    ' (3)   -   PhilHealth/Medicare PS
    ' (4)   -   PhilHealth/Medicare ES
    ' (5)   -   Medicare Total
    
    cSqlStmt = "select a.PERIODID, a.PERIOD_STAT, " & _
               "  a.EMPID, a.DEPID, a.TAXID, a.ACTIVE, a.EMP_STAT, a.date_res, a.DATE_HIRE, " & _
               "  a.RATE_AMT, a.COLA_AMT, a.POS_ALLOW, a.WAP, " & _
               "  a.REG_PAY, a.REG_OT_PAY, " & _
               "  a.NDIFF_PAY, a.NDIFF_OT_PAY, " & _
               "  a.HOLIDAY, a.HOL_PAY, a.ADJ_PAY, a.OTHER_PAY, a.LEAVE_PAY, a.M13PAY, " & _
               "  a.GROSS16231, a.SSER, a.SSPREM, a.SSS01, a.EC001, a.SSER1215, a.SSPREM1215, a.MEDICARE, a.MEDICARE2, a.MED01, a.PS1215, a.ES1215, " & _
               "  a.DED_AMT, a.GROSS_PAY, a.NET_PAY, " & _
               "  a.PAYSTATUS, a.BASIC1215, " & _
               "  a.BASICPAY , a.COLA, a.TAX1215, a.SUN_COLA, a.sa_ndiff_pay, a.sa_reg_pay, a.sa_adj_pay, a.sun_pay, a.sun_ot_pay," & _
               "  ifnull(b.ytd_gross,0) as ytd_gross, ifnull(b.ytd_gross_sa,0) as ytd_gross_sa, ifnull(b.ytd_basic,0) as ytd_basic, ifnull(b.ytd_cola,0) as ytd_cola " & _
               "from pa87260 a left join di3670 b on a.empid=b.empid " & _
               "where (a.ndiff_day>0)"
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oTempADO.EOF
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
            '    nGrossAmt = RegPay +
            '                RegOTPay +
            '                NDiffPay +
            '                HolPay +
            '                PosAllow +
            '                COLA +
            '                Incentive Leave +
            '                Adjustment             --> ala p 2 computation d2...
            
'            MsgBox DateDiff("d", Format(oTempADO("date_hire"), "yyyy-mm-dd"), "2006-11-03") & vbCrLf & _
                   oTempADO("date_hire")
            
            nGrossAmt = Round(oTempADO("reg_pay") + _
                        oTempADO("reg_ot_pay") + _
                        oTempADO("ndiff_pay") + _
                        oTempADO("ndiff_ot_pay") + _
                        oTempADO("hol_pay") + _
                        oTempADO("pos_allow") + _
                        oTempADO("cola") + _
                        oTempADO("leave_pay") + _
                        oTempADO("adj_pay"), 2)
                        
            '    nNetAmt = SARegOTPay +
            '              SANDiffOTPay +
            '              SunCola +
            '              SunPay +
            '              SunOTPay +
            '              SAAdjPay         --> ala p 2 computation d2...
            nSAAmt = Round(oTempADO("sa_reg_pay") + _
                    oTempADO("sa_ndiff_pay") + _
                    oTempADO("sun_cola") + _
                    oTempADO("sun_pay") + _
                    oTempADO("sun_ot_pay") + _
                    oTempADO("sa_adj_pay"), 2)
                        
            If oTempADO("active") > 0 Then
                If oTempADO("emp_stat") = 1 Then
                    n13mopay = Round((oTempADO("YTD_BASIC") + oTempADO("basicpay")) / 12, 2)
                ElseIf oTempADO("emp_stat") = 2 Then
                    n13mopay = Round((oTempADO("YTD_GROSS") + oTempADO("YTD_GROSS_SA") - oTempADO("YTD_COLA") + nGrossAmt + oTempADO("sa_net_pay") - oTempADO("cola")) / 12, 2)
                Else
                    n13mopay = 0
                End If
            End If
            
            nTotDed = 0
            
            cSqlStmt = "select dedid, ded_amt, ded_amt2, ded_amt3 from pa87263 where empid=" & cQuote & oTempADO("empid") & cQuote & " order by dedid"
            OpenQueryDNS cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
            
                aDedAmt = Array(0#, 0#, 0#, 0#, 0#, 0#, "")
                
                While Not oRecordSet.EOF
                
                    nDedAmt = 0
                    nDedAmt2 = 0
                    nDedAmt3 = 0
                    lFound = False
                
                    If nGrossAmt > 0 Then
                        Select Case oRecordSet("dedid")
                            Case "001"      ' --> sss premium
                                cSqlStmt = "select ER_SS, EE_SS, ER_EC  from pa7770 where " & (oTempADO("gross16231") + nGrossAmt) & " between range1 and range2"
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                nDedAmt = IIf(objdbRs.RecordCount > 0, objdbRs("EE_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSPREM1215"), 0)
                                nDedAmt2 = IIf(objdbRs.RecordCount > 0, objdbRs("ER_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSER1215"), 0)
                                nDedAmt3 = IIf(objdbRs.RecordCount > 0, objdbRs("ER_EC"), 0)
                                aDedAmt(0) = nDedAmt
                                aDedAmt(1) = nDedAmt2
                                aDedAmt(5) = IIf(objdbRs.RecordCount > 0, objdbRs("ER_EC"), 0)
                                lFound = True
                                
                            Case "005"      ' --> PhilHealth/Medicare
                                If aDedAmt(0) > 0 Then
                                    cSqlStmt = "select def_amt from di3673 " & _
                                               " where (empid=" & cQuote & oTempADO("EMPID") & cQuote & ")" & _
                                               " and (dedid=" & cQuote & oRecordSet("DEDID") & cQuote & ")" & _
                                               " and (" & IIf(oTempADO("period_stat") = 1, "period1=1", "period2=1") & ")"
                                    OpenQueryDNS cSqlStmt, oRset1, False
                                    
                                    cSqlStmt = "select PS, ES from PA7454 where " & (oTempADO("gross16231") + nGrossAmt) & " between range1 and range2"
                                    OpenQueryDNS cSqlStmt, objdbRs, False
                                    If oRset1.RecordCount > 0 Then
                                        nDedAmt = oRset1("def_amt") - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0)
                                    Else
                                        nDedAmt = IIf(objdbRs.RecordCount > 0, objdbRs("PS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0)
                                    End If
                                    nDedAmt2 = IIf(objdbRs.RecordCount > 0, objdbRs("ES"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("ES1215"), 0)
                                    aDedAmt(3) = nDedAmt
                                    aDedAmt(4) = nDedAmt2
                                Else
                                    aDedAmt(3) = 0
                                    aDedAmt(4) = 0
                                End If
                                lFound = True
                                
                            Case Else
                                nDedAmt = oRecordSet("ded_amt")
                                
                        End Select
                    End If
                    
                    nTotDed = nTotDed + nDedAmt
                    
                    If lFound Then
'                        If oTempADO("empid") = "297319" Then MsgBox "test"
                        If IfExists("pa87263", _
                                    "(periodid=" & cQuote & oTempADO("periodid") & cQuote & ") and (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (dedid=" & cQuote & oRecordSet("dedid") & cQuote & ")") Then
                            cSqlStmt = "update pa87263 set ded_amt = " & Round(nDedAmt, 2) & "," & _
                                       " ded_amt2 = " & Round(nDedAmt2, 2) & "," & _
                                       " ded_amt3 = " & Round(nDedAmt3, 2) & _
                                       " where (periodid=" & cQuote & oTempADO("periodid") & cQuote & ") and (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (dedid=" & cQuote & oRecordSet("dedid") & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                    End If
                    
                    oRecordSet.MoveNext
                    
                Wend
                
            End If
            
            If nTotDed > 0 Then
                cSqlStmt = "UPDATE PA87260 SET " & _
                           " DED_AMT=" & Round(nTotDed, 2) & "," & _
                           " SSPREM=" & Round(aDedAmt(0), 2) & "," & _
                           " SSER=" & Round(aDedAmt(1), 2) & "," & _
                           " SSS01=" & Round(aDedAmt(0) + aDedAmt(1), 2) & "," & _
                           " EC001=" & aDedAmt(5) & "," & _
                           " MEDICARE=" & Round(aDedAmt(3), 2) & "," & _
                           " MEDICARE2=" & Round(aDedAmt(4), 2) & "," & _
                           " MED01=" & Round(aDedAmt(3) + aDedAmt(4), 2) & "," & _
                           " m13pay=" & Round(n13mopay, 2) & "," & _
                           " gross_pay=" & Round(nGrossAmt + n13mopay, 2) & "," & _
                           " sa_net_pay=" & Round(nSAAmt, 2) & "," & _
                           " NET_PAY=" & Round(nGrossAmt + n13mopay - nTotDed, 2) & _
                           " WHERE PERIODID=" & cQuote & oTempADO("periodid") & cQuote & _
                           " AND EMPID=" & cQuote & oTempADO("EMPID") & cQuote
            Else
                cSqlStmt = "UPDATE PA87260 SET DED_AMT=0," & _
                           " SSPREM=0," & _
                           " SSER=0," & _
                           " SSS01=0," & _
                           " EC001=0," & _
                           " MEDICARE=0," & _
                           " MEDICARE2=0," & _
                           " MED01=0," & _
                           " sa_net_pay=" & Round(nSAAmt, 2) & "," & _
                           " m13pay=" & Round(n13mopay, 2) & "," & _
                           " gross_pay=" & Round(nGrossAmt + n13mopay, 2) & "," & _
                           " NET_PAY=" & Round(nGrossAmt + n13mopay - nTotDed, 2) & _
                           " WHERE PERIODID=" & cQuote & oTempADO("periodid") & cQuote & _
                           " AND EMPID=" & cQuote & oTempADO("EMPID") & cQuote
            End If
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            oTempADO.MoveNext
        Wend
        
        ShowProgress 4
        
        MsgBox "Done!!!"
        
    End If
    
    Set oRecordSet = Nothing
    
    Set oRset1 = Nothing
End Sub

Private Sub mnuWAP_Click()
    Dim cDepid, _
        cSqlStmt As String, _
        nCtr As Integer, _
        nDedAmt As Double, _
        nAmount As Double
    
    
    If MsgBox("Rebuild Holiday pay?", vbYesNo, "System Advisory!!!") = vbYes Then
        RebuildHoliday
    End If
    
'    If MsgBox("Rebuild Sequence Number for WAP?", vbYesNo, "System Advisory!!!") = vbYes Then
'        cSqlStmt = "select a.empid, a.depid, if(a.active>0,1,0) as active2 " & _
'                   "from pa87260 a " & _
'                   "Where a.emp_stat = 0 " & _
'                   "order by active2, if(a.active>0,'',a.depid), a.fullname "
'        OpenQueryDNS cSqlStmt, oTempADO, False
'        If oTempADO.RecordCount > 0 Then
'
'            ShowProgress 0
'
'            While Not oTempADO.EOF
'
'                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
'
'                If (oTempADO("active2") = 0) Then
'                    If (oTempADO("depid") <> cDepid) Then
'                        cDepid = oTempADO("depid")
'                        nCtr = 0
'                    End If
'                Else
'                    If (cDepid <> "999") Then
'                        cDepid = "999"
'                        nCtr = 0
'                    End If
'                End If
'                nCtr = nCtr + 1
'
'                cSqlStmt = "update pa87260 set seq_no=" & nCtr & _
'                           " where empid=" & cQuote & oTempADO("empid") & cQuote
'                OpenQueryDNS cSqlStmt, objdbRs, True
'                Script2File cSqlStmt
'
'                oTempADO.MoveNext
'
'            Wend
'
'            ShowProgress 4
'
'            MsgBox "Done"
'        End If
'    End If

'    If MsgBox("Rebuild COLA?", vbYesNo, "System Advisory!!!") = vbYes Then
'        ShowProgress 0
'        cSqlStmt = "select empid, cola_amt, reg_day, ndiff_day from pa87260"
'        OpenQueryDNS cSqlStmt, oTempADO, False
'        If oTempADO.RecordCount > 0 Then
'            While Not oTempADO.EOF
'                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
'                cSqlStmt = "update pa87260 set cola=" & Round(oTempADO("cola_amt") * (oTempADO("reg_day") + oTempADO("ndiff_day")), 2) & _
'                           " where empid=" & cQuote & oTempADO("empid") & cQuote
'                OpenQueryDNS cSqlStmt, objdbRs, True
'                Script2File cSqlStmt
'                oTempADO.MoveNext
'            Wend
'        End If
'        ShowProgress 4
'        MsgBox "Done"
'    End If

    
' --> taxable amount for december re-computation...
'    If MsgBox("Rebuild Taxable amount?", vbYesNo, "System Advisory!!!") = vbYes Then
'        ShowProgress 0
'
'        cDepid = "pah"
'
'loopd2:
'
'        cSqlStmt = "select a.periodid, a.empid, a.taxable, a.gross_pay, sum(b.ded_amt) as ded_amt " & _
'                   "from " & cDepid & "87260 a left join " & cDepid & "87263 b on a.periodid=b.periodid and a.empid=b.empid " & _
'                   "where b.dedid in ('001','003','005','007') " & _
'                   "group by b.periodid, b.empid"
'        OpenQueryDNS cSqlStmt, oTempADO, False
'        If oTempADO.RecordCount > 0 Then
'
'            While Not oTempADO.EOF
'
'                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
'
'                cSqlStmt = "update " & cDepid & "87260 set taxable=" & oTempADO("gross_pay") - oTempADO("ded_amt") & _
'                           " where periodid=" & cQuote & oTempADO("periodid") & cQuote & _
'                           " and empid=" & cQuote & oTempADO("empid") & cQuote
'                OpenQueryDNS cSqlStmt, objdbRs, True
'                Script2File cSqlStmt
'
'                oTempADO.MoveNext
'
'            Wend
'
'        End If
'
'        ShowProgress 4
'
'        If cDepid = "pah" Then
'            cDepid = "pa"
'            GoTo loopd2
'        End If
'
'    End If


'    If MsgBox("Rebuild Withholding Tax?", vbYesNo, "System Advisory!!!") = vbYes Then
'
'        ShowProgress 0
'
'        cSqlStmt = "select a.empid, a.emp_stat, a.taxid, a.taxable, a.tax1215, a.wtax, a.gross_pay, a.ded_amt as deduction, sum(b.ded_amt) as ded_amt " & _
'                   "from pa87260 a left join pa87263 b on a.periodid=b.periodid and a.empid=b.empid " & _
'                   "where b.dedid in ('001','003','005','007') " & _
'                   "group by b.empid"
'        OpenQueryDNS cSqlStmt, oTempADO, False
'        If oTempADO.RecordCount > 0 Then
'
'            While Not oTempADO.EOF
'
'                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
'
'                If oTempADO("emp_stat") = 2 Then
'
'                    cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
'                               " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("tax1215") + (oTempADO("gross_pay") - oTempADO("ded_amt"))) & ">=ded_amt2)" & _
'                               " order by ded_amt2 desc limit 1"
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'                        If objdbRs("DED_PCT") > 0 Then
'                            nDedAmt = objdbRs("DED_AMT") + (((oTempADO("tax1215") + (oTempADO("gross_pay") - oTempADO("ded_amt"))) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
'                        Else
'                            nDedAmt = 0
'                        End If
'                    End If
'
'                    ' --> update correct withholding tax...
'                    cSqlStmt = "update pa87263 set ded_amt = " & Round(nDedAmt, 2) & _
'                               " where (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (dedid='006')"
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'                    Script2File cSqlStmt
'
'                    cSqlStmt = "select sum(ded_amt) as tot_ded from pa87263 " & _
'                               "where empid=" & cQuote & oTempADO("empid") & cQuote & _
'                               " group by empid"
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'                        nAmount = objdbRs("tot_ded")
'                    Else
'                        nAmount = 0
'                    End If
'
'                    ' --> update other necessary field in header...
'                    cSqlStmt = "update pa87260 set taxable=" & oTempADO("gross_pay") - oTempADO("ded_amt") & "," & _
'                               "                   ded_amt = " & nAmount & "," & _
'                               "                   net_pay = gross_pay - " & oTempADO("deduction") - oTempADO("wtax") + Round(nDedAmt, 2) & "," & _
'                               "                   wtax = " & Round(nDedAmt, 2) & _
'                               " where empid=" & cQuote & oTempADO("empid") & cQuote
''                    MsgBox cSqlStmt
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'                    Script2File cSqlStmt
'
'                End If
'
'                oTempADO.MoveNext
'
'            Wend
'
'        End If
'
'        ShowProgress 4
'
'    End If
End Sub

Private Sub mnuWklyCon_Click()
 With frmReport
        .GetRpt 10
        .Show
    End With
End Sub

Private Sub mnuWLCostRpt_Click()
    With frmReport
        .GetRpt 6
        .Show
    End With
End Sub

Private Sub mnuXPSideMenu_Click()
    XPSideMenu1.Visible = Not XPSideMenu1.Visible
    mnuXPSideMenu.Caption = IIf(XPSideMenu1.Visible, "Hide", "Display") & " Side Menu"
End Sub

Private Sub XPSideMenu1_Action(Frame As Integer, Button As Integer)
    Dim nCtr, nCount As Integer
    Select Case Frame
        Case 0
            For nCtr = 0 To Forms.Count - 1
                nCount = nCount + IIf(Forms(nCtr).Name <> "frmSplash", 1, 0)
                If Forms(nCtr).Visible Then Forms(nCtr).Show
            Next nCtr
            
            If nCount > 1 Then
                MsgBox "Please close all the other module to terminate...", vbOKOnly, App.Title
            Else
                Select Case Button
                    Case 0      ' --> logged-off
                        nTimeCnt = 0
                        OpenQueryDNS "UPDATE PA2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote, objdbRs, True
                        Script2File "UPDATE PA2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote
                        Log2Audit "frmMain", "User logged-off."
                        Write2File gUserID & "logged-off at " & Now
                        Write2File ""
                        Hide
                        frmSplash.Show
                        showLogin
                    Case 1      ' --> exit app
                        OpenQueryDNS "UPDATE PA2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote, objdbRs, True
                        Script2File "UPDATE PA2360 SET STATUS=0 WHERE USERID=" & cQuote & gUserID & cQuote
                        Log2Audit "frmMain", "User end application."
                        Write2File gUserID & "end application at " & Now
                        Write2File ""
                        MDIForm_Unload 0
                        
                End Select
            End If
            
        Case 1
            Select Case Button
                Case 0
                    mnuDept_Click
                Case 1
                    mnuEmployee_Click
            End Select
            
        Case 2
            Select Case Button
                Case 0      ' --> shifting schedule
                    mnuDepShiftSched_Click
                Case 1      ' --> time Management
                    mnuTMS_Click
                Case 2      ' --> upload from bio-clock
                    mnuUpload_Click
            End Select
            
    End Select
End Sub

Private Sub mnuUpload_Click()
    GetUserRights PadStr(mnuUpload.Name, " ", 100, PadRight), gUserID
    frmBioProcess.Show
End Sub

'Sub CreateTmpPaySlip(ByVal nMode As Integer)
'    On Error GoTo ErrCreate
'    Dim cSqlStmt As String, _
'        cParam As String
'
'    If nMode = 0 Then
'        cSqlStmt = " CREATE TABLE tmp7297655(   [PERIODNAME] CHAR(100),     [SEQ_NO] INTEGER, " & _
'                   " [P_DAY] DOUBLE,            [P_HOLIDAY] DOUBLE, " & _
'                   " [DEPID] CHAR(3),           [DEPTNAME] CHAR(100),       [DEPID2] CHAR(3),           [DEPTNAME2] CHAR(100), " & _
'                   " [EMPID] CHAR(6),           [ACTIVE] INTEGER,           [EMP_STAT] INTEGER," & _
'                   " [FULLNAME] CHAR(100),      [FNAME] CHAR(50),           [LNAME] CHAR(50),           [MNAME] CHAR(50), " & _
'                   " [RATE_AMT] DOUBLE,         [COLA_AMT] DOUBLE,          [SUN_COLA] DOUBLE,          [POS_ALLOW] DOUBLE, " & _
'                   " [REG_DAY] DOUBLE,          [REG_PAY] DOUBLE,           [REG_OT_HR] DOUBLE,         [REG_OT_PAY] DOUBLE, " & _
" [NDIFF_DAY] DOUBLE,        [NDIFF_PAY] DOUBLE,         [NDIFF_OT_HR] DOUBLE,       [NDIFF_OT_PAY] DOUBLE, " & _
'                   " [HOLIDAY] DOUBLE,          [HOL_PAY] DOUBLE, " & _
'                   " [SA_REG_OT] DOUBLE,        [SA_REG_PAY] DOUBLE,        [SA_NDIFF_OT] DOUBLE,       [SA_NDIFF_PAY] DOUBLE, " & _
'                   " [SUN_HR] DOUBLE,           [SUN_PAY] DOUBLE,           [SUN_OT] DOUBLE,            [SUN_OT_PAY] DOUBLE, " & _
'                   " [SUN_ND] DOUBLE,           [SUN_ND_PAY] DOUBLE,        [SUN_ND_OT] DOUBLE,         [SUN_ND_OT_PAY] DOUBLE, " & _
'                   " [ADJ_PAY] DOUBLE,          [SA_ADJ_PAY] DOUBLE, " & _
'                   " [OTHER_PAY] DOUBLE,        [LEAVE_PAY] DOUBLE, " & _
'                   " [DED_AMT] DOUBLE,          [GROSS_PAY] DOUBLE, " & _
'                   " [NET_PAY] DOUBLE,          [SA_NET_PAY] DOUBLE, " & _
'                   " [SIGNATORY1] char(50),     [POSNAME1] char(50)," & _
'                   " [SIGNATORY2] char(50),     [POSNAME2] char(50)," & _
'                   " [SIGNATORY3] char(50),     [POSNAME3] char(50)," & _
'                   " [SIGNATORY4] char(50),     [POSNAME4] char(50)," & _
'                   " [SIGNATORY5] char(50),     [POSNAME5] char(50)," & _
'                   " [SIGNATORY6] char(50),     [POSNAME6] char(50)," & _
'                   " [SIGNATORY7] char(50),     [POSNAME7] char(50)," & _
'                   " [13MO_PAY] DOUBLE,         [CMPID] char(4), [REG_OT_RATE] DOUBLE )"
'    Else
'        cSqlStmt = " CREATE TABLE tmp7297655d(" & _
'                   " [PERIODID] CHAR(5),        [EMPID] CHAR(6)," & _
'                   " [DEDID] CHAR(3),           [DEDNAME] CHAR(100)," & _
'                   " [AMOUNT] DOUBLE,           [SEQ_NO] INTEGER)"
'
'    End If
'    'MsgBox cSqlStmt
'
'    oTempConn.Execute cSqlStmt
'    While oTempConn.State = adStateExecuting
'        DoEvents
'    Wend
'ErrCreate:
'    ' in case table is already existing, let's clear it...
'    QueryTemp "DELETE FROM " & IIf(nMode = 0, "tmp7297655", "tmp7297655d"), oTempADO, True
'End Sub


