VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{30DA1A2F-A970-4238-AC17-5773BA9DC841}#1.1#0"; "CIAXPDatePicker.ocx"
Begin VB.Form frmR3sss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS R3 Migration Tool"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1740
      TabIndex        =   25
      Tag             =   "1"
      ToolTipText     =   "TXT:REC_BY"
      Top             =   2970
      Width           =   660
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   315
      Left            =   2430
      TabIndex        =   24
      Top             =   2955
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   5715
      TabIndex        =   22
      Top             =   1965
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   21
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   1965
      Width           =   3945
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   1365
      Width           =   1290
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmR3sss.frx":0000
      Left            =   1725
      List            =   "frmR3sss.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2595
      Width           =   1470
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmR3sss.frx":0042
      Left            =   1725
      List            =   "frmR3sss.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2265
      Width           =   1920
   End
   Begin VB.Frame Frame2 
      Height          =   885
      Left            =   15
      TabIndex        =   12
      Top             =   3315
      Width           =   6255
      Begin VB.CommandButton Command7 
         Caption         =   "&Gen. List"
         Height          =   660
         Left            =   3420
         Picture         =   "frmR3sss.frx":0084
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   150
         Width           =   840
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   2460
         Picture         =   "frmR3sss.frx":1A06
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Close"
         Height          =   660
         Left            =   5235
         Picture         =   "frmR3sss.frx":3388
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Gen. R3"
         Height          =   660
         Left            =   4275
         Picture         =   "frmR3sss.frx":4D0A
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   1665
      Width           =   3945
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "TXT:WORKCENTERID"
      Top             =   1065
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   5715
      TabIndex        =   5
      Top             =   1635
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:LINEID"
      Top             =   135
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1725
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:LINENAME"
      Top             =   435
      Width           =   4500
   End
   Begin ciaXPDatePicker.XPDatePicker XPDatePicker1 
      Height          =   315
      Left            =   1725
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_SCHED"
      Top             =   735
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   556
      FormatString    =   "long date"
      MouseIcon       =   "frmR3sss.frx":668C
      CalendarDayBorder=   -1  'True
      CalendarDayBorderColor=   -2147483646
      CalendarMonthBorderColor=   8421504
      LicValid        =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Certified By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   27
      Top             =   2970
      Width           =   1530
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Prepared By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2880
      TabIndex        =   26
      Top             =   3000
      Width           =   3930
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Exell Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   23
      Top             =   1995
      Width           =   1560
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   20
      Top             =   1380
      Width           =   1530
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   16
      Top             =   2640
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer SSS ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   465
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   3615
      Left            =   0
      Top             =   -270
      Width           =   1725
   End
End
Attribute VB_Name = "frmR3sss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmR3sss
' description   :   Module for SSS R3 Migration Tool
' programmer    :   _-=[ srm ]=-_
' date          :   01 Jul 2015

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Sub Command1_Click()
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.InitDir = CheckPath(cUploadPath)
    CommonDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx"
    CommonDialog1.ShowOpen
    
    Text6.Text = CommonDialog1.FileName
End Sub

Private Sub Command10_Click()

    SSSR3_Migration True, False

'    On Error GoTo ErrLoad
'
'    Dim ExcelObj As Object
'    Dim ExcelBook As Object
'    Dim ExcelSheet As Object
'
'    Dim nRowlast As Long, _
'        nRowCnt As Long
'
'    Dim aSSSInfo, aFirstRowInfo, aSizeInfor, aUserInfo As Variant
'
'    Dim oTextFile As New FileSystemObject, _
'        oFile As File, _
'        oTxtStream As TextStream, _
'        cLogFirstRow, cLogDatFileName, cDatFileName, cLogSSSDetail, cLogLastData As String
'
'    Dim cSqlStmt As String, _
'        nCombo As Integer, _
'        oRecordSet As New ADODB.Recordset, _
'        oRSet As New ADODB.Recordset, _
'        cParam2 As String, _
'        nCtr As Integer
'
'    Dim nSSSTotAmt, nSSSTotEC, nSSSTolEmp As Long
'
'    aSSSInfo = Array("", "", "", "", 0#, 0#)
''    aSSSInfo(0) = "SSS Number"
''    aSSSInfo(1) = "Last Name"
''    aSSSInfo(2) = "First Name"
''    aSSSInfo(3) = "MI"
''    aSSSInfo(4) = "SSS Amount"
''    aSSSInfo(5) = "SSS EE"
'
'    aFirstRowInfo = Array("", "", "", "", "", "", "", "", "", "")
''    aFirstRowInfo(0) = "Company Name"
''    aFirstRowInfo(1) = "For the month format mmyyyy"
''    aFirstRowInfo(2) = "Employer SSS Number"
''    aFirstRowInfo(3) = "Reciept + GE"
''    aFirstRowInfo(4) = "Date Paid format mmddyyyy"
''    aFirstRowInfo(5) = "Total Amoount XXXXXXXXX.XX"
''    aFirstRowInfo(6) = "current date format mmdd"
''    aFirstRowInfo(7) = "current time format hhmm"
''    aFirstRowInfo(8) = "Constant vaue XXXX0.00"
''    aFirstRowInfo(9) = "NO"
'
'    aSizeInfor = Array("", "", "", "", "", "", "", "", "")
''    aSizeInfor(0) = "6"
''    aSizeInfor(1) = "7"
''    aSizeInfor(2) = "8"
''    aSizeInfor(3) = "9"
''    aSizeInfor(4) = "10"
''    aSizeInfor(5) = "11"
''    aSizeInfor(6) = "12"
''    aSizeInfor(7) = "13"
''    aSizeInfor(8) = "14"
'
'    nRowlast = 5000
'    nSSSTotAmt = 0
'    nSSSTotEC = 0
'    nSSSTolEmp = 0
'
'    'Check info if meron value
'    If Text3.Text <> "" And Text4.Text <> "" And Text6.Text <> "" Then
'
'        CreateTMPSSSREM
'
'        CreateTMPSSSDAT
'
'        aUserInfo = Array("", "", "", "", "")
'
'        If Not ChkPersonnel(Text7) Then Exit Sub
'        OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            objdbRs.Requery adAsyncFetch
'            objdbRs.Find "USERID='" & Text7.Text & "'"
'            aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
'        End If
'
'        ShowProgress 0
'
'        Set ExcelObj = CreateObject("Excel.Application")
'        Set ExcelSheet = CreateObject("Excel.Sheet")
'
'        ExcelObj.Workbooks.Open Text6.Text
'
'        Set ExcelBook = ExcelObj.Workbooks(1)
'        Set ExcelSheet = ExcelBook.Worksheets(1)
'
'        aFirstRowInfo(0) = UCase(cCompany)
'        aFirstRowInfo(1) = Combo1.ListIndex + 1
'        aFirstRowInfo(1) = IIf(Len(aFirstRowInfo(1)) > 1, aFirstRowInfo(1), "0" & aFirstRowInfo(1)) & Combo2.Text
'        aFirstRowInfo(2) = Replace(gSSSNum, "-", "")
'        aFirstRowInfo(3) = IIf(Val(Text3.Text) = 0, "NOPAY", Val(Text3.Text) & "GE")
'        aFirstRowInfo(4) = Format(XPDatePicker1.CurrentDate, "MMDDYYYY")
'        aFirstRowInfo(5) = PadStr(Format(Round(Val(Text5.Text), 2), "0.00"), "0", 12)
'        aFirstRowInfo(6) = Format(Now, "MMYY")
'        aFirstRowInfo(7) = Format(Now, "HHMM")
'        aFirstRowInfo(8) = "NO"
'
'        aSizeInfor(0) = PadStr("0.00", " ", 6)
'        aSizeInfor(1) = PadStr("0.00", " ", 7)
'        aSizeInfor(2) = PadStr("0.00", " ", 8)
'        aSizeInfor(3) = PadStr("0.00", " ", 9)
'        aSizeInfor(4) = PadStr("0.00", " ", 10)
'        aSizeInfor(5) = PadStr("0.00", " ", 11)
'        aSizeInfor(6) = PadStr("0.00", " ", 12)
'        aSizeInfor(7) = PadStr("0.00", " ", 13)
'        aSizeInfor(8) = PadStr("0.00", " ", 14)
'
'        cLogDatFileName = "R3" & aFirstRowInfo(2) & aFirstRowInfo(1) & "." & aFirstRowInfo(6) & aFirstRowInfo(7)
'        cDatFileName = cLogDatFileName
'        cLogDatFileName = CheckPath(Text4.Text) & cLogDatFileName
'
'        cLogFirstRow = "00" & PadStr(aFirstRowInfo(0), " ", 30, 1) & aFirstRowInfo(1) & aFirstRowInfo(2) & PadStr(aFirstRowInfo(3), " ", 10, 1) & _
'                       aFirstRowInfo(4) & aFirstRowInfo(5)
'
'        If Dir(cLogDatFileName) = "" Then
'            Set oTxtStream = oTextFile.CreateTextFile(cLogDatFileName, True)
'        Else
'            Set oFile = oTextFile.GetFile(cLogDatFileName)
'            Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
'        End If
'
'        oTxtStream.WriteLine cLogFirstRow
'
''        read excel data
'        For nRowCnt = 3 To nRowlast Step 1
'
'            With ExcelSheet
'                If .Cells(nRowCnt, 1) <> "" Then
'
'                    ShowProgress 2, (1 / nRowlast) * 100, , , "Copying " & .Cells(nRowCnt, 1) & "... " & Trim(.Cells(nRowCnt, 2) & ", " & .Cells(nRowCnt, 3)) & "..."
'
'                    cLogSSSDetail = ""
'
'                    aSSSInfo(0) = .Cells(nRowCnt, 1)
'                    aSSSInfo(1) = left(.Cells(nRowCnt, 2), 15)
'                    aSSSInfo(2) = left(.Cells(nRowCnt, 3), 15)
'                    aSSSInfo(3) = IIf(.Cells(nRowCnt, 4) = "", " ", .Cells(nRowCnt, 4))
'                    aSSSInfo(4) = Format(Round(IIf(.Cells(nRowCnt, 5) = "", 0, .Cells(nRowCnt, 5)), 2), "0.00")
'                    aSSSInfo(5) = Format(Round(IIf(.Cells(nRowCnt, 6) = "", 0, .Cells(nRowCnt, 6)), 2), "0.00")
'
'                    nSSSTotAmt = nSSSTotAmt + Val(.Cells(nRowCnt, 5))
'                    nSSSTotEC = nSSSTotEC + Val(.Cells(nRowCnt, 6))
'                    nSSSTolEmp = nSSSTolEmp + 1
'
'                    cSqlStmt = "INSERT INTO TMPSSSDAT(LASTNAME,FIRSTNAME,MI,SSSNO,SSSAMT,SSSEE)VALUES(" & _
'                                cQuote & aSSSInfo(1) & cQuote & "," & _
'                                cQuote & aSSSInfo(2) & cQuote & "," & _
'                                cQuote & aSSSInfo(3) & cQuote & "," & _
'                                cQuote & aSSSInfo(0) & cQuote & "," & _
'                                cQuote & aSSSInfo(4) & cQuote & "," & _
'                                cQuote & aSSSInfo(5) & cQuote & ")"
'                    QueryTemp cSqlStmt, objdbRs, True
'
''                    cLogSSSDetail = "20" & _
''                                    PadStr(UCase(aSSSInfo(1)), " ", 15, 1) & PadStr(UCase(aSSSInfo(2)), " ", 15, 2) & UCase(aSSSInfo(3)) & _
''                                    aSSSInfo(0) & aSizeInfor(2) & PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2) & _
''                                    aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & _
''                                    PadStr(aSSSInfo(5), " ", 6) & _
''                                    aSizeInfor(0) & _
''                                    PadStr(aFirstRowInfo(8), " ", 8)
''
''                    oTxtStream.WriteLine cLogSSSDetail
'                Else
'
'                    ShowProgress 4
'
'                    ExcelObj.Workbooks.Close
'                    ExcelObj.Quit
'                    Set ExcelObj = Nothing
'                    Exit For
'                End If
'            End With
'
'        Next nRowCnt
'
'
'        ' data from mysql
'
'        ShowProgress 0
'
'        nCombo = IIf((Combo1.ListIndex + 1) > 11, 12, (Combo1.ListIndex + 1))
'
'        cSqlStmt = " SELECT periodid,date_start,date_end FROM pa7730 " & _
'                   " where (13month=0) and (month(date_start) = " & nCombo & " and  month(date_end) = " & nCombo & ") and " & _
'                   " (year(date_start) = " & cQuote & Combo2.Text & cQuote & " and year(date_end) = " & cQuote & Combo2.Text & cQuote & ")"
'
''        Script2File cSqlStmt
'        OpenQueryDNS cSqlStmt, oRSet, False
'        cParam2 = ""
'        If oRSet.RecordCount > 0 Then
'            While Not oRSet.EOF
'                cParam2 = IIf(cParam2 = "", cQuote & oRSet("periodid"), cParam2 & cQuote & "," & cQuote & oRSet("periodid") & cQuote)
'                oRSet.MoveNext
'            Wend
'        End If
'
'        If cParam2 = "" Then GoTo ErrLoad:
'
'        cSqlStmt = "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
'                   "  left(ifnull(b.mname,c.mname),1) as mname, " & _
'                   "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
'                   "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215  " & _
'                   "FROM pah87263 a left join di3670 b on a.empid=b.empid " & _
'                   "  left join pah87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
'                   "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
'                   " and b.paystatus <> 1 " & _
'                   "group by a.empid " & _
'                   "Union All " & _
'                   "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
'                   "  left(ifnull(b.mname,c.mname),1) as mname, " & _
'                   "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
'                   "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215 " & _
'                   "FROM pa87263 a left join di3670 b on a.empid=b.empid " & _
'                   "  left join pa87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
'                   "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
'                   " and b.paystatus <> 1 " & _
'                   "group by a.empid " & _
'                   "order by empid, periodid "
'        Script2File cSqlStmt
'
'        OpenQueryDNS cSqlStmt, oRecordSet, False
'
'        If oRecordSet.RecordCount > 0 Then
'
'            While Not oRecordSet.EOF
'
'                cSqlStmt = "select empid from TMPSSSDAT where SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
'                QueryTemp cSqlStmt, objdbRs, False
'                If Not objdbRs.RecordCount > 0 Then
'
'
'                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
'
'                    aSSSInfo = Array("", "", "", "", 0#, 0#)
'
'                    aSSSInfo(0) = oRecordSet("ssnum")
'                    aSSSInfo(1) = left(UCase(EncodeStr2(DecodeStr(oRecordSet("lastname")))), 15)
'                    aSSSInfo(2) = left(UCase(EncodeStr2(DecodeStr(oRecordSet("firstname")))), 15)
'                    aSSSInfo(3) = UCase(EncodeStr2(DecodeStr(left(IIf(oRecordSet("mname") = "", "  ", oRecordSet("mname")), 1))))
'                    aSSSInfo(4) = Format(Round(IIf(oRecordSet("ded_amt") = "", 0, oRecordSet("ded_amt")), 2), "0.00")
'                    aSSSInfo(5) = IIf(aSSSInfo(4) >= 1650, "30.00", IIf(aSSSInfo(4) > 0, "10.00", "0.00"))
'
'
'
'                aEmpInfo(0) = oRecordSet("empid")                                                           'ID
'                aEmpInfo(1) = left(oRecordSet("ssnum"), 10)                                                 'ssnum
'                aEmpInfo(2) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("lastname"), 15))))                'esurn
'                aEmpInfo(3) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("firstname"), 15))))               'ename
'                aEmpInfo(4) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("mname"), 1))))                    'emidinit
'                aEmpInfo(5) = oRecordSet("ded_amt")                                                         'ssamt
'                aEmpInfo(6) = oRecordSet("sser") + oRecordSet("sser1215")
'                aEmpInfo(7) = oRecordSet("ssprem") + oRecordSet("ssprem1215")
'
'
'
'                    nSSSTotAmt = nSSSTotAmt + Val(aSSSInfo(4))
'                    nSSSTotEC = nSSSTotEC + Val(aSSSInfo(5))
'                    nSSSTolEmp = nSSSTolEmp + 1
'
'
'
'
'
'
'                    cSqlStmt = "INSERT INTO TMPSSSDAT(LASTNAME,FIRSTNAME,MI,SSSNO,SSSAMT,SSSEE)VALUES(" & _
'                                cQuote & aSSSInfo(1) & cQuote & "," & _
'                                cQuote & aSSSInfo(2) & cQuote & "," & _
'                                cQuote & aSSSInfo(3) & cQuote & "," & _
'                                cQuote & aSSSInfo(0) & cQuote & "," & _
'                                cQuote & aSSSInfo(4) & cQuote & "," & _
'                                cQuote & aSSSInfo(5) & cQuote & ")"
'                    QueryTemp cSqlStmt, objdbRs, True
'
'
'
''                cLogSSSDetail = "20" & _
''                                PadStr(UCase(aSSSInfo(1)), " ", 15, 1) & PadStr(UCase(aSSSInfo(2)), " ", 15, 2) & UCase(aSSSInfo(3)) & _
''                                aSSSInfo(0) & aSizeInfor(2) & PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2) & _
''                                aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & _
''                                PadStr(aSSSInfo(5), " ", 6) & _
''                                aSizeInfor(0) & _
''                                PadStr(aFirstRowInfo(8), " ", 8)
''                oTxtStream.WriteLine cLogSSSDetail
'
'                End If
'                oRecordSet.MoveNext
'            Wend
'        End If
'
'
'        cSqlStmt = "select LASTNAME,FIRSTNAME,MI,SSSNO,SSSAMT,SSSEE from TMPSSSDAT order by lastname,firstname,mi,sssno"
'        QueryTemp cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'         While Not objdbRs.EOF
'
'                aSSSInfo(0) = objdbRs("SSSNO")
'                aSSSInfo(1) = left(UCase(EncodeStr2(DecodeStr(objdbRs("LASTNAME")))), 15)
'                aSSSInfo(2) = left(UCase(EncodeStr2(DecodeStr(objdbRs("FIRSTNAME")))), 15)
'                aSSSInfo(3) = UCase(EncodeStr2(DecodeStr(left(IIf(objdbRs("MI") = "", "  ", objdbRs("MI")), 1))))
'                aSSSInfo(4) = Format(Round(IIf(objdbRs("SSSAMT") = "", 0, objdbRs("SSSAMT")), 2), "0.00")
'                aSSSInfo(5) = Format(Round(IIf(objdbRs("SSSEE") = "", 0, objdbRs("SSSEE")), 2), "0.00")
'
'
'                cLogSSSDetail = "20" & _
'                                PadStr(UCase(aSSSInfo(1)), " ", 15, 1) & PadStr(UCase(aSSSInfo(2)), " ", 15, 2) & UCase(aSSSInfo(3)) & _
'                                aSSSInfo(0) & aSizeInfor(2) & PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2) & _
'                                aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & _
'                                PadStr(aSSSInfo(5), " ", 6) & _
'                                aSizeInfor(0) & _
'                                PadStr(aFirstRowInfo(8), " ", 8)
'                oTxtStream.WriteLine cLogSSSDetail
'
'
'            objdbRs.MoveNext
'
'         Wend
'        End If
'
'
'        'end data
'        cLogLastData = "99" & _
'                       aSizeInfor(6) & _
'                       PadStr(Format(nSSSTotAmt, "0.00"), " ", 12) & _
'                       aSizeInfor(6) & _
'                       aSizeInfor(3) & _
'                       aSizeInfor(5) & _
'                       aSizeInfor(5) & _
'                       aSizeInfor(3) & _
'                       PadStr(Format(nSSSTotEC, "0.00"), " ", 10) & _
'                       aSizeInfor(4)
'
'        oTxtStream.WriteLine cLogLastData
'
'        'Transmmital report
'        cSqlStmt = "INSERT INTO TMPSSSREM(DATFNAME,COMPANYNAME,SSSEMP,APP_PER,TR_SBR_NO,DATE_PAY,PAY_AMT,SS_AMT,EC_AMT,TOT_AMT,ECOUNT,SIGNATORY1,POSNAME1)VALUES(" & _
'                    cQuote & cDatFileName & cQuote & "," & _
'                    cQuote & UCase(aFirstRowInfo(0)) & cQuote & "," & _
'                    cQuote & gSSSNum & cQuote & "," & _
'                    cQuote & aFirstRowInfo(1) & cQuote & "," & _
'                    cQuote & aFirstRowInfo(3) & cQuote & "," & _
'                    cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & "," & _
'                    Val(aFirstRowInfo(5)) & "," & _
'                    nSSSTotAmt & "," & _
'                    nSSSTotEC & "," & _
'                    nSSSTotAmt + nSSSTotEC & "," & _
'                    nSSSTolEmp & "," & _
'                    cQuote & UCase(EncodeStr2(DecodeStr(Label11.Caption))) & cQuote & "," & _
'                    cQuote & UCase(aUserInfo(0)) & cQuote & ")"
'
''        MsgBox cSqlStmt
'        QueryTemp cSqlStmt, objdbRs, True
'
'        ShowProgress 3
'
'        GenerateReport " ", "rptSSSR3TRANS.rpt"
'
'        ShowProgress 4
'
'        oTxtStream.Close
'        Set oTxtStream = Nothing
'        Set oTextFile = Nothing
'        Set oFile = Nothing
'    Else
'        MsgBox "Please supply missing Entry!!!"
'    End If
'
'    Exit Sub
'
'ErrLoad:
'    MsgBox "Error Generating SSS..."
End Sub

Private Sub Command11_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

    'Opens a Browse Folders Dialog Box that displays the
    'directories in your computer
    Dim lpIDList As Long ' Declare Varibles
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = "File Path for SSS R3"
    ' Text to appear in the the gray area under the title bar
    ' telling you what to do
    
    With tBrowseInfo
       .hWndOwner = Me.hwnd ' Owner Form
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       Text4.Text = sBuffer
    End If
End Sub

Sub CreateTMPSSS()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE TMPSSS( " & _
               " [EMPID] char(6),       [FIRSTNAME] CHAR(20)," & _
               " [MINAME] CHAR(20),     [LASTNAME] CHAR(20)," & _
               " [SSSNO] char(10),      [SSSAMT] double," & _
               " [EC_AMT] double,       [REM] char(2)," & _
               " [DATE_REM] integer,    [COSTID] char(10)," & _
               " [CDESC] char(100),     [WORKID] char(10)," & _
               " [WDESC] char(100),     [ER_AMT] double," & _
               " [EE_AMT] double)"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPSSS"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Sub CreateTMPSSSREM()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE TMPSSSREM( " & _
               " [DATFNAME] char(30)," & _
               " [COMPANYNAME] char(100)," & _
               " [SSSEMP] char(30)," & _
               " [APP_PER] char(30)," & _
               " [TR_SBR_NO] char(30)," & _
               " [DATE_PAY] date," & _
               " [PAY_AMT] double," & _
               " [SS_AMT] double," & _
               " [EC_AMT] double," & _
               " [TOT_AMT] double," & _
               " [ECOUNT] double," & _
               " [SIGNATORY1] char(50), " & _
               " [POSNAME1] char(50))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPSSSREM"
    QueryTemp cSqlStmt, oTempADO, True
End Sub


Sub CreateTMPSSSDAT()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE TMPSSSDAT( " & _
               " [DATFNAME] char(30)," & _
               " [COMPANYNAME] char(100)," & _
               " [CURDATE] date," & _
               " [LASTNAME] char(50)," & _
               " [FIRSTNAME] char(50)," & _
               " [MI] char(1)," & _
               " [SSSNO] char(15)," & _
               " [SSSAMT] DOUBLE," & _
               " [SSSEE] DOUBLE, " & _
               " [REM] char(2)," & _
               " [DTHRD] integer," & _
               " [ECOUNT] double)"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPSSSDAT"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Private Sub Command5_Click()
    cmdClick Text7, Label11
    Text5.SetFocus
End Sub

Private Sub Command6_Click()
    SSSR3_Migration False
'    On Error GoTo ErrLoad
'    Dim cSqlStmt As String, _
'        nCombo As Integer, _
'        oFile As File, oTextFile As New FileSystemObject, _
'        oRecordSet As New ADODB.Recordset, _
'        oRSet As New ADODB.Recordset, _
'        oRSet2 As New ADODB.Recordset, _
'        cParam2 As String, cCompanyName As String, cR3file As String, _
'        nPClose As Integer, nCtr As Integer, _
'        aEmpInfo As Variant, aERPInfo As Variant
'
'    CreateTMPSSS
'
'    aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
'
'    aERPInfo = Array("", "")
'
'
'    nCombo = IIf((Combo1.ListIndex + 1) > 11, 12, (Combo1.ListIndex + 1))
'
'
'    cSqlStmt = " SELECT periodid,date_start,date_end FROM pa7730 " & _
'               " where (13month=0) and (month(date_start) = " & nCombo & " and  month(date_end) = " & nCombo & ") and " & _
'               " (year(date_start) = " & cQuote & Combo2.Text & cQuote & " and year(date_end) = " & cQuote & Combo2.Text & cQuote & ")"
'
'    'Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oRSet, False
'    cParam2 = ""
'    If oRSet.RecordCount > 0 Then
'        If oRSet.RecordCount = 1 Then
'            GoTo ErrLoad
'        End If
'
'        While Not oRSet.EOF
'            cParam2 = IIf(cParam2 = "", cQuote & oRSet("periodid"), cParam2 & cQuote & "," & cQuote & oRSet("periodid") & cQuote)
'            oRSet.MoveNext
'        Wend
'    End If
'
'    If cParam2 = "" Then GoTo ErrLoad:
'
'    cSqlStmt = "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
'               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
'               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
'               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215  " & _
'               "FROM pah87263 a left join di3670 b on a.empid=b.empid " & _
'               "  left join pah87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
'               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
'               "group by a.empid " & _
'               "Union All " & _
'               "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
'               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
'               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
'               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215 " & _
'               "FROM pa87263 a left join di3670 b on a.empid=b.empid " & _
'               "  left join pa87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
'               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
'               "group by a.empid " & _
'               "order by empid, periodid "
'    Script2File cSqlStmt
'
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'
'        ShowProgress 0
'
'        While Not oRecordSet.EOF
'
'            aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
'
'            aERPInfo = Array("", "")
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
'
'            cSqlStmt = "select empid from TMPSSS where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
'            QueryTemp cSqlStmt, objdbRs, False
'            If Not objdbRs.RecordCount > 0 Then
''                MsgBox "insert"
'
'                OpenQueryDNS "Select * FROM pa37722 where costcenterid = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'                        aERPInfo(0) = objdbRs("DESCRIPTION")
'                    Else
'                        aERPInfo(0) = ""
'                    End If
'
'
'                OpenQueryDNS "SELECT * FROM pa97722 where workcenterid = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'                            aERPInfo(1) = objdbRs("DESCRIPTION")
'                    Else
'                            aERPInfo(1) = ""
'                    End If
'
'                aEmpInfo(0) = oRecordSet("empid")                                                           'ID
'                aEmpInfo(1) = left(oRecordSet("ssnum"), 10)                                                 'ssnum
'                aEmpInfo(2) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("lastname"), 15))))                'esurn
'                aEmpInfo(3) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("firstname"), 15))))               'ename
'                aEmpInfo(4) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("mname"), 1))))                    'emidinit
'                aEmpInfo(5) = oRecordSet("ded_amt")                                                         'ssamt
'                aEmpInfo(6) = oRecordSet("sser") + oRecordSet("sser1215")
'                aEmpInfo(7) = oRecordSet("ssprem") + oRecordSet("ssprem1215")
'
'                cSqlStmt = "INSERT INTO TMPSSS(EMPID,SSSNO,LASTNAME,FIRSTNAME,MINAME,SSSAMT,EC_AMT,DATE_REM,REM,ER_AMT,EE_AMT,COSTID,CDESC,WORKID,WDESC)VALUES(" & _
'                       cQuote & oRecordSet("empid") & cQuote & "," & _
'                       cQuote & aEmpInfo(1) & cQuote & "," & _
'                       cQuote & aEmpInfo(2) & cQuote & "," & _
'                       cQuote & aEmpInfo(3) & cQuote & "," & _
'                       cQuote & aEmpInfo(4) & cQuote & "," & _
'                       aEmpInfo(5) & "," & _
'                       IIf(aEmpInfo(5) >= 1650, "30.00", IIf(aEmpInfo(5) > 0, "10.00", "0.00")) & ",0," & cQuote & "N" & cQuote & "," & _
'                       aEmpInfo(6) & "," & _
'                       aEmpInfo(7) & "," & _
'                       cQuote & oRecordSet("costcenterid") & cQuote & "," & _
'                       cQuote & aERPInfo(0) & cQuote & "," & _
'                       cQuote & oRecordSet("workcenterid") & cQuote & "," & _
'                       cQuote & aERPInfo(1) & cQuote & ")"
'
''                    MsgBox cSqlStmt
'                QueryTemp cSqlStmt, objdbRs, False
''                nCtr = nCtr + 1
'
'            Else
'
'                cSqlStmt = "update TMPSSS set " & _
'                           " SSSAMT = SSSAMT + " & oRecordSet("ded_amt") & "," & _
'                           " ER_AMT = " & oRecordSet("sser") + oRecordSet("sser1215") & "," & _
'                           " EE_AMT = " & oRecordSet("ssprem") + oRecordSet("ssprem1215") & _
'                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
'
''                MsgBox cSqlStmt
'                QueryTemp cSqlStmt, objdbRs, True
'
'                '               ---> update 201506-03
'
'                cSqlStmt = "update TMPSSS set " & _
'                           " EC_AMT = iif (SSSAMT >= 1650 , 30.00, iif(SSSAMT > 0, 10, 0) ) " & _
'                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
'
''                MsgBox cSqlStmt
'                QueryTemp cSqlStmt, objdbRs, True
'
'            End If
'
'            oRecordSet.MoveNext
'        Wend
'
'        ShowProgress 3
'
'        GenerateReport "SSS R3 Data File Report", "rptSSSR3.rpt"
'
'        ShowProgress 4
'    End If
'
'    Exit Sub
'
'ErrLoad:
'    MsgBox "Error Generating SSS..."
End Sub

Private Sub Command7_Click()
    SSSR3_Migration True, True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()

    Text1.Text = gSSSNum
    Text2.Text = cCompany
    Text3.Text = "0"
'    Text4.Text = "D:\vb codes\pay4win\sss\casssey 2015-07-02"
    Text5.Text = "0"
'    Text6.Text = "D:\vb codes\pay4win\app\upload\Monthly May 2015.xlsx"
    
    Text1.Enabled = False
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    
    With Combo1
        .Clear
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
        .ListIndex = IIf((Month(Date) - 1) = 0, 0, Month(Date) - 2)
    End With
    
    With Combo2
       .Clear
       .AddItem Year(Now) - 2
       .AddItem Year(Now) - 1
       .AddItem Year(Now)
       .AddItem Year(Now) + 1
       .AddItem Year(Now) + 2
       .AddItem Year(Now) + 3
    End With
    
    MatchCombo Year(Now), Combo2
     
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text7.Text, Label11
End Sub

Sub txtKeyDown(nMode As Integer, cString As String, oLabel As Label)
'    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command5_Click
            End Select
        Else
            ShowData cString, oLabel
        End If
'    End If
End Sub

Sub cmdClick(ByVal oTxtBox As TextBox, ByVal oLabel As Label)
    frmLookup.showPopup 1   ', " where sysuser = 1"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTxtBox.Text = cResult
        ShowData cResult, oLabel
    End If
End Sub



Sub ShowData(cString As String, oLabel As Label)
    OpenQueryDNS "SELECT USERID,CONCAT(FIRSTNAME," & cQuote & " " & cQuote & ",LASTNAME) AS FULLNAME FROM PA2360 WHERE USERID=" & cQuote & cString & cQuote, objdbRs, False
    oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("FULLNAME"), "")
End Sub

Sub SSSR3_Migration(ByVal bMode As Boolean, Optional ByVal bRpt As Boolean)
    On Error GoTo ErrLoad
    Dim cSqlStmt As String, _
        nCombo As Integer, _
        oFile As File, oTextFile As New FileSystemObject, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        cParam2 As String, cCompanyName As String, cR3file As String, _
        nPClose As Integer, nCtr As Integer, _
        aEmpInfo As Variant, aERPInfo As Variant
        
    'Check info if meron value
    If Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
        MsgBox "Please supply missing Entry!!!"
        Exit Sub
    End If
        
        
    If bMode <> False Then
        Dim ExcelObj As Object
        Dim ExcelBook As Object
        Dim ExcelSheet As Object
        
        Dim nRowlast As Long, _
            nRowCnt As Long
        
        Dim aSSSInfo, aFirstRowInfo, aSizeInfor, aUserInfo As Variant
        
        Dim oTxtStream As TextStream, _
            cLogFirstRow, cLogDatFileName, cDatFileName, cLogSSSDetail, cLogLastData As String
            
        Dim nSSSTotAmt, nSSSTotEC, nSSSTolEmp As Long
            
        Dim nAnsTest As Integer
        
        Dim sAnsString_One, sAnsString_Two As String
        
        aSSSInfo = Array("", "", "", "", 0#, 0#)
    '    aSSSInfo(0) = "SSS Number"
    '    aSSSInfo(1) = "Last Name"
    '    aSSSInfo(2) = "First Name"
    '    aSSSInfo(3) = "MI"
    '    aSSSInfo(4) = "SSS Amount"
    '    aSSSInfo(5) = "SSS EE"
        
        aFirstRowInfo = Array("", "", "", "", "", "", "", "", "", "")
    '    aFirstRowInfo(0) = "Company Name"
    '    aFirstRowInfo(1) = "For the month format mmyyyy"
    '    aFirstRowInfo(2) = "Employer SSS Number"
    '    aFirstRowInfo(3) = "Reciept + GE"
    '    aFirstRowInfo(4) = "Date Paid format mmddyyyy"
    '    aFirstRowInfo(5) = "Total Amoount XXXXXXXXX.XX"
    '    aFirstRowInfo(6) = "current date format mmdd"
    '    aFirstRowInfo(7) = "current time format hhmm"
    '    aFirstRowInfo(8) = "Constant vaue XXXX0.00"
    '    aFirstRowInfo(9) = "NO"
    
        aSizeInfor = Array("", "", "", "", "", "", "", "", "")
    '    aSizeInfor(0) = "6"
    '    aSizeInfor(1) = "7"
    '    aSizeInfor(2) = "8"
    '    aSizeInfor(3) = "9"
    '    aSizeInfor(4) = "10"
    '    aSizeInfor(5) = "11"
    '    aSizeInfor(6) = "12"
    '    aSizeInfor(7) = "13"
    '    aSizeInfor(8) = "14"
        
        nRowlast = 5000
        nSSSTotAmt = 0
        nSSSTotEC = 0
        nSSSTolEmp = 0
    End If
    
    CreateTMPSSS
    
    aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
    
    aERPInfo = Array("", "")
    
    
    nCombo = IIf((Combo1.ListIndex + 1) > 11, 12, (Combo1.ListIndex + 1))
    
    
    aUserInfo = Array("", "", "", "", "")
    
    If Not ChkPersonnel(Text7) Then Exit Sub
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If
    
    
    cSqlStmt = " SELECT periodid,date_start,date_end FROM pa7730 " & _
               " where (13month=0) and (month(date_start) = " & nCombo & " and  month(date_end) = " & nCombo & ") and " & _
               " (year(date_start) = " & cQuote & Combo2.Text & cQuote & " and year(date_end) = " & cQuote & Combo2.Text & cQuote & ")"
               
    'Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRSet, False
    cParam2 = ""
    If oRSet.RecordCount > 0 Then
        If oRSet.RecordCount = 1 Then
            GoTo ErrLoad
        End If

        While Not oRSet.EOF
            cParam2 = IIf(cParam2 = "", cQuote & oRSet("periodid"), cParam2 & cQuote & "," & cQuote & oRSet("periodid") & cQuote)
            oRSet.MoveNext
        Wend
    End If
    
    If cParam2 = "" Then GoTo ErrLoad:
    
    cSqlStmt = "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215  " & _
               "FROM pah87263 a left join di3670 b on a.empid=b.empid " & _
               "  left join pah87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
               "group by a.empid " & _
               "Union All " & _
               "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215 " & _
               "FROM pa87263 a left join di3670 b on a.empid=b.empid " & _
               "  left join pa87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & "001" & cQuote & " ) and ((a.ded_amt + a.ded_amt2) <> 0) " & _
               "group by a.empid " & _
               "order by empid, periodid "
    'Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
            
            aERPInfo = Array("", "")
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            
            cSqlStmt = "select empid from TMPSSS where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
            QueryTemp cSqlStmt, objdbRs, False
            If Not objdbRs.RecordCount > 0 Then
'                MsgBox "insert"
                   
                OpenQueryDNS "Select * FROM pa37722 where costcenterid = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
                    If objdbRs.RecordCount > 0 Then
                        aERPInfo(0) = objdbRs("DESCRIPTION")
                    Else
                        aERPInfo(0) = ""
                    End If
    
    
                OpenQueryDNS "SELECT * FROM pa97722 where workcenterid = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
                    If objdbRs.RecordCount > 0 Then
                            aERPInfo(1) = objdbRs("DESCRIPTION")
                    Else
                            aERPInfo(1) = ""
                    End If
                
                aEmpInfo(0) = oRecordSet("empid")                                                           'ID
                aEmpInfo(1) = left(oRecordSet("ssnum"), 10)                                                 'ssnum
                aEmpInfo(2) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("lastname"), 15))))                'esurn
                aEmpInfo(3) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("firstname"), 15))))               'ename
                aEmpInfo(4) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("mname"), 1))))                    'emidinit
                aEmpInfo(5) = oRecordSet("ded_amt")                                                         'ssamt
                aEmpInfo(6) = oRecordSet("sser") + oRecordSet("sser1215")
                aEmpInfo(7) = oRecordSet("ssprem") + oRecordSet("ssprem1215")
                
                cSqlStmt = "INSERT INTO TMPSSS(EMPID,SSSNO,LASTNAME,FIRSTNAME,MINAME,SSSAMT,EC_AMT,DATE_REM,REM,ER_AMT,EE_AMT,COSTID,CDESC,WORKID,WDESC)VALUES(" & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & aEmpInfo(1) & cQuote & "," & _
                       cQuote & aEmpInfo(2) & cQuote & "," & _
                       cQuote & aEmpInfo(3) & cQuote & "," & _
                       cQuote & aEmpInfo(4) & cQuote & "," & _
                       aEmpInfo(5) & "," & _
                       IIf(aEmpInfo(5) >= 1650, "30.00", IIf(aEmpInfo(5) > 0, "10.00", "0.00")) & ",0," & cQuote & "N0" & cQuote & "," & _
                       aEmpInfo(6) & "," & _
                       aEmpInfo(7) & "," & _
                       cQuote & oRecordSet("costcenterid") & cQuote & "," & _
                       cQuote & aERPInfo(0) & cQuote & "," & _
                       cQuote & oRecordSet("workcenterid") & cQuote & "," & _
                       cQuote & aERPInfo(1) & cQuote & ")"
                
'                    MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, False
'                nCtr = nCtr + 1
            
            Else
                
                cSqlStmt = "update TMPSSS set " & _
                           " SSSAMT = SSSAMT + " & oRecordSet("ded_amt") & "," & _
                           " ER_AMT = " & oRecordSet("sser") + oRecordSet("sser1215") & "," & _
                           " EE_AMT = " & oRecordSet("ssprem") + oRecordSet("ssprem1215") & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote

'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
                
                '               ---> update 201506-03

                cSqlStmt = "update TMPSSS set " & _
                           " EC_AMT = iif (SSSAMT >= 1650 , 30.00, iif(SSSAMT > 0, 10, 0) ) " & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
                           
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
                
            End If
            
            oRecordSet.MoveNext
        Wend
        
        
        If bMode <> False Then
            'dito na papasok yung r3
            CreateTMPSSSREM
            
            CreateTMPSSSDAT
            
            Set ExcelObj = CreateObject("Excel.Application")
            Set ExcelSheet = CreateObject("Excel.Sheet")
    
            ExcelObj.Workbooks.Open Text6.Text
    
            Set ExcelBook = ExcelObj.Workbooks(1)
            Set ExcelSheet = ExcelBook.Worksheets(1)
        
            aFirstRowInfo(0) = UCase(cCompany)
            aFirstRowInfo(1) = Combo1.ListIndex + 1
            
            aFirstRowInfo(1) = IIf(Len(aFirstRowInfo(1)) > 1, aFirstRowInfo(1), "0" & aFirstRowInfo(1)) & Combo2.Text
            aFirstRowInfo(2) = Replace(gSSSNum, "-", "")
            aFirstRowInfo(3) = IIf(Text3.Text = "", "NOPAY", Text3.Text)
            aFirstRowInfo(4) = Format(XPDatePicker1.CurrentDate, "MMDDYYYY")
            aFirstRowInfo(5) = PadStr(Format(Round(Val(Text5.Text), 2), "0.00"), "0", 12)
            aFirstRowInfo(6) = Format(Now, "MMDD")
            aFirstRowInfo(7) = Format(Now, "HHMM")
            aFirstRowInfo(8) = "NO"
            
            aSizeInfor(0) = PadStr("0.00", " ", 6)
            aSizeInfor(1) = PadStr("0.00", " ", 7)
            aSizeInfor(2) = PadStr("0.00", " ", 8)
            aSizeInfor(3) = PadStr("0.00", " ", 9)
            aSizeInfor(4) = PadStr("0.00", " ", 10)
            aSizeInfor(5) = PadStr("0.00", " ", 11)
            aSizeInfor(6) = PadStr("0.00", " ", 12)
            aSizeInfor(7) = PadStr("0.00", " ", 13)
            aSizeInfor(8) = PadStr("0.00", " ", 14)
            
            cLogDatFileName = "R3" & aFirstRowInfo(2) & aFirstRowInfo(1) & "." & aFirstRowInfo(6) & aFirstRowInfo(7)
            cDatFileName = cLogDatFileName
            cLogDatFileName = CheckPath(Text4.Text) & cLogDatFileName
            
            cLogFirstRow = "00" & PadStr(aFirstRowInfo(0), " ", 30, 1) & aFirstRowInfo(1) & aFirstRowInfo(2) & PadStr(aFirstRowInfo(3), " ", 10, 1) & _
                           aFirstRowInfo(4) & aFirstRowInfo(5)
            
            If Dir(cLogDatFileName) = "" Then
                Set oTxtStream = oTextFile.CreateTextFile(cLogDatFileName, True)
            Else
                Set oFile = oTextFile.GetFile(cLogDatFileName)
                Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
            End If
            
            oTxtStream.WriteLine cLogFirstRow
            
    '        read excel data
            For nRowCnt = 3 To nRowlast Step 1
                
                With ExcelSheet
                    If .Cells(nRowCnt, 1) <> "" Then
                    
                        ShowProgress 2, (1 / nRowlast) * 100, , , "Copying " & .Cells(nRowCnt, 1) & "... " & Trim(.Cells(nRowCnt, 2) & ", " & .Cells(nRowCnt, 3)) & "..."
                    
                        cLogSSSDetail = ""
                        
                        aSSSInfo(0) = .Cells(nRowCnt, 1)
                        aSSSInfo(1) = left(.Cells(nRowCnt, 2), 15)
                        aSSSInfo(2) = left(.Cells(nRowCnt, 3), 15)
                        aSSSInfo(3) = IIf(.Cells(nRowCnt, 4) = "", " ", .Cells(nRowCnt, 4))
                        aSSSInfo(4) = Format(Round(IIf(.Cells(nRowCnt, 5) = "", 0, .Cells(nRowCnt, 5)), 2), "0.00")
                        aSSSInfo(5) = Format(Round(IIf(.Cells(nRowCnt, 6) = "", 0, .Cells(nRowCnt, 6)), 2), "0.00")
                        
'                            nSSSTotAmt = nSSSTotAmt + Val(.Cells(nRowCnt, 5))
'                            nSSSTotEC = nSSSTotEC + Val(.Cells(nRowCnt, 6))
'                            nSSSTolEmp = nSSSTolEmp + 1
                        cSqlStmt = "INSERT INTO TMPSSS(LASTNAME , FIRSTNAME, MINAME, SSSNO, SSSAMT, EC_AMT, REM)VALUES(" & _
                                    cQuote & aSSSInfo(1) & cQuote & "," & _
                                    cQuote & aSSSInfo(2) & cQuote & "," & _
                                    cQuote & aSSSInfo(3) & cQuote & "," & _
                                    cQuote & aSSSInfo(0) & cQuote & "," & _
                                    cQuote & aSSSInfo(4) & cQuote & "," & _
                                    cQuote & aSSSInfo(5) & cQuote & "," & _
                                    cQuote & "N0" & cQuote & ")"
                        QueryTemp cSqlStmt, objdbRs, True
                    Else
                    
                        ExcelObj.Workbooks.Close
                        ExcelObj.Quit
                        Set ExcelObj = Nothing
                        Exit For
                    End If
                End With
    
            Next nRowCnt
            
        
            cSqlStmt = "select LASTNAME , FIRSTNAME, MINAME, SSSNO, SSSAMT, EC_AMT, REM from TMPSSS ORDER BY LASTNAME, FIRSTNAME, MINAME"
            QueryTemp cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
                While Not oRecordSet.EOF
                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("SSSNO")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
                    
                    aSSSInfo(0) = oRecordSet("SSSNO")
                    aSSSInfo(1) = left(UCase(EncodeStr2(DecodeStr(oRecordSet("LASTNAME")))), 15)
                    aSSSInfo(2) = left(UCase(EncodeStr2(DecodeStr(oRecordSet("FIRSTNAME")))), 15)
                    aSSSInfo(3) = left(UCase(EncodeStr2(DecodeStr(oRecordSet("MINAME")))), 1)
                    aSSSInfo(4) = Format(Round(oRecordSet("SSSAMT"), 2), "0.00")
                    aSSSInfo(5) = Format(Round(oRecordSet("EC_AMT"), 2), "0.00")
                    
                    nSSSTotAmt = nSSSTotAmt + Val(aSSSInfo(4))
                    nSSSTotEC = nSSSTotEC + Val(aSSSInfo(5))
                    nSSSTolEmp = nSSSTolEmp + 1
                    
                    'sAnsString_One , sAnsString_Two
                    
                    If Val(aFirstRowInfo(1)) = 1 Or Val(aFirstRowInfo(1)) = 4 Or Val(aFirstRowInfo(1)) = 7 Or Val(aFirstRowInfo(1)) = 10 Then
                        sAnsString_One = PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2) & aSizeInfor(2)
                        sAnsString_Two = PadStr(aSSSInfo(5), " ", 6) & aSizeInfor(0) & aSizeInfor(0)
                    ElseIf Val(aFirstRowInfo(1)) = 2 Or Val(aFirstRowInfo(1)) = 5 Or Val(aFirstRowInfo(1)) = 8 Or Val(aFirstRowInfo(1)) = 11 Then
                        sAnsString_One = aSizeInfor(2) & PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2)
                        sAnsString_Two = aSizeInfor(0) & PadStr(aSSSInfo(5), " ", 6) & aSizeInfor(0)
                    Else
                        sAnsString_One = aSizeInfor(2) & aSizeInfor(2) & PadStr(aSSSInfo(4), " ", 8)
                        sAnsString_Two = aSizeInfor(0) & aSizeInfor(0) & PadStr(aSSSInfo(5), " ", 6)
                    End If
                    
'                    cLogSSSDetail = "20" & _
'                                    PadStr(UCase(aSSSInfo(1)), " ", 15, 1) & PadStr(UCase(aSSSInfo(2)), " ", 15, 2) & UCase(aSSSInfo(3)) & _
'                                    aSSSInfo(0) & _
'                                    aSizeInfor (2) & PadStr(aSSSInfo(4), " ", 8) & aSizeInfor(2) & _
'                                    aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & _
'                                    aSizeInfor(0) & PadStr(aSSSInfo(5), " ", 6) & aSizeInfor(0) & _
'                                    PadStr(oRecordset("REM"), " ", 8)


                    cLogSSSDetail = "20" & _
                                    PadStr(UCase(aSSSInfo(1)), " ", 15, 1) & PadStr(UCase(aSSSInfo(2)), " ", 15, 2) & UCase(aSSSInfo(3)) & _
                                    aSSSInfo(0) & _
                                    sAnsString_One & _
                                    aSizeInfor(0) & aSizeInfor(0) & aSizeInfor(0) & _
                                    sAnsString_Two & _
                                    PadStr(oRecordSet("REM"), " ", 8)
                    
                    oTxtStream.WriteLine cLogSSSDetail
                    
                    
                        cSqlStmt = "INSERT INTO TMPSSSDAT(DATFNAME, COMPANYNAME, CURDATE, LASTNAME, FIRSTNAME, MI, SSSNO, SSSAMT,SSSEE, REM,DTHRD)VALUES(" & _
                                    cQuote & Trim(cDatFileName) & cQuote & "," & _
                                    cQuote & Trim(aFirstRowInfo(0)) & cQuote & "," & _
                                    cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                                    cQuote & Trim(oRecordSet("LASTNAME")) & cQuote & "," & _
                                    cQuote & Trim(oRecordSet("FIRSTNAME")) & cQuote & "," & _
                                    cQuote & Trim(oRecordSet("MINAME")) & cQuote & "," & _
                                    cQuote & oRecordSet("SSSNO") & cQuote & "," & _
                                    Val(oRecordSet("SSSAMT")) & "," & _
                                    Val(oRecordSet("EC_AMT")) & "," & _
                                    cQuote & Trim(oRecordSet("REM")) & cQuote & ",0)"
                    
'                        Script2File cSqlStmt
                        QueryTemp cSqlStmt, objdbRs, True
                    
                    oRecordSet.MoveNext
                    
                Wend
            End If
        
            sAnsString_One = ""
            sAnsString_Two = ""
            
            If Val(aFirstRowInfo(1)) = 1 Or Val(aFirstRowInfo(1)) = 4 Or Val(aFirstRowInfo(1)) = 7 Or Val(aFirstRowInfo(1)) = 10 Then
                sAnsString_One = PadStr(Format(nSSSTotAmt, "0.00"), " ", 12) & aSizeInfor(6) & aSizeInfor(6)
                sAnsString_Two = PadStr(Format(nSSSTotEC, "0.00"), " ", 10) & aSizeInfor(3) & aSizeInfor(4)
            ElseIf Val(aFirstRowInfo(1)) = 2 Or Val(aFirstRowInfo(1)) = 5 Or Val(aFirstRowInfo(1)) = 8 Or Val(aFirstRowInfo(1)) = 11 Then
                sAnsString_One = aSizeInfor(6) & PadStr(Format(nSSSTotAmt, "0.00"), " ", 12) & aSizeInfor(6)
                sAnsString_Two = aSizeInfor(3) & PadStr(Format(nSSSTotEC, "0.00"), " ", 10) & aSizeInfor(4)
            Else
                sAnsString_One = aSizeInfor(6) & aSizeInfor(6) & PadStr(Format(nSSSTotAmt, "0.00"), " ", 12)
                sAnsString_Two = aSizeInfor(3) & aSizeInfor(4) & PadStr(Format(nSSSTotEC, "0.00"), " ", 10)
            End If
       
            'end data
'            cLogLastData = "99" & _
'                           aSizeInfor(6) & _
'                           PadStr(Format(nSSSTotAmt, "0.00"), " ", 12) & aSizeInfor(6) & aSizeInfor(3) & _
'                           aSizeInfor(5) & _
'                           aSizeInfor(5) & _
'                           aSizeInfor(3) & PadStr(Format(nSSSTotEC, "0.00"), " ", 10) & aSizeInfor(4)
                           
            cLogLastData = "99" & _
                           sAnsString_One & _
                           aSizeInfor(3) & _
                           aSizeInfor(5) & _
                           aSizeInfor(5) & _
                           sAnsString_Two
                           
            oTxtStream.WriteLine cLogLastData
            
            'Transmmital report
            cSqlStmt = "INSERT INTO TMPSSSREM(DATFNAME,COMPANYNAME,SSSEMP,APP_PER,TR_SBR_NO,DATE_PAY,PAY_AMT,SS_AMT,EC_AMT,TOT_AMT,ECOUNT,SIGNATORY1,POSNAME1)VALUES(" & _
                        cQuote & cDatFileName & cQuote & "," & _
                        cQuote & UCase(aFirstRowInfo(0)) & cQuote & "," & _
                        cQuote & gSSSNum & cQuote & "," & _
                        cQuote & aFirstRowInfo(1) & cQuote & "," & _
                        cQuote & aFirstRowInfo(3) & cQuote & "," & _
                        cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & "," & _
                        IIf(Val(aFirstRowInfo(5)) = 0, nSSSTotAmt + nSSSTotEC, Val(aFirstRowInfo(5))) & "," & _
                        nSSSTotAmt & "," & _
                        nSSSTotEC & "," & _
                        nSSSTotAmt + nSSSTotEC & "," & _
                        nSSSTolEmp & "," & _
                        cQuote & UCase(EncodeStr2(DecodeStr(Label11.Caption))) & cQuote & "," & _
                        cQuote & UCase(aUserInfo(0)) & cQuote & ")"
                        
    '        MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
        
        End If
        
        ShowProgress 3
        If bMode <> False Then
            If bRpt Then
                'dito and list
                GenerateReport "SSS R3 Data File Report", "rptSSSR3LIST.rpt"
            Else
            
                GenerateReport "SSS R3 Data File Report", "rptSSSR3TRANS.rpt"
            End If
        Else
            GenerateReport "SSS R3 Data File Report", "rptSSSR3.rpt"
        End If
        
        ShowProgress 4
    End If
    
    Exit Sub
    
ErrLoad:
    MsgBox "Error Generating SSS..."

End Sub
