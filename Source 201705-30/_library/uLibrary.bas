Attribute VB_Name = "uLibrary"
'                  $`````$
'                $( o  o )$
'    >------oOO--(_)--OOo------------------------------------------------------------------------------<
'    "Intelligent people can be bored. They just know a lot. But Smart people
'    are never bored, because they're always looking for something to engage
'    their minds."
'    >------oooo(O) (0)oooo----------------------------------------------------------------------------<

' project name  :   Dong-in Payroll & Time Management System
' module        :   uLibrary
' programmer    :   _-=[ srm ]=-_
' date          :   22 oct 2005

Option Explicit
    
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const mrNone = 0
Public Const mrOk = 1
Public Const mrCancel = 2
    
Public Const PadLeft = 0
Public Const PadRight = 1

Public Const SW_SHOWNOACTIVATE = 40

    
Public nAccess_Tag As Integer

Public lSuperUser As Boolean
    
Public gUserID As String, _
       gUserPW As String, _
       gUserName As String, _
       gWSID As String, _
       gCompanyID As String, _
       gAddress As String, _
       gUserGroup As Integer, _
       gUserLevel As Integer        ' --> 02072005
       
Public cMGR_CODE As String, _
       cMGR_NAME As String
       
Public cConnString, _
       cCompany, _
       cTempPath, _
       cReportPath, _
       cScriptPath, _
       cLogPath, _
       cImgFile, _
       cResult As String
       
' --> Prefix and codelen, for Cost Center code use... 20120822
Public cPrefix As String, _
       nCodeLen As Integer, _
       nCompCode As Integer

Public cODBC, _
       cUser, _
       cPwd, _
       cDriver

Public Const cQuote = """"

Public oTempConn As New ADODB.Connection        ' --> for access connection
Public objdbConn As New ADODB.Connection        ' --> for mysql connection
Public oSSSConn As New ADODB.Connection        ' --> for access connection
Public objdbRs As New ADODB.Recordset
Public oControl As Control

Public ModalResult As Integer

Public nStart, nDuration, nCurTime, nElapse As Double
Public nOldValue As Integer

Public aTaxAmt As Variant
Public aTaxPct As Variant

Public cDownloadPath, _
       cUploadPath As String     ' --> download/upload path for transfer,po,down/up 2 server, etc...
       
Public gBasicRate, _
       gColaAmt As Double

Public m_FrmProgess As frmProgress      ' --> 20060324


' --> 20060524, variable to handle Special Assessment ID & Amount...
Public gAssessID As String, _
       nAssessAmt As Double
       
       
' --> 20060726, variable for Company's misc #...
Public gPHealthNum, _
       gSSSNum, _
       gTINNum, _
       gTelNum, _
       gAreaNo As String
       
Public lCheckFC As Boolean


' --> 20060822, variable for salary division use...
Public gAdmin, _
       gStaff As String

Public gServerIP        ' --> variable for server's IP add... 20060908
Public nPort            ' --> for FTP use only...


' --> 20061003, Deduction ID to be excluded for tax computation
Public aTaxExempt As Variant


' --> 20061129, path for TimeKeeper database (timekeeper.mdb)
Public gTimeKeeperPath As String

Public gCashAdvance As String

' --> 20070117, flag to check late...
Public lCheckLate As Boolean


' --> 20070328, flag to check Extension
Public lExtension As Boolean


' --> 20070601
' divisor for OT (default to 1800 - 30 minutes)
Public nOTInterval As Double

' night diff time (default to 10:00 PM)
Public gNDiffTime As String


' --> flag for automatic downloading of data from bioclock, 20070621
Public lAttendance As Double


Public gAgency As String



' --> tag to determine whether gross or basic for computation of 13th month pay
Public g13Month As Boolean
' where
'   0   gross   (default)
'   1   basic

Public gPostal As String

Public gDepid As String

' --> 20090620, flag to check Audit (Combine or not)
Public lAudit As Integer

Public gBAccntNo As String

Public gRCBCNo As String

'20091210
Public gECA As String

'20131119
Public gServer As String

Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal SectionName$, ByVal KeyName$, ByVal Default$, ByVal ReturnedString$, ByVal Size&, ByVal FileName$)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


Public Sub Main()
    ' check instance of running application
    If App.PrevInstance Then
        App.Title = "... duplicate instance."      'Pretty, eh?
        MsgBox "Application is already running", vbExclamation
        End
    End If
        
    If Not ReadINI Then
        MsgBox "Initial configuration not found!"
        End
    End If
        
        
    If gServer = 1 Then
        
        If Not DetectServer Then
            MsgBox "Connection failed!"
            End
        End If
    
        Load frmSplashConfig
        frmSplashConfig.Show
    Else
        Load frmSplash
        frmSplash.Show 0
        Add2List "Loading Cost Accounting System, please wait..."
        DoEvents
    
        Add2List "Please wait reading from initial configuration"
    
        If Not ReadINI Then
            Add2List "Initial configuration not found!"
            End
        End If
        Add2List "Initial configuration successfully initiated..."
    
        Write2File "Session start at " & Now
    
        Add2List "Please wait establishing connection..."
        If Not DetectServer Then
            Add2List "Connection failed!"
            End
        End If
        Add2List "Connection successful!"
    
    '    ChkTable
    
        Log2Audit "frmMain", "Open " & App.Title & " version " & PadStr(App.Major, "0", 2) & "." & PadStr(App.Minor, "0", 3) & "." & PadStr(App.Revision, "0", 4)
    
        Add2List "Please wait initializing temporary database"
        If Not DetectTemp Then
            Add2List "Temporary database initialization failed!"
            End
        End If
    
        Add2List "Accessing the system..."
    
        Log2Audit "frmMain", "Session started."
    
        Load frmMain
        frmMain.showLogin
    
    '    Load frmWebReport
    '    frmWebReport.WebBrowser1.Navigate "http://localhost"
    '    frmWebReport.Show
    
        With frmMain
            .Show
        End With
    End If
    
    
    
    
    
'    Load frmSplash
'    frmSplash.Show 0
'    Add2List "Loading Cost Accounting System, please wait..."
'    DoEvents
'
'    Add2List "Please wait reading from initial configuration"
'
'    If Not ReadINI Then
'        Add2List "Initial configuration not found!"
'        End
'    End If
'    Add2List "Initial configuration successfully initiated..."
'
'    Write2File "Session start at " & Now
'
'    Add2List "Please wait establishing connection..."
'    If Not DetectServer Then
'        Add2List "Connection failed!"
'        End
'    End If
'    Add2List "Connection successful!"
'
''    ChkTable
'
'    Log2Audit "frmMain", "Open " & App.Title & " version " & PadStr(App.Major, "0", 2) & "." & PadStr(App.Minor, "0", 3) & "." & PadStr(App.Revision, "0", 4)
'
'    Add2List "Please wait initializing temporary database"
'    If Not DetectTemp Then
'        Add2List "Temporary database initialization failed!"
'        End
'    End If
'
'    Add2List "Accessing the system..."
'
'    Log2Audit "frmMain", "Session started."
'
'    Load frmMain
'    frmMain.showLogin
'
''    Load frmWebReport
''    frmWebReport.WebBrowser1.Navigate "http://localhost"
''    frmWebReport.Show
'
'    With frmMain
'        .Show
'    End With
    
End Sub


Public Function ReadINI() As Boolean
    Dim nRetVal As Long
    Dim cINIFile, _
        cINIValue As String, _
        cString As String
    Dim nImpex, _
        nOffer1, _
        nOffer2, _
        nOffer3, _
        nOffer4, _
        nOffer5 As Double, _
        nCtr As Integer
        
    cINIFile = App.Path & "\" & App.EXEName & ".INI"
    cINIValue = String(255, 0)
    
    If Trim(Dir(cINIFile, vbNormal)) <> "" Then
        nRetVal = GetPrivateProfileString("MAIN", "APPLICATION", "Payroll and Time Management System", cINIValue, 255, cINIFile)
        App.Title = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "Payroll and Time Management System")
        
        Add2List "Retrieving company information..."
        nRetVal = GetPrivateProfileString("MAIN", "COMPANY", "Creative Mind", cINIValue, 255, cINIFile)
        cCompany = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "Creative Mind")
        
        DoEvents
'        frmInit.lblCompany.Caption = App.Title
'        frmInit.Label1.Caption = "Copyright " & Chr$(169) & " 2005 " & cCompany & Chr$(13) & Chr$(10) & _
                                 "All Rights " & Chr$(174) & " Reserved."
    
        Add2List "Retrieving company address..."
        nRetVal = GetPrivateProfileString("MAIN", "ADDRESS", "Balanga, Bataan 2100", cINIValue, 255, cINIFile)
        gAddress = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "Balanga, Bataan 2100")
        
        
        Add2List "Please wait retrieving workstation ID"
        ' --> added 2005 feb 05
        nRetVal = GetPrivateProfileString("MAIN", "WORKSTATION", "", cINIValue, 255, cINIFile)
        gWSID = IIf(nRetVal > 0, left$(cINIValue, nRetVal), gWSID)
        ' --> end of WorkStation ID tag
        Add2List "Workstation ID set to " & gWSID
        
        ' --> added 2005 feb 07
        nRetVal = GetPrivateProfileString("MAIN", "COMPANYID", "", cINIValue, 255, cINIFile)
        gCompanyID = IIf(nRetVal > 0, left$(cINIValue, nRetVal), gCompanyID)
        ' --> end of company ID tag
        
        nRetVal = GetPrivateProfileString("MAIN", "WALLPAPER", "", cINIValue, 255, cINIFile)
        cImgFile = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        
        Add2List "Please wait retrieving temporary path location..."
        nRetVal = GetPrivateProfileString("PATH", "TEMPPATH", "", cINIValue, 255, cINIFile)
        cTempPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cTempPath)
        If Trim(cTempPath) = "" Then
            MsgBox "Temporary path is undefined!", vbCritical, App.Title
            ReadINI = False
            Exit Function
        Else
            cTempPath = CheckPath(cTempPath)
            If Dir(cTempPath, vbDirectory) = "" Then
                MkDir cTempPath
            End If
        End If
        Add2List "Temporary path set to " & cTempPath
        
        
        Add2List "Please wait retrieving report path location..."
        nRetVal = GetPrivateProfileString("PATH", "REPORTPATH", "", cINIValue, 255, cINIFile)
        cReportPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cReportPath)
        If Trim(cReportPath) = "" Then
            MsgBox "Temporary path is undefined!", vbCritical, App.Title
            ReadINI = False
            Exit Function
        Else
            cReportPath = CheckPath(cReportPath)
            If Dir(cReportPath, vbDirectory) = "" Then
                MkDir cReportPath
            End If
        End If
        Add2List "Report path set to " & cReportPath
        
        
        Add2List "Please wait retrieving script path location..."
        nRetVal = GetPrivateProfileString("PATH", "SCRIPTPATH", "", cINIValue, 255, cINIFile)
        cScriptPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cScriptPath)
        If Trim(cScriptPath) = "" Then
            MsgBox "Script path is undefined!", vbCritical, App.Title
            ReadINI = False
            Exit Function
        Else
            cScriptPath = CheckPath(cScriptPath)
            If Dir(cScriptPath, vbDirectory) = "" Then
                MkDir cScriptPath
            End If
        End If
        Add2List "Script path set to " & cScriptPath
        
        
        Add2List "Please wait retrieving path for log file"
        nRetVal = GetPrivateProfileString("PATH", "LOGPATH", "", cINIValue, 255, cINIFile)
        cLogPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cLogPath)
        If Trim(cLogPath) = "" Then
            cLogPath = CheckPath(App.Path)
        End If
        If Dir(cLogPath, vbDirectory) = "" Then
            MkDir cLogPath
        End If
        Add2List "Log path set to " & cLogPath
        
        
        ' --> upload path
        Add2List "Please wait retrieving path for upload"
        nRetVal = GetPrivateProfileString("PATH", "UPLOAD", "", cINIValue, 255, cINIFile)
        cUploadPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cUploadPath)
        If Trim(cUploadPath) = "" Then
'            cUploadPath = CheckPath(App.Path)
            MsgBox "Path for uploading is undefined!", vbCritical, App.Title
            ReadINI = False
            Exit Function
        Else
            cUploadPath = CheckPath(cUploadPath)
        End If
        If Dir(cUploadPath, vbDirectory) = "" Then
            MkDir cUploadPath
        End If
        Add2List "Upload path set to " & cUploadPath
        
        
        ' --> 20060309 - auto-create bioclock folder here...
        If Dir(cUploadPath & "Bioclock", vbDirectory) = "" Then
            MkDir cUploadPath & "Bioclock"
        End If
        
        
        ' --> added 20060908
        Add2List "Retrieving Server information..."
        nPort = GetPrivateProfileInt("SOCKET", "PORT", 1007, cINIFile)
        
        nRetVal = GetPrivateProfileString("SOCKET", "IPADD", "127.0.0.1", cINIValue, 255, cINIFile)
        gServerIP = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "127.0.0.1")
        
        
        ' --> added 20050127
        ' --> for odbc connection
        Add2List "Please wait building connection string..."
        nRetVal = GetPrivateProfileString("ODBC", "DSN", cODBC, cINIValue, 255, cINIFile)
        cODBC = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cODBC)
        
        nRetVal = GetPrivateProfileString("ODBC", "USERID", cUser, cINIValue, 255, cINIFile)
        cUser = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cUser)
        
        nRetVal = GetPrivateProfileString("ODBC", "PASSWORD", cPwd, cINIValue, 255, cINIFile)
        cPwd = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cPwd)
        
        ' --> dsn-less connection - 20060908
        '
        cConnString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                      "SERVER=" & gServerIP & ";" & _
                      "DATABASE=" & cODBC & ";" & _
                      "USER=" & cUser & ";" & _
                      "PASSWORD=" & cPwd & ";" & _
                      "OPTION=11;"
        ' --> end of odbc
        
        ' --> added prefix and codelen for Cost Center code 20120822
        nRetVal = GetPrivateProfileString("COSTCENTER", "PREFIX", "D", cINIValue, 255, cINIFile)
        cPrefix = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "D")
        nCodeLen = GetPrivateProfileInt("COSTCENTER", "CODELEN", 10, cINIFile)
        
        nCompCode = GetPrivateProfileInt("COSTCENTER", "COMPCODE", 4, cINIFile)
        ' --> end of prefix
        
        ' --> Tax Exempt Amount
        aTaxAmt = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT1", "", cINIValue, 255, cINIFile)
        aTaxAmt(0) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(0))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT2", "", cINIValue, 255, cINIFile)
        aTaxAmt(1) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(1))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT3", "", cINIValue, 255, cINIFile)
        aTaxAmt(2) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(2))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT4", "", cINIValue, 255, cINIFile)
        aTaxAmt(3) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(3))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT5", "", cINIValue, 255, cINIFile)
        aTaxAmt(4) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(4))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT6", "", cINIValue, 255, cINIFile)
        aTaxAmt(5) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(5))
        
        nRetVal = GetPrivateProfileString("TAX", "AMOUNT7", "", cINIValue, 255, cINIFile)
        aTaxAmt(6) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxAmt(6))
        
        ' --> Tax Exemption Percent
        aTaxPct = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        nRetVal = GetPrivateProfileString("TAX", "PERCENT1", "", cINIValue, 255, cINIFile)
        aTaxPct(0) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(0))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT2", "", cINIValue, 255, cINIFile)
        aTaxPct(1) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(1))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT3", "", cINIValue, 255, cINIFile)
        aTaxPct(2) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(2))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT4", "", cINIValue, 255, cINIFile)
        aTaxPct(3) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(3))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT5", "", cINIValue, 255, cINIFile)
        aTaxPct(4) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(4))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT6", "", cINIValue, 255, cINIFile)
        aTaxPct(5) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(5))
        
        nRetVal = GetPrivateProfileString("TAX", "PERCENT7", "", cINIValue, 255, cINIFile)
        aTaxPct(6) = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), aTaxPct(6))
        
        
        ' --> default Basic Rate
        nRetVal = GetPrivateProfileString("SETUP", "RATE", "", cINIValue, 255, cINIFile)
        gBasicRate = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), 0)
        
        ' --> default COLA Amount
        nRetVal = GetPrivateProfileString("SETUP", "COLA", "", cINIValue, 255, cINIFile)
        gColaAmt = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), 0)
        
        
        ' --> special assessment id
        nRetVal = GetPrivateProfileString("setup", "assessment", "", cINIValue, 255, cINIFile)
        gAssessID = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        lCheckFC = GetPrivateProfileInt("SETUP", "CHECKFC", 0, cINIFile)
        
        
        ' --> company's philhealth #
        nRetVal = GetPrivateProfileString("main", "philhealth", "", cINIValue, 255, cINIFile)
        gPHealthNum = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> company's sss #
        nRetVal = GetPrivateProfileString("main", "sss", "", cINIValue, 255, cINIFile)
        gSSSNum = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> company's tin #
        nRetVal = GetPrivateProfileString("main", "tin", "", cINIValue, 255, cINIFile)
        gTINNum = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> company's tel #
        nRetVal = GetPrivateProfileString("main", "phone", "", cINIValue, 255, cINIFile)
        gTelNum = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> company's area code #
        nRetVal = GetPrivateProfileString("main", "area", "", cINIValue, 255, cINIFile)
        gAreaNo = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> download path - 20060815
        Add2List "Please wait retrieving path for download"
        nRetVal = GetPrivateProfileString("PATH", "DOWNLOAD", "", cINIValue, 255, cINIFile)
        cDownloadPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), cDownloadPath)
        If Trim(cDownloadPath) = "" Then
'            cUploadPath = CheckPath(App.Path)
            MsgBox "Path for download is undefined!", vbCritical, App.Title
            ReadINI = False
            Exit Function
        Else
            cDownloadPath = CheckPath(cDownloadPath)
        End If
        If Dir(cDownloadPath, vbDirectory) = "" Then
            MkDir cDownloadPath
        End If
        Add2List "Download path set to " & cDownloadPath
        
        
        ' --> Admin ID - 20060822
        nRetVal = GetPrivateProfileString("setup", "admin", "", cINIValue, 255, cINIFile)
        gAdmin = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> Staff ID
        nRetVal = GetPrivateProfileString("setup", "staff", "", cINIValue, 255, cINIFile)
        gStaff = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("tax", "exempt", "", cINIValue, 255, cINIFile)
        cString = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> for tax exemption...
        ReDim aTaxExempt(99)
        While InStr(1, cString, ",") > 0
            aTaxExempt(nCtr) = cQuote & left(cString, InStr(1, cString, ",") - 1) & cQuote
            cString = Mid(cString, InStr(1, cString, ",") + 1, Len(cString) - InStr(1, cString, ","))
            nCtr = nCtr + 1
        Wend
        If Trim(cString) <> "" Then aTaxExempt(nCtr) = cQuote & cString & cQuote
        
        
        ' --> 20061129
        nRetVal = GetPrivateProfileString("path", "timekeeper", "", cINIValue, 255, cINIFile)
        gTimeKeeperPath = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        If Trim(gTimeKeeperPath) <> "" Then gTimeKeeperPath = CheckPath(gTimeKeeperPath)
        
        
        ' --> code for cash advance 20061211
        nRetVal = GetPrivateProfileString("setup", "advance", "", cINIValue, 255, cINIFile)
        gCashAdvance = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        
        ' --> tag to check late - 20070117
        lCheckLate = GetPrivateProfileInt("setup", "late", 0, cINIFile)
        
        
        ' --> tag to check extension - 20070328
        lExtension = GetPrivateProfileInt("setup", "extension", 1, cINIFile)
        
        
        ' --> divisor for OT interval (default to 30 mins - 1800) -- 20070601
        nRetVal = GetPrivateProfileString("SETUP", "OTInterval", "", cINIValue, 255, cINIFile)
        nOTInterval = IIf(nRetVal > 0, Val(left$(cINIValue, nRetVal)), 1800)
        
        ' --> pre-defined Night Diff Time, default to 10:00 PM
        nRetVal = GetPrivateProfileString("setup", "NDTIME", "", cINIValue, 255, cINIFile)
        gNDiffTime = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "22:00:00")
        
        
        ' --> tag to check attendance - 20070621
        lAttendance = GetPrivateProfileInt("setup", "attendance", 0, cINIFile)
        
        
        ' --> tag for 13th month pay computation, default to gross - 20071207
        g13Month = GetPrivateProfileInt("setup", "13month", 0, cINIFile)
        
        nRetVal = GetPrivateProfileString("main", "postal", "", cINIValue, 255, cINIFile)
        gPostal = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("setup", "depid", "", cINIValue, 255, cINIFile)
        gDepid = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("setup", "audit", 0, cINIValue, 255, cINIFile)
        lAudit = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("main", "baccntno", "", cINIValue, 255, cINIFile)
        gBAccntNo = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("main", "RCBC", "", cINIValue, 255, cINIFile)
        gRCBCNo = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        nRetVal = GetPrivateProfileString("costcenter", "agency", "", cINIValue, 255, cINIFile)
        gAgency = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> code for cash advance 20091210
        nRetVal = GetPrivateProfileString("setup", "ECA", "", cINIValue, 255, cINIFile)
        gECA = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ' --> code for sever selection 20131119
        nRetVal = GetPrivateProfileString("setup", "server", "", cINIValue, 255, cINIFile)
        gServer = IIf(nRetVal > 0, left$(cINIValue, nRetVal), "")
        
        ReadINI = True
    Else
        MsgBox "INI file not found!", vbCritical, "Payroll and Time Management System"
        ReadINI = False
    End If
End Function


Public Function DetectTemp() As Boolean
    On Error GoTo ErrDetect
    Dim oCatalog As ADOX.Catalog, _
        oTextFile As New FileSystemObject
    
    DoEvents
    
    ' --> delete temporary file if it's existing...
    If Dir(cTempPath & App.EXEName & ".MDB", vbNormal) <> "" Then
        oTextFile.DeleteFile cTempPath & App.EXEName & ".MDB"
        Set oTextFile = Nothing
    End If

    Set oCatalog = New ADOX.Catalog
    oCatalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & cTempPath & App.EXEName & ".MDB" & ";"
    Set oCatalog = Nothing

    With oTempConn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cTempPath & App.EXEName & ".MDB"
    '    oTempConn.ConnectionString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" & cTempPath
        .Open
    End With
    
    DetectTemp = True
    
    Exit Function
    
ErrDetect:
    MsgBox "Error retrieving temporary file", vbCritical
    DetectTemp = False
End Function


Public Function DetectServer() As Boolean
    On Error GoTo EndConnect
    
'    Add2List "Please wait trying to connect to server named " & cODBC
    objdbConn.CursorLocation = adUseClient
    objdbConn.ConnectionString = cConnString
'    MsgBox cConnString
    'objdbConn.ConnectionString = Provider=MSDASQL.1;Persist Security Info=False;User ID=micostaff;Data Source=mico;Password=123
    objdbConn.Open
    
    DetectServer = True
    Exit Function
    
EndConnect:
    MsgBox "Please check your connection string!", vbCritical, App.Title
    DetectServer = False
End Function


' for system menu and access right use
Public Function GetMenu(oForm As Form, Optional ByVal lReset As Boolean = False)

    On Error Resume Next

    Dim cString As String, cSqlStmt As String, cParent As String, _
        nCtr As Integer
    
    If lReset Then
        OpenQueryDNS "DELETE FROM PA6368", objdbRs, True
        OpenQueryDNS "DELETE FROM PA7668", objdbRs, True
        
        ' --> 20050317
        Log2Audit oForm.Name, "Resetting of Menu..."
        Script2File "DELETE FROM PA6368"
        Script2File "DELETE FROM PA7668"
        
        ' --> addendum 20050321
        ShowProgress 0
    End If
    
    For Each oControl In oForm
'        If nCtr > 100 Then nCtr = 0
        nCtr = nCtr + 1
'        If nCtr > 100 Then nCtr = 1
        
        If lReset Then ShowProgress 2, ((nCtr / 150) * 100)
        
'        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
        
'        If lReset Then      ' --> addendum 20050321
'            frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + 1
'        End If
        
        If (TypeName(oControl) = "Menu") Then
        
            If (oControl.Caption <> "-") And _
                (UCase(oControl.Name) <> "MNUSYSTEM") And _
                (Val(oControl.HelpContextID) <> 0) Then
                
                cString = ConcatStr(oControl.Caption, "&")
                
                If lReset Then ShowProgress 2, (nCtr / 150) * 100, , , "Setting status of " & cString
                
                If oControl.Visible Then
                
                    ' --> let's add 2 MENU table if not existing...
                    If Not IfExists("PA6368", "MNUNAME=" & cQuote & oControl.Name & cQuote) Then
                        cSqlStmt = "INSERT INTO PA6368(PA6368.MNUNAME,PA6368.CAPTION,PA6368.MENUID)VALUES(" & _
                                   cQuote & oControl.Name & cQuote & "," & _
                                   cQuote & cString & cQuote & "," & _
                                   oControl.HelpContextID & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt    ' --> 20050317
                    End If
                    
                    ' let's add 2 SYSMENU table if menu is not existing...
                    If Not IfExists("PA7668", "MNUNAME=" & cQuote & oControl.Name & cQuote) Then
                        ' get parent menu here...
                        If Mid(Trim(Str(oControl.HelpContextID)), 2, 1) = "0" Then
                            cParent = "PATMS"
                        Else
                            If Val(Mid(Trim(Str(oControl.HelpContextID)), 3, 2)) = 0 Then
                                OpenQueryDNS "SELECT * FROM PA6368 WHERE MENUID=" & left(Trim(Str(oControl.HelpContextID)), 1) + "000", objdbRs, False
                            Else
                                OpenQueryDNS "SELECT * FROM PA6368 WHERE MENUID=" & left(Trim(Str(oControl.HelpContextID)), 2) + "00", objdbRs, False
                            End If
                            
                            If objdbRs.RecordCount > 0 Then
                                cParent = objdbRs("MNUNAME")
                            End If
                        End If
                        
                        cSqlStmt = "INSERT INTO PA7668(PA7668.MNUNAME,PA7668.PARENT,PA7668.CAPTION,PA7668.AVAIL)VALUES(" & _
                                   cQuote & oControl.Name & cQuote & "," & _
                                   cQuote & cParent & cQuote & "," & _
                                   cQuote & cString & cQuote & ",1)"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        
                        ' --> 20050317
                        Log2Audit oForm.Name, "Add " & oControl.Name & " to system menu..."
                        Script2File cSqlStmt
                    End If
                    
                    OpenQueryDNS "SELECT * FROM PA7668 WHERE MNUNAME=" & cQuote & oControl.Name & cQuote, objdbRs, False
                    ' set visibility of d menu here...
                    oControl.Visible = (objdbRs.RecordCount > 0) And (objdbRs("AVAIL") = 1)
                    
                    'MsgBox oControl.Caption & Chr$(13) & Chr$(10) & Trim(Str(oControl.HelpContextID)) & Chr$(13) & Chr$(10) & cParent
                
                End If
                
            End If
            
        End If
        
    Next
    
    If lReset Then ShowProgress 4       ' --> added 20050321
End Function

'Public Function GetMenu(oForm As Form, Optional ByVal lReset As Boolean = False)
'
'    Dim cString As String, cSqlStmt As String, cParent As String, _
'        nCtr As Integer
'
'    If lReset Then
'        OpenQueryDNS "DELETE FROM PA6368", objdbRs, True
'        OpenQueryDNS "DELETE FROM PA7668", objdbRs, True
'
'        ' --> 20050317
'        Log2Audit oForm.Name, "Resetting of Menu..."
'        Script2File "DELETE FROM PA6368"
'        Script2File "DELETE FROM PA7668"
'
'        ' --> addendum 20050321
'        ShowProgress 0
'    End If
'
'    For Each oControl In oForm
''        If nCtr > 100 Then nCtr = 0
'        nCtr = nCtr + 1
''        If nCtr > 100 Then nCtr = 1
'
'        If lReset Then ShowProgress 2, ((nCtr / 150) * 100)
'
''        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
'
''        If lReset Then      ' --> addendum 20050321
''            frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + 1
''        End If
'
'        If (TypeName(oControl) = "Menu") Then
'
'            If (oControl.Caption <> "-") And _
'                (UCase(oControl.Name) <> "MNUSYSTEM") And _
'                (Val(oControl.HelpContextID) <> 0) Then
'
'                cString = ConcatStr(oControl.Caption, "&")
'
'                If lReset Then ShowProgress 2, (nCtr / 150) * 100, , , "Setting status of " & cString
'
'                If oControl.Visible Then
'
'                    ' --> let's add 2 MENU table if not existing...
'                    If Not IfExists("PA6368", "MNUNAME=" & cQuote & oControl.Name & cQuote) Then
'                        cSqlStmt = "INSERT INTO PA6368(PA6368.MNUNAME,PA6368.CAPTION,PA6368.MENUID)VALUES(" & _
'                                   cQuote & oControl.Name & cQuote & "," & _
'                                   cQuote & cString & cQuote & "," & _
'                                   oControl.HelpContextID & ")"
'                        OpenQueryDNS cSqlStmt, objdbRs, True
'                        Script2File cSqlStmt    ' --> 20050317
'                    End If
'
'                    ' let's add 2 SYSMENU table if menu is not existing...
'                    If Not IfExists("PA7668", "MNUNAME=" & cQuote & oControl.Name & cQuote) Then
'                        ' get parent menu here...
'                        If Mid(Trim(Str(oControl.HelpContextID)), 2, 1) = "0" Then
'                            cParent = "PATMS"
'                        Else
'                            If Val(Mid(Trim(Str(oControl.HelpContextID)), 3, 2)) = 0 Then
'                                OpenQueryDNS "SELECT * FROM PA6368 WHERE MENUID=" & left(Trim(Str(oControl.HelpContextID)), 1) + "000", objdbRs, False
'                            Else
'                                OpenQueryDNS "SELECT * FROM PA6368 WHERE MENUID=" & left(Trim(Str(oControl.HelpContextID)), 2) + "00", objdbRs, False
'                            End If
'
'                            If objdbRs.RecordCount > 0 Then
'                                cParent = objdbRs("MNUNAME")
'                            End If
'                        End If
'
'                        cSqlStmt = "INSERT INTO PA7668(PA7668.MNUNAME,PA7668.PARENT,PA7668.CAPTION,PA7668.AVAIL)VALUES(" & _
'                                   cQuote & oControl.Name & cQuote & "," & _
'                                   cQuote & cParent & cQuote & "," & _
'                                   cQuote & cString & cQuote & ",1)"
'                        OpenQueryDNS cSqlStmt, objdbRs, True
'
'                        ' --> 20050317
'                        Log2Audit oForm.Name, "Add " & oControl.Name & " to system menu..."
'                        Script2File cSqlStmt
'                    End If
'
'                    OpenQueryDNS "SELECT * FROM PA7668 WHERE MNUNAME=" & cQuote & oControl.Name & cQuote, objdbRs, False
'                    ' set visibility of d menu here...
'                    oControl.Visible = (objdbRs.RecordCount > 0) And (objdbRs("AVAIL") = 1)
'
'                    'MsgBox oControl.Caption & Chr$(13) & Chr$(10) & Trim(Str(oControl.HelpContextID)) & Chr$(13) & Chr$(10) & cParent
'
'                End If
'
'            End If
'
'        End If
'
'    Next
'
'    If lReset Then ShowProgress 4       ' --> added 20050321
'End Function


' retrieve user rights on each menu...
Public Sub GetUserMenu(ByVal oForm As Form, ByVal cUserID As String)
    On Error Resume Next
    
    OpenQueryDNS "SELECT * FROM PA2798 WHERE USERID=" & cQuote & cUserID & cQuote, objdbRs, False
    
    DoEvents
    
    For Each oControl In oForm
    
        If (TypeName(oControl) = "Menu") Then
        
            If lSuperUser Then
                oControl.Visible = True
            Else
                If oControl.Enabled Then
                    objdbRs.Requery adAsyncFetch
                    objdbRs.Find "MNUNAME='" & PadStr(oControl.Name, " ", 100, PadRight) & "'"
                    oControl.Visible = Not objdbRs.EOF
                End If
            End If
        End If
        
    Next
    
End Sub


' get user's right...
Public Sub GetUserRights(ByVal cMenu As String, cUserID As String)
    nAccess_Tag = 0
    OpenQueryDNS "SELECT * FROM PA2798 WHERE USERID=" & cQuote & cUserID & cQuote & " AND MNUNAME=" & cQuote & cMenu & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        nAccess_Tag = 1000 + (objdbRs("BIT1") * 100) + (objdbRs("BIT2") * 10) + objdbRs("BIT3")
    End If
    If lSuperUser Then nAccess_Tag = 1111
End Sub


' to concat string, gamit pa lang sa menu item
Public Function ConcatStr(ByVal cString As String, Optional ByVal cSearch As String) As String
    Dim nPos As Long
    nPos = InStr(1, cString, cSearch, vbTextCompare)
    If (cString <> "") And (cSearch <> "") And (nPos <> 0) Then
        ConcatStr = left(cString, nPos - 1) & right(cString, Len(cString) - nPos)
    Else
        ConcatStr = cString
    End If
End Function


Public Function PadStr(ByVal cString As String, ByVal cPad As String, ByVal nLength As Integer, Optional ByVal nDirection As Integer = 0) As String
    Dim nCtr As Integer
    Dim cRetVal As String
    cString = Trim(cString)
    If Not ((nLength = 0) Or (cString = "") Or (cPad = "") Or (Len(cString) >= nLength)) Then
        For nCtr = 1 To nLength - Len(cString)
            cRetVal = cRetVal + cPad
        Next
        If nDirection = PadLeft Then
            PadStr = cRetVal & cString
        Else
            PadStr = cString & cRetVal
        End If
    Else
        PadStr = cString
    End If
End Function


' for enter key to act as tab key...
Public Sub K_Press(ByVal KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


' add '\' in a defined path
Public Function CheckPath(cPath)
'    MsgBox cPath
    CheckPath = Trim(cPath) & IIf(right(Trim(cPath), 1) = "\", "", "\")
End Function


Public Sub OpenQueryDNS(ByVal cSqlStmt As String, oADORSet As ADODB.Recordset, ByVal lState As Boolean)
On Error GoTo ErrQuery
    
    DoEvents
    If Not lState Then
        Set oADORSet = objdbConn.Execute(cSqlStmt)
    Else
        objdbConn.Execute (cSqlStmt)
        While objdbConn.State = adStateExecuting
            DoEvents
        Wend
    End If
    
    Exit Sub
    
ErrQuery:
    ErrorMsg Err.Number, Err.Description, "Open Query", "uSRM"
End Sub

Public Sub QueryTemp(ByVal cSqlStmt As String, oADORSet As ADODB.Recordset, ByVal lState As Boolean)
    On Error GoTo ErrQuery
    
    DoEvents
    If Not lState Then
        Set oADORSet = oTempConn.Execute(cSqlStmt)
    Else
        oTempConn.Execute (cSqlStmt)
        While oTempConn.State = adStateExecuting
            DoEvents
        Wend
    End If
    Exit Sub
    
ErrQuery:
    ErrorMsg Err.Number, Err.Description, "Open Temporary Query", "uSRM"
End Sub

Public Sub QuerySSS(ByVal cSqlStmt As String, oADORSet As ADODB.Recordset, ByVal lState As Boolean)
    On Error GoTo ErrQuery
    
    DoEvents
    If Not lState Then
        Set oADORSet = oSSSConn.Execute(cSqlStmt)
    Else
        oSSSConn.Execute (cSqlStmt)
        While oSSSConn.State = adStateExecuting
            DoEvents
        Wend
    End If
    Exit Sub
    
ErrQuery:
    ErrorMsg Err.Number, Err.Description, "Open SSS Query", "uSRM"
End Sub

Public Sub ClearAll(ByVal oForm As Form, ByVal lEnabled As Boolean, ByVal lClear As Boolean)
    For Each oControl In oForm
        If oControl.Tag = 1 Then
            Select Case TypeName(oControl)
                Case "TextBox"
                    oControl.Text = IIf(lClear, "", oControl.Text)
                Case "ComboBox"
                    If lClear Then oControl.ListIndex = -1
                Case "DTPicker"
                    If lClear Then oControl.Value = Now
                Case "XPDatePicker"
                    If lClear Then oControl.CurrentDate = Now
                Case "CheckBox", "XPSpin"
                    If lClear Then oControl.Value = 0
            End Select
            oControl.Enabled = lEnabled
        End If
    Next
End Sub


' get Field value assigned on each control...
Public Function GetFields(ByVal oForm As Form, oADOTemp As ADODB.Recordset, Optional ByVal cTable As String)
    Dim cType As String
    Dim cField As String
    
    DoEvents
    
    For Each oControl In oForm
    
        If oControl.Tag = 1 Then
            If Trim(oControl.ToolTipText) <> "" Then
                cType = UCase(left(oControl.ToolTipText, InStr(1, oControl.ToolTipText, ":") - 1))
                cField = right(oControl.ToolTipText, Len(oControl.ToolTipText) - InStr(1, oControl.ToolTipText, ":"))
                
                Select Case TypeName(oControl)
                
                    Case "TextBox"
                        If oADOTemp.Fields.Item(cField).Type = adCurrency Then
                            oControl.MaxLength = 15   ' 27 jun 2005
                        Else
                            oControl.MaxLength = oADOTemp.Fields.Item(cField).DefinedSize   ' 03 jan 2005
                        End If
                        If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                            oControl.Text = ""
                        Else
                            If cType = "TXT" Then
                                oControl.Text = IIf(cField = "PASSWORD", oADOTemp(cField), DecodeStr(oADOTemp(cField)))
                            Else
    '                            oControl.Alignment = vbAlignRight
                                oControl.Text = Format(IIf(Trim(oADOTemp(cField)) = "", "0", oADOTemp(cField)), "#,##0.##00;-#,##0.##00")
                            End If
                        End If
                        
                    Case "ComboBox"
                        If cType = "TXT" Then
                            If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                                oControl.ListIndex = -1
                            Else
                                MatchCombo oADOTemp(cField), oControl
                            End If
                        Else
                            If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                                oControl.ListIndex = -1
                            Else
                                oControl.ListIndex = Val(oADOTemp(cField))
                            End If
                        End If
                        
                    Case "DTPicker"
                        If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                            oControl.Value = Now
                        Else
                            oControl.Value = IIf(IsDate(oADOTemp(cField)), oADOTemp(cField), Now)
                        End If
                        
                    Case "XPDatePicker"
                        If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                            oControl.CurrentDate = Now
                        Else
                            oControl.CurrentDate = IIf(IsDate(oADOTemp(cField)), oADOTemp(cField), Now)
                        End If
                        
                    Case "CheckBox", "XPSpin"
                        If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                            oControl.Value = 0
                        Else
                            oControl.Value = oADOTemp(cField)
                        End If
                        
                    Case "OptionButton"
                        If IsNull(oADOTemp(cField)) Or (oADOTemp.RecordCount = 0) Then
                            oControl.Value = False
                        Else
                            oControl.Value = oADOTemp(cField) = 1
                        End If
                        
                End Select
            End If
        End If
        
        ' --> disable Control Panel Button if recordcount is empty
        If oADOTemp.RecordCount = 0 Then
            If (oControl.Tag = 18) Or _
               (oControl.Tag = 19) Or _
               ((oControl.Tag >= 11) And (oControl.Tag <= 16)) Or _
               (oControl.Tag = 22) Or _
               (oControl.Tag = 23) Then
                oControl.Enabled = False
            End If
        End If
    Next
End Function


' create insert query...
Public Function InsertFields(ByVal oForm As Form, ByVal cTableName As String) As String
    Dim cType As String
    Dim cField As String
    Dim cCollectionValue As String
    Dim cCollectionField As String
    
    For Each oControl In oForm
    
        If oControl.Tag = 1 Then
            If Trim(oControl.ToolTipText) <> "" Then
                cType = left(oControl.ToolTipText, InStr(1, oControl.ToolTipText, ":") - 1)
                cField = cTableName & "." & right(oControl.ToolTipText, Len(oControl.ToolTipText) - InStr(1, oControl.ToolTipText, ":"))
                cCollectionField = cCollectionField & cField & ","
                        
                Select Case TypeName(oControl)
                
                    Case "TextBox"
                        If cType = "TXT" Then
                            cCollectionValue = cCollectionValue & cQuote & EncodeStr(oControl.Text) & cQuote & ","
                        Else
                            cCollectionValue = cCollectionValue & Format(IIf(Trim(oControl.Text) = "", "0", oControl.Text), "###.##00") & ","
                        End If
                        
                    Case "ComboBox"
                        If cType = "TXT" Then
                            cCollectionValue = cCollectionValue & cQuote & GetCombo(oControl.Text) & cQuote & ","
                        Else
                            cCollectionValue = cCollectionValue & oControl.ListIndex & ","
                        End If
                        
                    Case "DTPicker"
                        If cType = "TIM" Then
                            cCollectionValue = cCollectionValue & cQuote & Format(oControl.Value, "HH:MM:SS") & cQuote & ","
                        Else
                            cCollectionValue = cCollectionValue & cQuote & Format(oControl.Value, "yyyy-mm-dd") & cQuote & ","
                        End If
                        
                    Case "XPDatePicker"
                        cCollectionValue = cCollectionValue & cQuote & Format(oControl.CurrentDate, "yyyy-mm-dd") & cQuote & ","
                        
                    Case "CheckBox", "XPSpin"
                        cCollectionValue = cCollectionValue & oControl.Value & ","
                
                    Case "OptionButton"
                        cCollectionValue = cCollectionValue & IIf(oControl.Value = True, 1, 0) & ","
                
                End Select
            End If
        End If
        
    Next
    
    InsertFields = "INSERT INTO " & cTableName & "(" & _
                   left(cCollectionField, Len(cCollectionField) - 1) & ")VALUES(" & _
                   left(cCollectionValue, Len(cCollectionValue) - 1) & ")"
End Function


' create update query...
Public Function EditField(ByVal oForm As Form, ByVal cTableName As String, ByVal cCondition As String) As String
    Dim cType As String
    Dim cField As String
    Dim cCollectionField As String

    For Each oControl In oForm
    
        If oControl.Tag = 1 Then
            If Trim(oControl.ToolTipText) <> "" Then
                cType = left(oControl.ToolTipText, InStr(1, oControl.ToolTipText, ":") - 1)
                cField = cTableName & "." & right(oControl.ToolTipText, Len(oControl.ToolTipText) - InStr(1, oControl.ToolTipText, ":"))
                        
                Select Case TypeName(oControl)
                
                    Case "TextBox"
                        If cType = "TXT" Then
                            cCollectionField = cCollectionField & cField & " = " & cQuote & EncodeStr(oControl.Text) & cQuote & " ,"
                        Else
                            cCollectionField = cCollectionField & cField & " = " & Format(IIf(Trim(oControl.Text) = "", "0", oControl.Text), "###.##00") & " ,"
                        End If
                        
                    Case "ComboBox"
                        If cType = "TXT" Then
                            cCollectionField = cCollectionField & cField & " = " & cQuote & GetCombo(oControl.Text) & cQuote & " ,"
                        Else
                            cCollectionField = cCollectionField & cField & " = " & oControl.ListIndex & " ,"
                        End If
                        
                    Case "DTPicker"
                        If cType = "TIM" Then
                            cCollectionField = cCollectionField & cField & " = " & cQuote & Format(oControl.Value, "HH:MM:SS") & cQuote & ","
                        Else
                            cCollectionField = cCollectionField & cField & " = " & cQuote & Format(oControl.Value, "yyyy-mm-dd") & cQuote & ","
                        End If
                        
                    Case "XPDatePicker"
                        If cType = "TIM" Then
                            cCollectionField = cCollectionField & cField & " = " & cQuote & Format(oControl.CurrentDate, "HH:MM:SS") & cQuote & ","
                        Else
                            cCollectionField = cCollectionField & cField & " = " & cQuote & Format(oControl.CurrentDate, "yyyy-mm-dd") & cQuote & ","
                        End If
                        
                    Case "CheckBox", "XPSpin"
                        cCollectionField = cCollectionField & cField & " = " & oControl.Value & " ,"
                
                    Case "OptionButton"
                        cCollectionField = cCollectionField & cField & " = " & IIf(oControl.Value = True, 1, 0) & " ,"
                
                End Select
            End If
        End If
        
    Next
    
    EditField = "UPDATE " & cTableName & " SET " & left(cCollectionField, Len(cCollectionField) - 1) & _
                "WHERE " & cCondition
End Function


' for Error Messages...
Function ErrorMsg(ErrNum As Long, ErrDesc As String, _
    strFunction As String, strModule As String)
    On Error Resume Next
    Dim anErrorMessage As String
    anErrorMessage = "Error Number      : " & ErrNum & "." & vbCrLf & _
                     "Error Description : " & ErrDesc & vbCrLf & _
                     "Time Occured      : " & Format(Now, "hh:mm:ss") & vbCrLf & _
                     "Module Name       : " & strModule & vbCrLf & _
                     "Sub/Function      : " & strFunction & vbCrLf
        
    MsgBox anErrorMessage, vbCritical
    
    Write2File anErrorMessage
    
End Function


' --> QueryAttach Procedure
'   where:
'       oRecordSet      RecordSet to pass
'       oFlexGrid       MSHFlexGrid to display
'       cHeader         array containing information about the flexgrid header
'       lProgress       option to display progressbar, default to True
'       lAutoNumber     option to display number on 1st column, default to False
'       lWithHiLyt      option to highlight alternate row for display enhancement only, default to False
'       nMode           option to use fast attachment, default to 0 - standard
'   sample:
'   QueryAttach objDbRS, MSHFlexGrid1, myArray, False, , True, 1

Public Sub QueryAttach(ByVal oRecordSet As ADODB.Recordset, _
                       ByVal oFlexGrid As MSHFlexGrid, _
                       ByVal cHeader As Variant, _
                       Optional ByVal lProgress As Boolean = True, _
                       Optional ByVal lAutoNumber As Boolean = False, _
                       Optional ByVal lWithHilyt As Boolean = False, _
                       Optional ByVal nMode As Integer = 0)
    Dim nCtr, nRecNo, nStart, nEnd As Integer, _
        cValue As String, nValue As Double

    If lProgress Then ShowProgress 0
    
    If nMode = 0 Then
    
        SetGridColumn cHeader, oFlexGrid, lAutoNumber
        
        DoEvents
        
        With oFlexGrid
        
            .Redraw = False
            .Rows = oRecordSet.RecordCount + 1
            
            DoEvents
                
            While Not oRecordSet.EOF
            
                nRecNo = nRecNo + 1
                
                If lProgress Then
                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving " & Trim(Str(nRecNo)) & " of " & Trim(Str(oRecordSet.RecordCount)) & " item(s)..."
                End If
                
                If lWithHilyt Then
                    HiLyt nRecNo, Int(nRecNo / 2) = nRecNo / 2, oFlexGrid, , IIf(Int(nRecNo / 2) = nRecNo / 2, &HE0E0E0, vbWhite)
                End If
                
                .RowHeight(nRecNo) = 285
                For nCtr = 1 To oRecordSet.Fields.Count
                    cValue = Replace(cHeader(nCtr - 1), ",", ".", 1, Len(cHeader(nCtr - 1)), vbTextCompare)
                    nStart = InStr(1, cValue, "]") + 2
                    nEnd = InStrRev(cValue, ":") - 1
                    
                    nValue = Val(Mid(cValue, nStart, nEnd - nStart + 1))
                    
                    If lAutoNumber Then .TextMatrix(nRecNo, 0) = nRecNo
                    Select Case oRecordSet.Fields.Item(nCtr - 1).Type
                        Case adChar, adVarChar, adVarWChar, adWChar, adDBDate '202, 130  '133
                            If IsNull(Trim(oRecordSet(oRecordSet.Fields.Item(nCtr - 1).Name))) Then
                                .TextMatrix(nRecNo, nCtr) = ""
                            Else
                                .TextMatrix(nRecNo, nCtr) = Trim(DecodeStr(oRecordSet(oRecordSet.Fields.Item(nCtr - 1).Name)))
                            End If
                        
                        Case adSmallInt, adDouble, adCurrency, adInteger, adSmallInt, adBigInt, adNumeric '131
                            If Int(nValue) = nValue Then
                                .TextMatrix(nRecNo, nCtr) = oRecordSet(oRecordSet.Fields.Item(nCtr - 1).Name)
                            Else
                                .TextMatrix(nRecNo, nCtr) = FormatNumber(oRecordSet(oRecordSet.Fields.Item(nCtr - 1).Name), (nValue - Int(nValue)) * 10, , , vbFalse)
                            End If
                    End Select
                Next nCtr
                
                oRecordSet.MoveNext
                
            Wend
            
            .Redraw = True
            
        End With
        
    Else
    
        oFlexGrid.Clear
        Set oFlexGrid.Recordset = oRecordSet
        SetGridColumn cHeader, oFlexGrid, lAutoNumber, False
        If lWithHilyt Then RefreshGrid oFlexGrid, True
'        oFlexGrid.Redraw = True
        
    End If
    
    If lProgress Then ShowProgress 4
End Sub


Public Sub SetGridColumn(cColumnDef As Variant, oFlexGrid As MSHFlexGrid, Optional ByVal lAutoNumber As Boolean = False, Optional lAutoClear As Boolean = True)
    Dim nCtr As Integer, nStart As Integer, nEnd As Integer
    Dim cValue As String
    
    With oFlexGrid
        If lAutoClear Then
            .Clear
            .Rows = 2
        End If
        
        .Redraw = False
        
        If UBound(cColumnDef) >= 0 Then
            .Cols = UBound(cColumnDef) + 2
            .ColWidth(0) = IIf(lAutoNumber, 400, 300)
            
            For nCtr = 0 To UBound(cColumnDef)
                cValue = cColumnDef(nCtr)
                .TextMatrix(0, nCtr + 1) = Mid(cValue, InStr(1, cValue, "[") + 1, InStr(1, cValue, "]") - InStr(1, cValue, "[") - 1)
                
                nStart = InStr(1, cValue, "]") + 2
                nEnd = InStrRev(cValue, ":") - 1
                
                If UCase(right(cValue, Len(cValue) - InStrRev(cValue, ":"))) = "TRUE" Then
                    '.ColWidth(nCtr + 1) = Val(Mid(cValue, nStart, nEnd - nStart + 1)) * 90 '105
                    .ColWidth(nCtr + 1) = Val(Mid(cValue, nStart, nEnd - nStart + 1)) * (.Font.Size * 10.5)
                Else
                    .ColWidth(nCtr + 1) = 0
                End If
                
                Select Case left(cValue, InStr(1, cValue, ":") - 1)
                    Case "TXT"
                        .ColAlignment(nCtr + 1) = flexAlignLeftCenter
                    Case "NUM"
                        .ColAlignment(nCtr + 1) = flexAlignRightCenter
                End Select
            Next nCtr
            
        End If
        
        .RowHeight(1) = 285
        .Redraw = True
    End With
End Sub


' --> check for duplicate unique code in a string grid
Public Function ChkDupInGrid(ByVal cString As String, ByVal nPos As Integer, ByVal oFlexGrid As MSHFlexGrid) As Boolean
    Dim lFound As Boolean
    With oFlexGrid
        .Redraw = False
        .Row = 1
        lFound = False
        Do While .Row < .Rows - 1 And Not lFound
                 If UCase(left(.TextMatrix(.Row, nPos), Len(Trim(cString)))) = UCase(Trim(cString)) Then
                    lFound = True
'                    Exit Do
                 End If
                 .Row = .Row + 1
        Loop
        .Redraw = True
        ChkDupInGrid = lFound
    End With
End Function

' --> created 20050218
Public Function TableExist(ByVal cTable As String) As Boolean
    OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & cTable & cQuote, objdbRs, False
    TableExist = objdbRs.RecordCount > 0
End Function


Public Function IfExists(ByVal cTable As String, ByVal cCondition As String) As Boolean
    OpenQueryDNS "SELECT * FROM " & cTable & " WHERE " & cCondition, objdbRs, False
    IfExists = objdbRs.RecordCount > 0
End Function


' used for auto-generation of field ID...
Function GenerateSeries(ByVal cCounterID As String, Optional ByVal nPadLen As Long, Optional ByVal cWSID As String = "", Optional ByVal cTmpCMPID As String = "") As String
    Dim nPos As Long, cSqlStmt As String
    
    cCounterID = Trim(UCase(cCounterID))
    
    If Not IfExists("PA7287", "CTRID=" & cQuote & cCounterID & cQuote & IIf(Trim(cWSID) = "", "", " AND WSID=" & cQuote & cWSID & cQuote)) Then
        cSqlStmt = "INSERT INTO PA7287(CTRID,COUNTER,UNUSED,WSID)VALUES(" & cQuote & cCounterID & cQuote & ",0," & cQuote & cQuote & "," & cQuote & cWSID & cQuote & ")"
        OpenQueryDNS cSqlStmt, objdbRs, True
'        Script2File cSqlStmt
    End If
        
    OpenQueryDNS "SELECT COUNTER, UNUSED FROM PA7287 WHERE CTRID=" & cQuote & cCounterID & cQuote & IIf(Trim(cWSID) = "", "", " AND WSID=" & cQuote & cWSID & cQuote), objdbRs, False
    If Trim(objdbRs("UNUSED")) <> "" Then
        nPos = InStrRev(objdbRs("UNUSED"), ",")
        If nPos > 0 Then
            GenerateSeries = right(objdbRs("UNUSED"), Len(objdbRs("UNUSED")) - nPos - 1)
            cSqlStmt = "UPDATE PA7287 SET UNUSED=" & cQuote & left(objdbRs("UNUSED"), nPos - 1) & cQuote & _
                       " WHERE CTRID=" & cQuote & cCounterID & cQuote & IIf(Trim(cWSID) = "", "", " AND WSID=" & cQuote & cWSID & cQuote)
        Else
            GenerateSeries = objdbRs("UNUSED")
            cSqlStmt = "UPDATE PA7287 SET UNUSED=''" & _
                       " WHERE CTRID=" & cQuote & cCounterID & cQuote & IIf(Trim(cWSID) = "", "", " AND WSID=" & cQuote & cWSID & cQuote)
        End If
        OpenQueryDNS cSqlStmt, objdbRs, True
'        Script2File cSqlStmt
    Else
        GenerateSeries = Trim(Str(objdbRs("COUNTER") + 1))
        cSqlStmt = "UPDATE PA7287 SET COUNTER=COUNTER+1 WHERE CTRID=" & cQuote & cCounterID & cQuote & IIf(Trim(cWSID) = "", "", " AND WSID=" & cQuote & cWSID & cQuote)
        OpenQueryDNS cSqlStmt, objdbRs, True
'        Script2File cSqlStmt
    End If
End Function


' used to reset generated series...
Public Sub ResetSeries(ByVal cCounterID As String, ByVal cGeneratedCode As String)
    Dim oADORS As New ADODB.Recordset
    cCounterID = Trim(UCase(cCounterID))
    
    If Trim(cGeneratedCode) <> "" Then
        If IfExists("PA7287", "CTRID=" & cQuote & cCounterID & cQuote) Then
            OpenQueryDNS "SELECT COUNTER,UNUSED FROM PA7287 WHERE CTRID=" & cQuote & cCounterID & cQuote, oADORS, False
            
            If Val(oADORS("COUNTER")) = Val(cGeneratedCode) Then
            
                OpenQueryDNS "UPDATE PA7287 SET COUNTER=COUNTER-1 WHERE CTRID=" & cQuote & cCounterID & cQuote, oADORS, True
                
'                Script2File "UPDATE SCOUNTER SET COUNTER=COUNTER-1 WHERE CTRID=" & cQuote & cCounterID & cQuote
                
            Else
                If Trim(oADORS("UNUSED")) <> "" Then
                    If InStr(1, oADORS("UNUSED"), "," & cGeneratedCode) <> 0 Then
                        OpenQueryDNS "UPDATE PA7287 SET UNUSED=" & cQuote & _
                                     oADORS("UNUSED") & "," & Trim(Str(Val(cGeneratedCode))) & cQuote & _
                                     " WHERE CTRID=" & cQuote & cCounterID & cQuote, oADORS, True
                                     
'                        Script2File "UPDATE SCOUNTER SET UNUSED=" & cQuote & _
'                                    oADORS("UNUSED") & "," & Trim(Str(Val(cGeneratedCode))) & cQuote & _
'                                    " WHERE CTRID=" & cQuote & cCounterID & cQuote
                    End If
                Else
                    OpenQueryDNS "UPDATE PA7287 SET UNUSED=" & cQuote & Trim(Str(Val(cGeneratedCode))) & cQuote & _
                                 " WHERE CTRID=" & cQuote & cCounterID & cQuote, oADORS, True
                                 
'                    Script2File "UPDATE SCOUNTER SET UNUSED=" & cQuote & Trim(Str(Val(cGeneratedCode))) & cQuote & _
'                                " WHERE CTRID=" & cQuote & cCounterID & cQuote
                End If
                
            End If
            
        End If
    End If
    
    Set oADORS = Nothing
End Sub


' User-Defined Function for Combo Box
' LoadCombo
' MatchCombo
' GetCombo

' load item on a combo box...
Public Sub LoadCombo(ByVal cField1 As String, ByVal cField2 As String, oADOSource As ADODB.Recordset, ByVal oComboBox As ComboBox)
    oComboBox.Clear
    If oADOSource.RecordCount > 0 Then
        oADOSource.MoveFirst
        While Not oADOSource.EOF
            oComboBox.AddItem (DecodeStr(Trim(oADOSource(cField2))) & ", " & Trim(oADOSource(cField1)))
            oADOSource.MoveNext
        Wend
    End If
End Sub

' match an item in a combobox...
Public Sub MatchCombo(ByVal cString As String, oComboBox As ComboBox)
    Dim nCtr As Integer, cValue As String
    oComboBox.ListIndex = -1
    If Trim(cString) <> "" Then
        For nCtr = 0 To oComboBox.ListCount - 1
            cValue = Trim(right(oComboBox.List(nCtr), Len(oComboBox.List(nCtr)) - InStrRev(oComboBox.List(nCtr), ",")))
'            If InStrRev(oComboBox.List(nCtr), cString) > 0 Then
            If UCase(cValue) = UCase(cString) Then
                oComboBox.ListIndex = nCtr
                Exit For
            End If
        Next
    End If
End Sub

' get combo code in a combobox
Public Function GetCombo(ByVal cComboString As String) As String
    If Trim(cComboString) = "" Then
        GetCombo = cComboString
    Else
        GetCombo = right(cComboString, Len(cComboString) - InStrRev(cComboString, ",") - 1)
    End If
End Function


' replace quote/backslash in a string here...
Public Function EncodeStr(ByVal cString As String) As String
    'cString = IIf(IsNull(cString), "", cString)
    If Trim(cString) <> "" Then
        cString = Replace(cString, """", Chr(22), 1, Len(cString), vbTextCompare)
        cString = Replace(cString, "\", Chr(23), 1, Len(cString), vbTextCompare)
    End If
    EncodeStr = cString
End Function


Public Function EncodeStr2(ByVal cString As String) As String
    'cString = IIf(IsNull(cString), "", cString)
    If Trim(cString) <> "" Then
        cString = Replace(cString, """", cQuote & cQuote, 1, Len(cString), vbTextCompare)
    End If
    EncodeStr2 = cString
End Function


' decode encoded quote/backslash in a string...
Public Function DecodeStr(ByVal cString As String) As String
    'cString = IIf(IsNull(cString), "", cString)
    If Trim(cString) <> "" Then
        cString = Replace(cString, Chr(22), """", 1, Len(cString), vbTextCompare)
        cString = Replace(cString, Chr(23), "\", 1, Len(cString), vbTextCompare)
    End If
    DecodeStr = cString
End Function


Public Sub CtrlPanel(oForm As Form, ntag As Integer, Optional ByVal lEnabled As Boolean = True)
    For Each oControl In oForm
        If TypeName(oControl) = "CommandButton" And (Trim(oControl.Tag) <> "") Then
            oControl.Enabled = ntag = 0
        
            
            ' --> start of access right
            If Trim(oForm.Tag) <> "" Then
                If (oControl.Tag >= 17) And (oControl.Tag <= 19) Then
                    oControl.Enabled = (Mid(oForm.Tag, oControl.Tag - 15, 1) = "1") And (ntag = 0)
                End If
                
                If oControl.Tag = 20 Then
                    oControl.Enabled = IIf(ntag = 0, False, (Mid(oForm.Tag, 2, 1) = "1") Or (Mid(oForm.Tag, 3, 1) = "1"))
                End If
            End If
            ' --> end of access rights
            
            If oControl.Tag = 21 Then
                oControl.Caption = IIf(ntag = 0, "&Close", "&Cancel")
                oControl.Enabled = True
            End If
            
            If Not lEnabled Then
                If (oControl.Tag = 18) Or (oControl.Tag = 19) Or (oControl.Tag = 22) Or (oControl.Tag = 23) Then
                    oControl.Enabled = lEnabled
                End If
            End If

'            If (oControl.Tag = 22) Or (oControl.Tag = 23) Then
'                oControl.Enabled = lEnabled
'            End If
        End If
    Next
End Sub


Public Function addQuote(ByVal token) As String
    addQuote = Replace(token, "'", "''", 1)
End Function


Public Function enkryp(ByVal cValue As String, Optional ByVal cID As String = "") As String
    Dim cSqlStmt As String
    cSqlStmt = "SELECT AES_ENCRYPT(" & cQuote & EncodeStr(cValue) & cQuote & "," & cQuote & UCase(cID) & cQuote & ") AS SRM"
    OpenQueryDNS cSqlStmt, objdbRs, False
    enkryp = objdbRs("SRM")
End Function


Public Function Dekryp(ByVal cValue As String, Optional ByVal cID As String = "") As String
    Dim cSqlStmt As String
    cSqlStmt = "SELECT AES_DECRYPT(" & cQuote & cValue & cQuote & "," & cQuote & UCase(cID) & cQuote & ") AS SRM"
    OpenQueryDNS cSqlStmt, objdbRs, False
    Dekryp = IIf(objdbRs("SRM") = Null, "", DecodeStr(objdbRs("SRM")))
End Function


Public Sub dbNavigator(ByVal oButton As VB.CommandButton, ByVal oForm As Form, ByVal oRecordSet As ADODB.Recordset)
    If oRecordSet.RecordCount > 0 Then
        Select Case oButton.Tag
            Case 11     ' --> top
                oRecordSet.MoveFirst
            Case 12     ' --> bottom
                oRecordSet.MoveLast
            Case 13     ' --> previous
                oRecordSet.MovePrevious
                If oRecordSet.BOF Then oRecordSet.MoveFirst
            Case 14     ' --> next
                oRecordSet.MoveNext
                If oRecordSet.EOF Then oRecordSet.MoveLast
        End Select
        GetFields oForm, oRecordSet
    End If
End Sub


Public Sub Write2File(ByVal cString As String)
    Dim oTextFile As New FileSystemObject, _
        oFile As File, _
        oTxtStream As TextStream, _
        cLogFile As String
    
    cLogFile = CheckPath(cLogPath) & Format(Now, "yyyymmdd") & ".log"
    If Dir(cLogFile) = "" Then
        Set oTxtStream = oTextFile.CreateTextFile(cLogFile, True)
    Else
        Set oFile = oTextFile.GetFile(cLogFile)
        Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
    End If
    
    oTxtStream.WriteLine cString
    
    oTxtStream.Close
    
    Set oTxtStream = Nothing
    Set oTextFile = Nothing
    Set oFile = Nothing
End Sub


Public Sub Log2Audit(ByVal cModule As String, ByVal cActivity As String, Optional ByVal cMGR_CODE = "", Optional ByVal cMGR_NAME = "")
    On Error GoTo ErrLog
    
    Dim cSqlStmt As String, _
        oADOLog As New ADODB.Recordset
    
    DoEvents
    
    OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & "PA28348" & cQuote, oADOLog, False
    If oADOLog.RecordCount = 0 Then
        cSqlStmt = "CREATE TABLE PA28348(" & _
                   "    `WSID` CHAR(3) NOT NULL DEFAULT ''," & _
                   "    `USERID` CHAR(6) NOT NULL DEFAULT ''," & _
                   "    `USERNAME` CHAR(50) NOT NULL DEFAULT ''," & _
                   "    `DATE` DATE NOT NULL DEFAULT '" & Format(Now, "yyyy-mm-dd") & "'," & _
                   "    `TIME` CHAR(10) NOT NULL DEFAULT ''," & _
                   "    `MODULE` CHAR(50) NOT NULL DEFAULT ''," & _
                   "    `ACTIVITY` CHAR(250) NOT NULL DEFAULT ''," & _
                   "    `MGR_CODE` char(6) NOT NULL DEFAULT ''," & _
                   "    `MGR_NAME` char(50) NOT NULL DEFAULT ''," & _
                   "    FULLTEXT (`ACTIVITY`)" & _
                   ") Type=MyISAM;"
        OpenQueryDNS cSqlStmt, oADOLog, True
    End If
    
    OpenQueryDNS "SHOW COLUMNS FROM PA28348 LIKE 'MGR_CODE'", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cSqlStmt = "INSERT INTO PA28348(WSID,USERID,USERNAME,`DATE`,`TIME`,`MODULE`,`ACTIVITY`,`MGR_CODE`,`MGR_NAME`)VALUES(" & _
                   cQuote & gWSID & cQuote & "," & _
                   cQuote & gUserID & cQuote & "," & _
                   cQuote & gUserName & cQuote & "," & _
                   cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                   cQuote & Format(Now, "hh:mm:ss") & cQuote & "," & _
                   cQuote & cModule & cQuote & "," & _
                   cQuote & EncodeStr(cActivity) & cQuote & "," & _
                   cQuote & cMGR_CODE & cQuote & "," & _
                   cQuote & cMGR_NAME & cQuote & ")"
    Else
        cSqlStmt = "INSERT INTO PA28348(WSID,USERID,USERNAME,`DATE`,`TIME`,`MODULE`,`ACTIVITY`)VALUES(" & _
                   cQuote & gWSID & cQuote & "," & _
                   cQuote & gUserID & cQuote & "," & _
                   cQuote & gUserName & cQuote & "," & _
                   cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                   cQuote & Format(Now, "hh:mm:ss") & cQuote & "," & _
                   cQuote & cModule & cQuote & "," & _
                   cQuote & EncodeStr(cActivity) & cQuote & ")"
    End If
    OpenQueryDNS cSqlStmt, oADOLog, True
    Script2File cSqlStmt
ErrLog:
    Set oADOLog = Nothing
End Sub

Public Sub Script2File(ByVal cString As String)
    Dim oTextFile As New FileSystemObject, _
        oTxtStream As TextStream, _
        oFile As File, _
        cScriptFile As String, _
        oScriptRSet As New ADODB.Recordset, _
        cSqlStmt As String

    DoEvents
    
    ' --> check table existence first...
    OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & "PA7287" & cQuote, objdbRs, False
    If objdbRs.RecordCount = 0 Then
        cSqlStmt = "CREATE TABLE `PA7287` (`ctrID` char(10) NOT NULL default ''," & _
                   "   `counter` int(5) NOT NULL default '0'," & _
                   "   `unused` char(200) NOT NULL default ''," & _
                   "   `wsID` char(4) NOT NULL default ''," & _
                   "   `CMPID` char(4) NOT NULL default '0003'," & _
                   "   PRIMARY KEY  (`ctrID`,`wsID`,`CMPID`)" & _
                   " ) ENGINE=MyISAM DEFAULT CHARSET=latin1"
        OpenQueryDNS cSqlStmt, objdbRs, True
    End If
    
    ' --> check table existence first... (201703-02)
        OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & "DI36770A" & cQuote, objdbRs, False
    If objdbRs.RecordCount = 0 Then
        cSqlStmt = "CREATE TABLE  `DI36770A` (`EMPID` char(6) NOT NULL default ''," & _
                   "    `PERIODID` char(5) NOT NULL default ''," & _
                   "    `DATE` date NOT NULL default '2011-01-19'," & _
                   "    `SHIFTID` char(5) NOT NULL default ''," & _
                   "    `DESCRIPTION` char(100) NOT NULL default ''," & _
                   "    `TIME1` char(10) NOT NULL default ''," & _
                   "    `TIME2` char(10) NOT NULL default ''," & _
                   "    `allowance` double NOT NULL default '5'," & _
                   "    `reg_hr` double NOT NULL default '0'," & _
                   "    `reg_ot_hr` double NOT NULL default '0'," & _
                   "    `sa_reg_ot` double NOT NULL default '0'," & _
                   "    `tot_ot` double NOT NULL default '0'," & _
                   "    `nd_hr` double NOT NULL default '0'," & _
                   "    `nd_ot_hr` double NOT NULL default '0'," & _
                   "    `sa_nd_ot` double NOT NULL default '0'," & _
                   "    `nd_tot_ot` double NOT NULL default '0'," & _
                   "    `sun_hr` double NOT NULL default '0'," & _
                   "    `sun_ot_hr` double NOT NULL default '0'," & _
                   "    `sun_nd` double NOT NULL default '0'," & _
                   "    `sun_nd_ot` double NOT NULL default '0'," & _
                   "    `Inc_hr` double NOT NULL default '0'," & _
                   "    `REMARK` char(50) NOT NULL default ''," & _
                   "    `TAG` int(1) NOT NULL default '0'," & _
                   "    `CMPID` char(4) NOT NULL default '0019'," & _
                   "     PRIMARY KEY  (`EMPID`,`DATE`)" & " ) ENGINE=MyISAM DEFAULT CHARSET=latin1;"
                   
        OpenQueryDNS cSqlStmt, objdbRs, True
    End If
    
    
    ' --> check table existence first... (201704-07)
        OpenQueryDNS "SHOW TABLE STATUS LIKE " & cQuote & "DI7673" & cQuote, objdbRs, False
    If objdbRs.RecordCount = 0 Then
       cSqlStmt = "CREATE TABLE  `DI7673` (`POSID` char(3) NOT NULL default ''," & _
                  "    `DESIGNATION` int(1) NOT NULL default '0'," & _
                  "     PRIMARY KEY  (`POSID`)" & " ) ENGINE=MyISAM DEFAULT CHARSET=latin1;"
        OpenQueryDNS cSqlStmt, objdbRs, True
    End If


    
    If Dir(CheckPath(cScriptPath) & Format(Now, "yyyymmdd"), vbDirectory) = "" Then
        MkDir CheckPath(cScriptPath) & Format(Now, "yyyymmdd")
        OpenQueryDNS "UPDATE PA7287 SET COUNTER=1 WHERE CTRID='SCRIPT'" & IIf(Trim(gWSID) = "", "", " AND WSID=" & cQuote & gWSID & cQuote), objdbRs, True
    End If
            
    OpenQueryDNS "SELECT COUNTER FROM PA7287 WHERE CTRID='SCRIPT' AND WSID=" & cQuote & gWSID & cQuote, oScriptRSet, False
    
    If oScriptRSet.RecordCount > 0 Then
        cScriptFile = CheckPath(cScriptPath) & Format(Now, "yyyymmdd") & "\" & PadStr(oScriptRSet("COUNTER"), "0", 8) & ".SRM"
    Else
        cScriptFile = CheckPath(cScriptPath) & Format(Now, "yyyymmdd") & "\" & PadStr(GenerateSeries("SCRIPT", , gWSID), "0", 8) & ".SRM"
    End If
    
loopd2:
    If Dir(cScriptFile) = "" Then
        Set oTxtStream = oTextFile.CreateTextFile(cScriptFile, True)
    Else
        Set oFile = oTextFile.GetFile(cScriptFile)
        If oFile.Size > 1048576 Then    ' --> adjusted to 1MB file size - 20051212
            cScriptFile = CheckPath(cScriptPath) & Format(Now, "yyyymmdd") & "\" & PadStr(GenerateSeries("SCRIPT", , gWSID), "0", 8) & ".SRM"
            GoTo loopd2
        Else
            Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
        End If
    End If
    
    cString = cString & IIf(right(cString, 1) <> ";", ";", "")    ' --> 20051212
    
    oTxtStream.WriteLine cString
    oTxtStream.Close
    
    Set oScriptRSet = Nothing
    Set oTxtStream = Nothing
                            Set oTextFile = Nothing
    Set oFile = Nothing
End Sub
Public Sub ShowProgress(ByVal nMode As Integer, _
                        Optional ByVal nValue As Integer = 0, _
                        Optional ByVal nNoList As Integer = 0, _
                        Optional ByVal cTopTxt As String = "", _
                        Optional ByVal cBottomTxt As String = "")
                        
    With frmProgress
        Select Case nMode
            Case 0      ' --> initialize
                .Height = IIf(nNoList = 0, 1620, 3570)
                .Show
                .Label1.Caption = App.ProductName
                .Label2.Caption = cCompany & vbCrLf & gAddress
                .Label3.top = IIf(nNoList = 0, 990, 2955)
                .ProgressBar1.top = IIf(nNoList = 0, 1200, 3165)
                .ProgressBar1.Value = 0
                .ProgressBar1.Visible = True
                .List1.Visible = Not nNoList = 0
                .List1.Clear
                SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
                
            Case 1      ' --> reset progressbar
                .ProgressBar1.Visible = True
                .ProgressBar1.Value = 0
            
            Case 2      ' --> pass value of progressbar here
                If nStart = 0 Then nStart = (Hour(Time) * 3600) + (Minute(Time) * 60) + Second(Time)
                If nOldValue = 0 Then nOldValue = nValue

                .ProgressBar1.Value = nValue
                If .List1.Visible Then
                    .List1.AddItem cBottomTxt
                    .List1.ListIndex = .List1.ListCount - 1
                End If

                If nOldValue <> nValue Then
                    nOldValue = nValue
                    nCurTime = (Hour(Time) * 3600) + (Minute(Time) * 60) + Second(Time)
                    If nValue <> 0 Then
                        nDuration = ((nCurTime - nStart) * .ProgressBar1.Max) / nValue
                        nElapse = nDuration - (nCurTime - nStart)
                        If (nElapse - ((nElapse \ 5) * 5)) <= 0 Then
                            If (nElapse \ 60) > 0 Then
                                .Label3.Caption = Format((nElapse \ 60), "#0") & " minute(s) and " & Format((nElapse - ((nElapse \ 60) * 60)), "#0") & " second(s) remaining..."
                            Else
                                .Label3.Caption = Format(nElapse, "#0") & " second(s) remaining..."
                            End If
                        End If
                    End If
                End If
                
            Case 3      ' --> for reporting purposes...
                cBottomTxt = IIf(Trim(cBottomTxt) = "", "Preparing report...", cBottomTxt)
                .ProgressBar1.Visible = False
                .List1.Visible = False
                .Label3.Caption = cBottomTxt
                SetWindowPos .hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
            
            Case 4      ' --> unload frmProgress
                SetWindowPos .hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
                Unload frmProgress
            
        End Select
    End With
End Sub


Public Sub Lock2User(ByVal cModuleName As String, ByVal cField As String, cData As String, Optional ByVal isLocked As Boolean = False)
    On Error GoTo ErrLock
    Dim cSqlStmt As String
    
    cField = right(cField, Len(cField) - InStr(1, cField, ":"))
    
    If isLocked Then
        cSqlStmt = "INSERT INTO PA5625(WSID,USERID,`DATE`,`TIME`,`MODULE`,`FIELD`,`VALUE`)VALUES(" & _
                   cQuote & gWSID & cQuote & "," & _
                   cQuote & gUserID & cQuote & "," & _
                   cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                   cQuote & Format(Now, "hh:mm:ss") & cQuote & "," & _
                   cQuote & cModuleName & cQuote & "," & _
                   cQuote & cField & cQuote & "," & _
                   cQuote & cData & cQuote & ")"
    Else
        cSqlStmt = "DELETE FROM PA5625 WHERE USERID=" & cQuote & gUserID & cQuote & " AND " & _
                   "`MODULE`=" & cQuote & cModuleName & cQuote & " AND " & _
                   "`FIELD`=" & cQuote & cField & cQuote & " AND " & _
                   "`VALUE`=" & cQuote & cData & cQuote
    End If
    OpenQueryDNS cSqlStmt, objdbRs, True
    
    Exit Sub
    
ErrLock:
    ErrorMsg Err.Number, Err.Description, IIf(isLocked, "LOCK", "UNLOCK") & " MODULE", cModuleName
End Sub


Public Function isDataLock(ByVal cModuleName As String, ByVal cField As String, cData As String) As Boolean

    cField = right(cField, Len(cField) - InStr(1, cField, ":"))
    
    OpenQueryDNS "SELECT * FROM PA5625 WHERE " & _
                 "`MODULE`=" & cQuote & cModuleName & cQuote & " AND " & _
                 "`FIELD`=" & cQuote & cField & cQuote & " AND " & _
                 "`VALUE`=" & cQuote & cData & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If Not objdbRs("USERID") = gUserID Then
            MsgBox "This record is currently in exclusive edit mode by user " & Trim(objdbRs("USERID")) & "!", vbExclamation, App.Title
            isDataLock = True
        Else
            isDataLock = False
        End If
    Else
        isDataLock = False
    End If
End Function


Public Sub HiLyt(ByVal nRow As Integer, ByVal lHighlight As Boolean, ByVal oFlexGrid As MSHFlexGrid, Optional ByVal nColumn As Integer = 0, Optional ByVal nColor As Long = &H80000018)
    Dim nCol As Integer
    With oFlexGrid
        .Redraw = False
        
        nCol = .Col
        .Row = nRow
        .FillStyle = flexFillRepeat
        .Col = 1
        .ColSel = .Cols() - 1
        
        '.CellBackColor = IIf(lHighlight, nColor, &H80000005)
        If Not ((.CellBackColor = &H816DEB) Or (.CellBackColor = vbCyan)) Then
            If lHighlight Then
                .CellBackColor = nColor
            Else
                .CellBackColor = IIf(Int(.Row / 2) = .Row / 2, &HE0E0E0, vbWhite)
            End If
        End If
        
        .FillStyle = flexFillSingle
'        If nColor = vbInfoBackground Then .TextMatrix(nRow, 0) = IIf(lHighlight, Chr$(169), "")
        If nColor = vbInfoBackground Then .TextMatrix(nRow, 0) = IIf(lHighlight, Chr$(26), "")
        If nColumn > 0 Then .TextMatrix(nRow, nColumn) = IIf(lHighlight, "1", "0")
        .Col = nCol
        
        .Redraw = True
    End With
End Sub


Public Sub HiLyt2(ByVal nRowPos As Integer, ByVal oFlexGrid As MSHFlexGrid, Optional nColor As Long = vbBlack)
    Dim nOldRow As Integer, _
        lRedraw As Boolean
    With oFlexGrid
        lRedraw = .Redraw
        nOldRow = .Row
        If lRedraw Then .Redraw = False
        .Row = nRowPos
        .FillStyle = flexFillRepeat
        .Col = 1
        .ColSel = .Cols - 1
        .CellForeColor = nColor
        .FillStyle = flexFillSingle
        If lRedraw Then .Redraw = True
        .Row = nOldRow
    End With
End Sub


' --> function to check existing highlight in a mshflexgrid for deletion...
Public Function IsHilyt(ByVal oFlexGrid As MSHFlexGrid) As Boolean
    Dim nCtr As Integer, _
        nRow As Integer
        
    IsHilyt = False
    
    With oFlexGrid
    
        DoEvents
        .Redraw = False
        nRow = oFlexGrid.Row
        
        For nCtr = 1 To .Rows - 1
        
            .Row = nCtr
            
            If .CellBackColor = vbRed Then
                IsHilyt = True
                Exit For
            End If
            
        Next nCtr
        .Redraw = True
        .Row = nRow
    End With
End Function


' --> procedure to delete all highlighted cell in a mshflexgrid
Public Sub DelHilyt(ByVal oFlexGrid As MSHFlexGrid, Optional ByVal nColPos = 0, Optional ByVal cWarningMsg As String = "A link had been detected... Would you like to delete anyway?")
    Dim nCtr As Integer
        
    With oFlexGrid
        DoEvents
        .Redraw = False
        ShowProgress 0, , 100
        For nCtr = 1 To .Rows - 1
            If nCtr > .Rows - 1 Then Exit For
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            .Row = nCtr
ikot:
            If .CellBackColor = vbRed Then
                If (nColPos > 0) And (Trim(.TextMatrix(nCtr, nColPos)) <> "") Then
                    If MsgBox(cWarningMsg, vbYesNo, App.Title) = vbYes Then
                        If .Rows - 1 = 1 Then
                            .AddItem "", .Rows
                            .RowHeight(.RowSel + 1) = 285
                        End If
                        .RemoveItem nCtr
                        GoTo ikot   ' --> loop just 2 make sure...
                    End If
                Else
                    If .Rows - 1 = 1 Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                    End If
                    .RemoveItem nCtr
                    GoTo ikot   ' --> loop just 2 make sure...
                End If
            End If
        Next nCtr
        ShowProgress 4
        .Redraw = True
    End With
End Sub


Public Sub RefreshGrid(ByVal oFlexGrid As MSHFlexGrid, Optional ByVal lWithHilyt As Boolean = False, Optional ByVal lAutoNum As Boolean = False)
    Dim nOldRow, nRecNo As Integer
    
    With oFlexGrid
        nOldRow = .Row
        .Redraw = False
        DoEvents
        For nRecNo = 1 To .Rows - 1
            If lWithHilyt Then
                HiLyt nRecNo, Int(nRecNo / 2) = nRecNo / 2, oFlexGrid, , IIf(Int(nRecNo / 2) = nRecNo / 2, &HE0E0E0, vbWhite)
            End If
            If lAutoNum Then
                If .TextMatrix(nRecNo, 0) <> Chr(26) Then
                    .TextMatrix(nRecNo, 0) = nRecNo
                End If
            End If
        Next nRecNo
        .Row = nOldRow
        .Redraw = True
    End With
End Sub


' --> 20051121 - Copy & Paste row value from MSHFlexGrid to another row...
Public Function CopyFlexInfo(ByVal oFlexGrid As MSHFlexGrid, ByVal cColumnDef As Variant, Optional ByVal nMode As Integer = 1) As Variant
    Dim nCtr, nRow, nOldRow As Integer, _
        aFlexValue As Variant, _
        cValue As String
        
    With oFlexGrid
        .Redraw = False
        DoEvents
        
        If Not (nMode = 2) Then
            RefreshGrid oFlexGrid, True
            HiLyt oFlexGrid.Row, True, oFlexGrid, , IIf(nMode = 0, &H816DEB, vbCyan)
            aFlexValue = cColumnDef
        Else
            .Redraw = False
            DoEvents
            nOldRow = .Row
            For nRow = 1 To .Rows - 1
                .Row = nRow
                If .CellBackColor = &H816DEB Then Exit For
            Next nRow
            .Row = nOldRow
            .Redraw = True
        End If
        
        For nCtr = 1 To .Cols - 1
            If (nMode = 2) Then
                .TextMatrix(.Row, nCtr) = cColumnDef(nCtr - 1)
            Else
                aFlexValue(nCtr - 1) = .TextMatrix(.Row, nCtr)
            End If
        Next nCtr
        
        ' --> remove if masked is CUT...
        If (nMode = 2) And (nRow > 0) Then
            .RemoveItem nRow
            If nRow <= .Row Then .RowSel = .Row - 1
            RefreshGrid oFlexGrid, True
        End If
    
        .Redraw = True
        
        .SetFocus
        
    End With
    
    CopyFlexInfo = aFlexValue
End Function


Public Function GenerateReport(ByVal cReportTitle As String, ByVal cRptName As String, Optional ByVal cString As String, Optional ByVal lIsAccess As Boolean = False)
    Dim m_frmRptViewer As New frmRptViewer

On Error GoTo ErrGenerateRPT
    
    If Dir(cReportPath & cRptName) <> "" Then
        
        MsgBox "Press [OK] to continue...", vbInformation, "Preview Report..."
        
        With m_frmRptViewer
            .SetFilter cRptName, cReportTitle
            .Caption = cReportTitle
            .Show
        End With
        
        Log2Audit cRptName, "Generate " & cReportTitle

    Else
        MsgBox "Warning!" & vbCrLf & vbCrLf & "The report file " & cRptName & " is missing!" & vbCrLf & "Please verify that the report file exists in " & cReportPath, vbInformation, App.Title
    End If
    
    If Not (m_frmRptViewer Is Nothing) Then Set m_frmRptViewer = Nothing
    
    Exit Function

ErrGenerateRPT:
    ErrorMsg Err.Number, Err.Description, cReportTitle, cRptName
End Function


Public Sub Add2List(ByVal cString As String)
    With frmSplash.List1
        If .ListCount > 1000 Then .Clear
        .AddItem cString
        .ListIndex = .ListCount - 1
    End With
End Sub



'   usage:
'       0   n/a
'       1   sunday
'       2   holiday
'       3   regular day
Public Function ChkSwap(cDateValue As String) As Integer
    Dim cSqlStmt As String, _
        nValue As Integer, _
        oRecordSet As New ADODB.Recordset
    
    cSqlStmt = "select * from pa7927 where date1=" & cQuote & Format(cDateValue, "yyyy-mm-dd") & cQuote & " and status=1"
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ' --> check if date 2 is holiday...
        cSqlStmt = "select * from pa4329 " & _
                   " where (date=" & cQuote & Format(oRecordSet("date2"), "yyyy-mm-dd") & cQuote & ")" & _
                   " or ((month(date)=" & Month(oRecordSet("date2")) & ") and (day(date)=" & Day(oRecordSet("date2")) & ") and (fix_day=1))"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            nValue = 2
        Else
            nValue = IIf(Weekday(oRecordSet("date2")) = vbSunday, 1, 3)
        End If
    Else
        nValue = IIf(Weekday(cDateValue) = vbSunday, 1, 0)
        'nValue = IIf(Weekday(cDateValue) = vbSunday, 0, 1)
    End If
    
    ChkSwap = nValue
    
    Set oRecordSet = Nothing
End Function

' --> compute and breakdown dtr here...
' --> ComputeDays function
' | usage:
' |     cEmpID          Employee ID
' |     aPeriodInfo     Array consisting of
' |                         Date Start
' |                         Date End
' |                         Number of Holiday   - temporary
' |     aEmpStat        Array consisting of
' |                         Employment Status
' |                         WAP tag for Contractual
' |                         Paystatus tag for emergency
Public Function ComputeDays(ByVal cEmpID As String, _
                            aPeriodInfo As Variant, _
                            ByVal aEmpStat As Variant, _
                            Optional lClose As Boolean = False) As Variant
    On Error GoTo ErrCompute
    
    Dim aTimeInfo As Variant, aTmpTime As Variant, aShiftInfo As Variant, _
        cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, oRset1 As New ADODB.Recordset, _
        cDateIn As String, cTimein As String, cTimeOut As String, cTmpTime As String, aDateIn As String, _
        nTmpTime As Long, nTotTime As Long, nCtr As Integer, _
        lMultiDTR As Boolean, lDayBeyond As Boolean, _
        cShiftid As String, aShiftid As String, _
        cLeaveInfo As String, _
        nTrnType As Integer
    
    aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#)
    aTmpTime = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#)
'    (0) = regday
'    (1) = reg ot (max 2 hrs)
'    (2) = sa reg ot (excess of 2 hrs b4 10pm)
'    (3) = ndiff day
'    (4) = ndiff ot (excess of ndiff day)
'    (5) = Sunday
'    (6) = Sunday OT
'    (7) = holiday
'    (8) = Holiday OT hrs (reserved)
'    (9) = late (seconds)
'    (10) = tag for incomplete entry (daily)
'    (11) = tag for no entry/leave (daily)
'    (12) = sa ndiff ot --> added 20060314
'    (13) = Sunday ND
'    (14) = Sunday ND OT

    aShiftInfo = Array("", "", 0#, 0, 0#, 0#, 0, "", "", 0#, 0, 0#, 0#, 0)
'    aShiftInfo(0) = Start Time
'    aShiftInfo(1) = End Time
'    aShiftInfo(2) = Grace Period (minute)
'    aShiftInfo(3) = NDiff
'    aShiftInfo(4) = Reg Hour (hour)
'    aShiftInfo(5) = Break Time (hour)
'    aShiftInfo(6) = OT Tag (reg/nd)
    
    
    ' --> for leave...
    If aPeriodInfo(0) <> aPeriodInfo(1) Then
        cSqlStmt = "select if(date_start<" & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & ",date_start) as date_start, " & _
                   " if(date_end>" & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ",date_end) as date_end " & _
                   "From pa367583 " & _
                   "where (empid=" & cQuote & cEmpID & cQuote & ") and (status=1) " & _
                   "  and ((date_start between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ")" & _
                   "    or (date_end between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & "))"
    Else
        cSqlStmt = "select if(date_start<" & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & ",date_start) as date_start, " & _
                   " if(date_end>" & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ",date_end) as date_end " & _
                   "From pa367583 " & _
                   "where (empid=" & cQuote & cEmpID & cQuote & ") and (status=1) " & _
                   "  and ((" & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " between date_start and date_end)" & _
                   "    or (" & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & " between date_start and date_end))"
    End If
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            For nTotTime = 0 To DateDiff("d", objdbRs("date_start"), objdbRs("date_end")) 'objdbRs("date_end") - objdbRs("date_start")
                cDateIn = Format(DateAdd("d", nTotTime, objdbRs("date_start")), "yyyy-mm-dd")
                If Trim(cLeaveInfo) <> "" Then
                    If InStr(1, cLeaveInfo, cDateIn) = 0 Then
                        cLeaveInfo = cLeaveInfo & cQuote & cDateIn & cQuote & ","
                    End If
                Else
                    cLeaveInfo = cQuote & cDateIn & cQuote & ","
                End If
            Next nTotTime
            objdbRs.MoveNext
        Wend
        If Trim(cLeaveInfo) <> "" Then cLeaveInfo = left(cLeaveInfo, Len(cLeaveInfo) - 1)
    End If
    
    cSqlStmt = "select * from " & IIf(lClose, "pah84650", "pa84650") & " where (empid=" & cQuote & cEmpID & cQuote & ")" & _
               " and (logdate between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & _
               " and " & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ")" & _
               " order by empid, logdate, transdate, trantime"
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        DoEvents
        While Not oRecordSet.EOF
            
            If InStr(1, cLeaveInfo, Format(oRecordSet("logdate"), "yyyy-mm-dd")) > 0 Then
                aTimeInfo(11) = 2   ' --> tag for leave
                GoTo loopd2
            End If
            
'            MsgBox Format(oRecordSet("logdate"), "yyyy-mm-dd") + 1
'            If Format(oRecordSet("logdate"), "yyyy-mm-dd") = "2010-06-19" Then MsgBox "stop"
            
            cShiftid = oRecordSet("shiftid")
            
            ' --> retrieve time-in
            If (Trim(cDateIn) = "") Or (cDateIn <> oRecordSet("logdate")) Then
                
                ' --> accumulate running total here
                If Trim(cDateIn) <> "" Then
                    If (aTmpTime(0) = 0) And (aTmpTime(3) = 0) And (aTmpTime(5) = 0) Then aTmpTime(10) = 1
                    For nCtr = 0 To 14
'                        If nCtr = 0 Then
'                            aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ 3600) * 3600)
                        If (nCtr = 0) Or (nCtr = 3) Or (nCtr = 5) Then
                            aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ 900) * 900)
                        Else
                            If (nCtr < 10) Or (nCtr > 11) Then
                                aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ 1800) * 1800)
                            Else
                                aTimeInfo(nCtr) = aTimeInfo(nCtr) + aTmpTime(nCtr)
                            End If
                        End If
                    Next nCtr
                End If
                
                ' --> added 20060502
                nTrnType = oRecordSet("trantype")
                
                aTmpTime = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#)
                nTotTime = 0
                lMultiDTR = False
                
                If oRecordSet("trantype") = 0 Then
                    cDateIn = oRecordSet("logdate")
                    '---> Added 201703-02
                    aDateIn = Format(cDateIn, "yyyy-mm-dd")
'                    cDateIn = oRecordSet("transdate")
                    cTimein = oRecordSet("trantime")
                    cTimeOut = ""
                
                    ' --> retrieve shift info here
                    cSqlStmt = "select shiftid, time1, time2, `allowance`, ndiff, reg_hr, btime from pa74380 where shiftid=" & cQuote & cShiftid & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, False
                    If objdbRs.RecordCount > 0 Then
                        aShiftInfo(0) = objdbRs("time1")
                        aShiftInfo(1) = objdbRs("time2")
                        aShiftInfo(2) = objdbRs("allowance")
                        aShiftInfo(3) = objdbRs("ndiff")
                        aShiftInfo(4) = objdbRs("reg_hr") * 3600
                        aShiftInfo(5) = objdbRs("btime") * 3600
                        '--> For Alternate Shift (201703-02)
                        cSqlStmt = "select * from DI36770A where empid=" & cQuote & cEmpID & cQuote & _
                               " and date=" & cQuote & aDateIn & cQuote & _
                               " and periodid=" & cQuote & frmTMS2.Text1.Text & cQuote
                               OpenQueryDNS cSqlStmt, objdbRs, False
                          If IfExists("DI36770A", "(empid=" & cQuote & cEmpID & cQuote & ") and (date=" & cQuote & aDateIn & cQuote & ")") Then
                               aShiftid = objdbRs("shiftid")
                                '   get NDIFF,reghr and btime
                               cSqlStmt = "select * from pa74380 where shiftid=" & cQuote & aShiftid & cQuote
                               OpenQueryDNS cSqlStmt, objdbRs, False
                               aShiftInfo(10) = objdbRs("ndiff")
                               aShiftInfo(11) = objdbRs("reg_hr") * 3600
                               aShiftInfo(12) = objdbRs("btime") * 3600
                               cSqlStmt = "select * from DI36770A where empid=" & cQuote & cEmpID & cQuote & _
                               " and date=" & cQuote & aDateIn & cQuote & _
                               " and periodid=" & cQuote & frmTMS2.Text1.Text & cQuote
                               OpenQueryDNS cSqlStmt, objdbRs, False
                                If IfExists("DI36770A", "(shiftid=" & cQuote & aShiftid & cQuote & ") and (empid=" & cQuote & cEmpID & cQuote & ") and (date=" & cQuote & aDateIn & cQuote & ")") Then
                                    aShiftInfo(7) = objdbRs("time1")
                                    aShiftInfo(8) = objdbRs("time2")
                                       If (DateDiff("s", aShiftInfo(0), aShiftInfo(7))) = 3600 Or (DateDiff("s", aShiftInfo(0), aShiftInfo(7))) = -3600 Or (DateDiff("s", aShiftInfo(7), aShiftInfo(0))) = 3600 Or (DateDiff("s", aShiftInfo(7), aShiftInfo(0))) = -3600 Then
                                            If DateDiff("s", cTimein, aShiftInfo(0)) <= -3000 Or DateDiff("s", cTimein, aShiftInfo(0)) >= 3960 Then
                                                aShiftInfo(0) = aShiftInfo(7)
                                                aShiftInfo(1) = aShiftInfo(8)
                                                aShiftInfo(3) = aShiftInfo(10)
                                                aShiftInfo(4) = aShiftInfo(11)
                                                aShiftInfo(5) = aShiftInfo(12)
                                            End If
                                       Else
                                            If DateDiff("s", cTimein, aShiftInfo(0)) <= -5100 Or DateDiff("s", cTimein, aShiftInfo(0)) >= 5100 Then
                                                aShiftInfo(0) = aShiftInfo(7)
                                                aShiftInfo(1) = aShiftInfo(8)
                                                aShiftInfo(3) = aShiftInfo(10)
                                                aShiftInfo(4) = aShiftInfo(11)
                                                aShiftInfo(5) = aShiftInfo(12)
                                            End If
                                       End If
                                End If
                          End If
                    Else
                        cDateIn = ""
                        cTimein = ""
                        aShiftInfo = Array("", "", 0#, 0, 0#, 0#)
                        GoTo loopd2
                    End If
                    
                    
'                    (201704-26) TLC
                    If (DateDiff("n", aShiftInfo(0), cTimein) >= aShiftInfo(2) And DateDiff("n", aShiftInfo(0), cTimein) >= 0) Then
                               lCheckLate = True
                    Else
                               lCheckLate = False
                    End If
                    
'                    ' --> check late in shift's grace period
'                    If DateDiff("s", cTimein, DateAdd("n", aShiftInfo(2), aShiftInfo(0))) >= 0 Then
'                        aTmpTime(9) = 0
'                    Else    ' --> interval of 15 mins
'                        aTmpTime(9) = DateDiff("s", aShiftInfo(0), cTimein)
'                        aTmpTime(9) = ((aTmpTime(9) \ 900) + IIf((aTmpTime(9) Mod 900 > 0), 1, 0)) * 900
'                    End If

                    If DateDiff("s", cTimein, DateAdd("n", aShiftInfo(2), aShiftInfo(0))) >= 0 Then
'                        MsgBox DateAdd("n", aShiftInfo(2), aShiftInfo(0))
                        aTmpTime(9) = 0
                    Else    ' --> interval of 15 mins
'                        If gCompanyID = "0003" Then
                            aTmpTime(9) = DateDiff("s", aShiftInfo(0), cTimein)
'                        Else
'                            aTmpTime(9) = DateDiff("s", aShiftInfo(0), cTimein)
'                            aTmpTime(9) = ((aTmpTime(9) \ 900) + IIf((aTmpTime(9) Mod 900 > 0), 1, 0)) * 900
'                        End If
                    End If
                End If
                
            End If
            
            If oRecordSet("trantype") = 0 Then
            
'                lMultiDTR = IIf((Weekday(cDateIn) = vbSunday) Or (ChkSwap(cDateIn) = 1), aTmpTime(5), aTmpTime(0)) > 0
                lMultiDTR = IIf(ChkSwap(cDateIn) = 1, aTmpTime(5), aTmpTime(0)) > 0
                cTimein = oRecordSet("trantime")
                
            Else
            
                ' --> added 20060502
                If nTrnType = 1 Then GoTo loopd2
                
                lDayBeyond = oRecordSet("transdate") <> oRecordSet("logdate")
                cTimeOut = oRecordSet("trantime")
                
                If lMultiDTR Then
                
                    ' --> for multiple time in/out
'                    If (Weekday(cDateIn) = vbSunday) Or (ChkSwap(cDateIn) = 1) Then
                    If ChkSwap(cDateIn) = 1 Then
                        nTotTime = DateDiff("s", cTimein, cTimeOut) ' - aTmpTime(9)
                        If aTmpTime(5) < 28800 Then     ' --> if undertime
                            If (nTotTime - (28800 - aTmpTime(5))) >= 0 Then
                                aTmpTime(5) = 28800
                                ' --> add excess time to Sun OT
                                aTmpTime(6) = nTotTime - (28800 - aTmpTime(5))
                            Else
                                aTmpTime(5) = aTmpTime(5) + nTotTime
                                nTotTime = 0
                            End If
                        Else
                            ' --> add excess time to Sun OT
                            aTmpTime(6) = aTmpTime(6) + nTotTime
                        End If
                    Else
                        If aTmpTime(0) < 28800 Then     ' --> if undertime
                            If lDayBeyond Then
                                nTotTime = DateDiff("s", cTimein, gNDiffTime)
                                If nTotTime > 0 Then
                                    aTmpTime(0) = aTmpTime(0) + nTotTime
                                    
                                    If aTmpTime(0) > 28800 Then
                                        If aEmpStat(1) = 1 Then
                                            aTmpTime(2) = aTmpTime(2) + (28800 - aTmpTime(0))
                                        Else
                                            aTmpTime(1) = aTmpTime(1) + (28800 - aTmpTime(0))
                                            If aTmpTime(1) > 7200 Then
                                                aTmpTime(2) = aTmpTime(1) - 7200
                                                aTmpTime(1) = 7200
                                            End If
                                        End If
                                    End If
                                End If
                                
                                nTotTime = DateDiff("s", oRecordSet("logdate") & " " & gNDiffTime, oRecordSet("transdate") & " " & cTimeOut)
                                If nTotTime > 0 Then
                                    If aEmpStat(1) = 1 Then
                                        aTmpTime(12) = nTotTime
                                    Else
                                        aTmpTime(4) = nTotTime
                                        If aTmpTime(4) > 7200 Then
                                            aTmpTime(12) = (aTmpTime(4) - 7200)
                                            aTmpTime(4) = 7200
                                        End If
                                    End If
'                                    aTmpTime(IIf(aEmpStat(1) = 1, 12, 4)) = nTotTime
                                End If
                            Else
                                If DateDiff("s", gNDiffTime, cTimeOut) > 0 Then
                                    nTotTime = DateDiff("s", cTimein, gNDiffTime)
                                    If nTotTime > 0 Then
                                        aTmpTime(0) = aTmpTime(0) + nTotTime
                                        
                                        If aTmpTime(0) > 28800 Then
                                            If aEmpStat(1) = 1 Then
                                                aTmpTime(2) = aTmpTime(2) + (28800 - aTmpTime(0))
                                            Else
                                                aTmpTime(1) = aTmpTime(1) + (28800 - aTmpTime(0))
                                                If aTmpTime(1) > 7200 Then
                                                    aTmpTime(2) = aTmpTime(1) - 7200
                                                    aTmpTime(1) = 7200
                                                End If
                                            End If
                                        End If
                                    End If
                                
                                    nTotTime = DateDiff("s", gNDiffTime, cTimeOut)
                                    If nTotTime > 0 Then
                                        If aEmpStat(1) = 1 Then
                                            aTmpTime(12) = nTotTime
                                        Else
                                            aTmpTime(4) = nTotTime
                                            If aTmpTime(4) > 7200 Then
                                                aTmpTime(12) = (aTmpTime(4) - 7200)
                                                aTmpTime(4) = 7200
                                            End If
                                        End If
                                    End If
                                Else
                                    nTotTime = DateDiff("s", cTimein, cTimeOut) ' - aTmpTime(9)
                                    If nTotTime > 0 Then
                                        aTmpTime(0) = aTmpTime(0) + nTotTime
                                        
                                        If aTmpTime(0) > 28800 Then
                                            If aEmpStat(1) = 1 Then
                                                aTmpTime(2) = aTmpTime(2) + (aTmpTime(0) - 28800)
                                            Else
                                                aTmpTime(1) = aTmpTime(1) + (aTmpTime(0) - 28800)
                                                If aTmpTime(1) > 7200 Then
                                                    aTmpTime(2) = aTmpTime(1) - 7200
                                                    aTmpTime(1) = 7200
                                                End If
                                            End If
                                            aTmpTime(0) = 28800
                                        End If
                                    End If
                                End If
                            End If
                        Else
                        
                            If lDayBeyond Then
                                nTotTime = DateDiff("s", cTimein, gNDiffTime)
                                If nTotTime > 0 Then
                                    If aEmpStat(1) = 1 Then
                                        aTmpTime(2) = aTmpTime(2) + nTotTime
                                    Else
                                        aTmpTime(1) = aTmpTime(1) + nTotTime
                                        If aTmpTime(1) > 7200 Then
                                            aTmpTime(2) = aTmpTime(2) + (aTmpTime(1) - 7200)
                                            aTmpTime(1) = 7200
                                        End If
                                    End If
                                End If
                            End If
                        
                            If lDayBeyond Then
                                nTotTime = DateDiff("s", oRecordSet("logdate") & " " & gNDiffTime, oRecordSet("transdate") & " " & cTimeOut)
                            Else
                                nTotTime = DateDiff("s", cTimein, cTimeOut) ' - aTmpTime(9)
                            End If
                            
                            If nTotTime > 0 Then
                                If lDayBeyond Then
                                    If aEmpStat(1) = 1 Then
                                        aTmpTime(2) = aTmpTime(2) + nTotTime
                                    Else
                                        aTmpTime(12) = aTmpTime(12) + nTotTime
                                    End If
                                Else
                                    If aEmpStat(1) = 1 Then
                                        aTmpTime(2) = aTmpTime(2) + nTotTime
                                    Else
                                        aTmpTime(1) = aTmpTime(1) + nTotTime
                                        If aTmpTime(1) > 7200 Then
                                            aTmpTime(2) = aTmpTime(1) - 7200
                                            aTmpTime(1) = 7200
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        
                    End If
                    
                Else
                
'                    ' --> for single time in/out
'                    If cTimein <> "" Then   ' --> check if there's an IN transaction
'                        If aShiftInfo(3) = 1 Then
'                            nTotTime = ((DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), IIf(DateDiff("s", DateAdd("d", 1, oRecordSet("logdate")) & " " & aShiftInfo(1), oRecordSet("logdate") & " " & cTimeOut) > 0, aShiftInfo(1), cTimeOut)) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600))
'                        Else
'                            If gCompanyID = "0003" Then
'                                nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)) * IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)
'                            Else
'                                If (DateDiff("n", aShiftInfo(0), cTimein) <= aShiftInfo(2)) Then
'
'                                    nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600)
''                                    If nTotTime > 0 Then
''                                        nTotTime = 28800
''                                    Else
''                                        nTotTime = 0
''                                    End If
'                                Else
'                                    nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600)
'                                End If
'                            End If
'
'                        End If

                    ' --> revised 20161109(1)
                    ' --> for single time in/out
                    If cTimein <> "" Then   ' --> check if there's an IN transaction
                        If aShiftInfo(3) = 1 Then
                            nTotTime = ((DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), IIf(DateDiff("s", DateAdd("d", 1, oRecordSet("logdate")) & " " & aShiftInfo(1), oRecordSet("logdate") & " " & cTimeOut) > 0, aShiftInfo(1), cTimeOut)) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600))
                        Else
                            If gCompanyID = "0003" Then
                                If (DateDiff("n", aShiftInfo(0), cTimein) <= aShiftInfo(2) And DateDiff("n", aShiftInfo(0), cTimein) >= 0) Then
                                    nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) < 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)) * IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)
                                Else
                                     nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)) * IIf(lCheckLate, IIf(aTmpTime(9) = 0, 900, aTmpTime(9)), 3600)
                                End If
                                
                            Else
                                If (DateDiff("n", aShiftInfo(0), cTimein) <= aShiftInfo(2) And DateDiff("n", aShiftInfo(0), cTimein) >= 0) Then
                                   
                                    nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) < 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600)
'                                    If nTotTime > 0 Then
'                                        nTotTime = 28800
'                                    Else
'                                        nTotTime = 0
'                                    End If
                                Else
                                    nTotTime = (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), cTimeOut) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600)
                                End If
                            End If

                        End If
                    
                        If (nTotTime < 28800) And (Not lDayBeyond) Then
                        
                            ' --> re-compute if nTotTime > half day's work... 20070920
                            If nTotTime > 14400 Then
                                If (DateDiff("s", aShiftInfo(1), cTimein) < 0) Then 'And (DateDiff("s", cTimeOut, aShiftInfo(1)) > 0) Then
                                    If (DateDiff("s", cTimeOut, aShiftInfo(1)) > 0) Then
                                        nTotTime = 28800 - DateDiff("s", cTimeOut, aShiftInfo(1)) - (DateDiff("s", aShiftInfo(0), IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein)))
                                    Else
                                        nTotTime = 28800 - (DateDiff("s", aShiftInfo(0), IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein)))
                                        
                                    End If
                                End If
                            End If
                        
                            ' --> undertime d2...
'                            nTotTime = ((DateDiff("s", IIf(DateDiff("s", cTimeIn, aShiftInfo(0)) > 0, aShiftInfo(0), cTimeIn), IIf(DateDiff("s", aShiftInfo(1), cTimeOut) > 0, aShiftInfo(1), cTimeOut)) \ IIf(lCheckLate, 900, 3600)) * IIf(lCheckLate, 900, 3600))
                            If ChkSwap(cDateIn) = 1 Then
                                aTmpTime(5) = aTmpTime(5) + nTotTime
                            Else
                                If aShiftInfo(3) = 1 Then
                                    If (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), gNDiffTime) - aTmpTime(9)) < nTotTime Then
                                        aTmpTime(0) = aTmpTime(0) + (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), gNDiffTime) - aTmpTime(9))
                                        aTmpTime(3) = aTmpTime(3) + nTotTime - (DateDiff("s", IIf(DateDiff("s", cTimein, aShiftInfo(0)) > 0, aShiftInfo(0), cTimein), gNDiffTime) - aTmpTime(9))
                                    Else
                                        aTmpTime(0) = aTmpTime(0) + nTotTime
                                    End If
                                Else
                                    aTmpTime(0) = aTmpTime(0) + nTotTime
                                End If
                            End If
                            
                            If aShiftInfo(3) = 1 Then
                                nTotTime = DateDiff("s", DateAdd("d", 1, oRecordSet("logdate")) & " " & aShiftInfo(1), oRecordSet("logdate") & " " & cTimeOut)
                            Else
                                If DateDiff("s", aShiftInfo(1), cTimein) < 0 Then
                                    nTotTime = DateDiff("s", aShiftInfo(1), cTimeOut)
                                Else
                                    nTotTime = 0
                                End If
                            End If
                            
                            If nTotTime > 0 Then
                                If ChkSwap(cDateIn) = 1 Then
                                    aTmpTime(6) = aTmpTime(6) + nTotTime
                                Else
                                    If aEmpStat(1) = 1 Then
                                        aTmpTime(2) = nTotTime
                                    Else
                                        If (nTotTime - 7200) <= 0 Then
                                            ' --> Reg OT
                                            aTmpTime(1) = nTotTime
                                        Else
                                            ' --> reg OT & SA reg OT
                                            aTmpTime(1) = 7200
                                            aTmpTime(2) = nTotTime - 7200
                                        End If
                                    End If
                                End If
                            End If
                        
                        Else
                            
                            If aShiftInfo(3) = 1 Then
                                
                                If ChkSwap(cDateIn) = 1 Then
                                    ' pg sunday na night diff pa...
                                    aTmpTime(13) = 28800 - IIf(lCheckLate, aTmpTime(9), 0)
                                    
                                    nTotTime = DateDiff("s", aShiftInfo(1), cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
                                    If nTotTime < 0 Then
                                        aTmpTime(13) = 28800 + nTotTime
                                        nTotTime = 0
                                    Else
                                        aTmpTime(14) = nTotTime
                                    End If
                                    
                                Else
                                    ' --> NDiff Day & OT
                                    aTmpTime(3) = 28800 - IIf(lCheckLate, aTmpTime(9), 0)
                                    
                                    ' 20060516 aTmpTime(4) = DateDiff("s", aShiftInfo(1), cTimeOut) - aTmpTime(9)
                                    nTotTime = DateDiff("s", aShiftInfo(1), cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
                                    If nTotTime < 0 Then
                                        aTmpTime(3) = 28800 + nTotTime
                                        nTotTime = 0
                                    End If
                                        
                                    If DateDiff("s", aShiftInfo(1), "06:00:00") > 0 Then
                                        If aEmpStat(0) <> 0 Then
                                            If aEmpStat(1) = 1 Then
                                                aTmpTime(2) = nTotTime
                                            Else
'                                                aTmpTime(4) = 7200
'                                                aTmpTime(12) = nTotTime - 7200
                                            
                                                If (nTotTime - 7200) <= 0 Then
                                                    'aTmpTime(1) = nTotTime
                                                    aTmpTime(4) = nTotTime
                                                Else
                                                    '2009-06-25 reivise for the new computation of ndiff ot
                                                    'aTmpTime(2) = 7200
                                                    'aTmpTime(2) = nTotTime - 7200
                                                    aTmpTime(4) = 7200
                                                    aTmpTime(12) = nTotTime - 7200
                                                End If
                                            End If
                                        Else
                                            'aTmpTime(2) = nTotTime
                                            aTmpTime(12) = nTotTime
                                        End If
                                    Else
                                        If aEmpStat(0) <> 0 Then
                                            If aEmpStat(1) = 1 Then
                                                ' --> SA ND OT
                                                aTmpTime(12) = nTotTime
                                            Else
                                                If (nTotTime - 7200) <= 0 Then
                                                    ' --> ND OT
                                                    aTmpTime(4) = nTotTime
                                                Else
                                                    ' --> ND OT & SA ND OT
                                                    aTmpTime(4) = 7200
                                                    aTmpTime(12) = nTotTime - 7200
                                                End If
                                            End If
                                        Else
                                            aTmpTime(12) = nTotTime
                                        End If
                                    End If
                                End If
                                
                            Else
'                                If (Weekday(cDateIn) = vbSunday) Or (ChkSwap(cDateIn) = 1) Then
                                If ChkSwap(cDateIn) = 1 Then
'                                 If ChkSwap(oRecordSet("transdate")) = 1 Then
                                    ' --> Sunday Hr & OT
                                    aTmpTime(5) = 28800 - IIf(lCheckLate, aTmpTime(9), 0)
                                    
'                                     --> revised 20070703
'                                    If lDayBeyond Then
'                                        aTmpTime(6) = DateDiff("s", aShiftInfo(1), gNDiffTime) - IIf(lCheckLate, 0, aTmpTime(9))
'                                        aTmpTime(14) = DateDiff("s", oRecordSet("logdate") & " " & gNDiffTime, oRecordSet("transdate") & " " & cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
'                                    Else
'                                        aTmpTime(6) = DateDiff("s", oRecordSet("logdate") & " " & aShiftInfo(1), oRecordSet("transdate") & " " & cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
'                                    End If

                                    ' --> revised 20161109(2)
                                    If lDayBeyond Then
                                        aTmpTime(6) = DateDiff("s", aShiftInfo(1), gNDiffTime) - IIf(lCheckLate, 0, aTmpTime(9))
                                        aTmpTime(14) = DateDiff("s", oRecordSet("logdate") & " " & gNDiffTime, oRecordSet("transdate") & " " & cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
                                        nTotTime = DateDiff("s", aShiftInfo(1), gNDiffTime) - IIf(lCheckLate, 0, aTmpTime(9))
                                    Else
                                        aTmpTime(6) = DateDiff("s", oRecordSet("logdate") & " " & aShiftInfo(1), oRecordSet("transdate") & " " & cTimeOut) - IIf(lCheckLate, 0, aTmpTime(9))
                                        nTotTime = DateDiff("s", aShiftInfo(1), IIf(DateDiff("s", gNDiffTime, cTimeOut) <= 0, cTimeOut, gNDiffTime)) - IIf(lCheckLate, 0, aTmpTime(9))
                                    End If
                                
                                    If nTotTime < 0 Then
                                        aTmpTime(5) = 28800 + nTotTime
                                        nTotTime = 0
                                    End If

                                Else
                                                                        
                                    ' --> Reg Day
'                                    nTotTime
                                    'aTmpTime(0) = 28800 - IIf(lCheckLate, aTmpTime(9), 0)
                                    aTmpTime(0) = 28800 - IIf(lCheckLate, aTmpTime(9), 0)
                                    
                                    If lDayBeyond Then
                                        nTotTime = DateDiff("s", aShiftInfo(1), gNDiffTime) - IIf(lCheckLate, 0, aTmpTime(9))
                                    Else
                                        nTotTime = DateDiff("s", aShiftInfo(1), IIf(DateDiff("s", gNDiffTime, cTimeOut) <= 0, cTimeOut, gNDiffTime)) - IIf(lCheckLate, 0, aTmpTime(9))
                                    End If
                                    
                                    
                                    ' --> 20060504
                                    If nTotTime < 0 Then
                                        aTmpTime(0) = 28800 + nTotTime
                                        nTotTime = 0
                                    End If

'                                    ' --> 20100703
'                                    If nTotTime < 0 Then
'                                        If lDayBeyond = False Then
'                                            aTmpTime(0) = 28800 + nTotTime
'                                            nTotTime = 0
'                                        End If
'                                    End If
                                    
                                    ' --> check if Emp Stat is WAP
                                    If aEmpStat(0) <> 0 Then
                                        ' --> retrieve nDiff OT here, time consumed beyond 10pm...
                                        If (DateDiff("s", gNDiffTime, cTimeOut) >= 0) Or lDayBeyond Then
                                        
'                                            remarked 20060314
'                                            this should be treated as SA NDiff OT
'                                            If lDayBeyond Then
'                                                aTmpTime(4) = DateDiff("s", "22:00:00", "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + 1
'                                            Else
'                                                aTmpTime(4) = DateDiff("s", "22:00:00", cTimeOut)
'                                            End If
                                        
                                            
                                            If lDayBeyond Then
                                                
                                                OpenQueryDNS "select * from pa4329 where date =" & cQuote & Format(oRecordSet("transdate"), "yyyy-mm-dd") & cQuote, objdbRs, False
                                                If objdbRs.RecordCount > 0 Then
'                                                    MsgBox DateDiff("s", gNDiffTime, "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + DateDiff("s", "00:00:00", cTimeOut) + 1
                                                    '20080610
                                                    aTmpTime(12) = DateDiff("s", gNDiffTime, "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + 1
    '                                                aTmpTime(12) = DateDiff("s", gNDiffTime, "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + 1
                                                Else
                                                    aTmpTime(12) = DateDiff("s", gNDiffTime, "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + 1
                                                End If
                                                
                                            Else
                                                aTmpTime(12) = DateDiff("s", gNDiffTime, cTimeOut)
                                            End If
                                        End If
                                        
                                        If aEmpStat(1) = 1 Then
                                            aTmpTime(2) = nTotTime
                                        Else
                                            If (nTotTime - 7200) <= 0 Then
                                                ' --> Reg OT
                                                aTmpTime(1) = nTotTime
                                            Else
                                                ' --> reg OT & SA reg OT
                                                aTmpTime(1) = 7200
                                                aTmpTime(2) = nTotTime - 7200
                                            End If
                                        End If
                                    Else
                                        If (DateDiff("s", gNDiffTime, cTimeOut) >= 0) Or lDayBeyond Then
                                            If lDayBeyond Then
                                                aTmpTime(12) = DateDiff("s", gNDiffTime, "23:59:59") + DateDiff("s", "00:00:00", cTimeOut) + 1
                                            Else
                                                aTmpTime(12) = DateDiff("s", gNDiffTime, cTimeOut)
                                            End If
                                        End If
                                        aTmpTime(2) = nTotTime
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            ' --> added 20060502
            nTrnType = oRecordSet("trantype")
            
loopd2:
            oRecordSet.MoveNext
        Wend
    
        For nCtr = 0 To 14
'        MsgBox aTimeInfo(0)
            If (nCtr = 0) Or (nCtr = 3) Or (nCtr = 5) Then
'                aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ 3600) * 3600)
                
                aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ 900) * 900)
            Else
                If (nCtr < 10) Or (nCtr > 11) Then
                    aTimeInfo(nCtr) = aTimeInfo(nCtr) + ((aTmpTime(nCtr) \ nOTInterval) * nOTInterval)
                Else
                    aTimeInfo(nCtr) = aTimeInfo(nCtr) + aTmpTime(nCtr)
                End If
            End If
        Next nCtr
    
        If aTimeInfo(11) = 0 Then
            aTimeInfo(10) = IIf((aTimeInfo(0) = 0) And (aTimeInfo(3) = 0) And (aTimeInfo(5) = 0) And (aTimeInfo(13) = 0), 1, 0)
        End If
    
        ' --> retrieve holiday here...
        aTimeInfo(7) = IIf((aEmpStat(0) <> 0) And (Not ((aEmpStat(0) = 1) And (aEmpStat(1) = 1))), aPeriodInfo(2), 0)
    Else
        If aPeriodInfo(0) = aPeriodInfo(1) Then
            If InStr(1, cLeaveInfo, Format(aPeriodInfo(0), "yyyy-mm-dd")) > 0 Then
                aTimeInfo(11) = 2   ' --> tag for leave
            Else
                aTimeInfo(11) = 1
            End If
        Else
            aTimeInfo(11) = 1
        End If
    End If
    
EndCompute:
    Set oRecordSet = Nothing
    Set oRset1 = Nothing
    
    
    'Revised (201704-26)
    'If gCompanyID = "0003" Then
    'Revised (201705-29)
    If aTimeInfo(0) < 28800 Then
        aTimeInfo(0) = aTmpTime(0)
        aTimeInfo(3) = aTmpTime(3)
        aTimeInfo(5) = aTmpTime(5)
        aTimeInfo(13) = aTmpTime(13)
    'End If
    End If

    aTimeInfo(0) = Round((aTimeInfo(0) / 3600) / 8, 4)   ' --> Reg Day
    aTimeInfo(3) = Round((aTimeInfo(3) / 3600) / 8, 4)   ' --> NDiff Day
    
    ' --> for emergency manpower, 20070831
    If aEmpStat(2) > 0 Then
        aTimeInfo(1) = Round((aTimeInfo(1) + aTimeInfo(2)) / 3600, 4)   ' --> Reg OT
        aTimeInfo(4) = Round((aTimeInfo(4) + aTimeInfo(12)) / 3600, 4)  ' --> NDiff OT
        aTimeInfo(2) = 0                                                ' --> SA Reg OT
        aTimeInfo(12) = 0                                               ' --> SA NDiff OT
        aTimeInfo(7) = 0                                ' --> No Holiday
    Else
        ' --> revised 20070328
        If lExtension Then
            aTimeInfo(1) = Round(aTimeInfo(1) / 3600, 4)         ' --> Reg OT
            aTimeInfo(2) = Round(aTimeInfo(2) / 3600, 4)         ' --> SA Reg OT
            aTimeInfo(4) = Round(aTimeInfo(4) / 3600, 4)         ' --> NDiff OT
            aTimeInfo(12) = Round(aTimeInfo(12) / 3600, 4)       ' --> SA NDiff OT
        Else
'   ----------- for combine purpose only
'            If gCompanyID = "0002" Then
'                If lAudit = 0 Then
'                    aTimeInfo(1) = Round(aTimeInfo(1) / 3600, 4)    ' --> Reg OT
'                    aTimeInfo(2) = Round(aTimeInfo(2) / 3600, 4)    ' --> SA Reg OT
'                    aTimeInfo(4) = Round(aTimeInfo(4) / 3600, 4)    ' --> NDiff OT
'                    aTimeInfo(12) = Round(aTimeInfo(12) / 3600, 4)  ' --> SA NDiff OT
'                Else
'                    aTimeInfo(1) = Round((aTimeInfo(1) + aTimeInfo(2)) / 3600, 4)   ' --> Reg OT
'                    aTimeInfo(4) = Round((aTimeInfo(4) + aTimeInfo(12)) / 3600, 4)  ' --> NDiff OT
'                    aTimeInfo(2) = 0                                                ' --> SA Reg OT
'                    aTimeInfo(12) = 0                                               ' --> SA NDiff OT
'                End If
'            Else
'                aTimeInfo(1) = Round((aTimeInfo(1) + aTimeInfo(2)) / 3600, 4)   ' --> Reg OT
'                aTimeInfo(4) = Round((aTimeInfo(4) + aTimeInfo(12)) / 3600, 4)  ' --> NDiff OT
'                aTimeInfo(2) = 0                                                ' --> SA Reg OT
'                aTimeInfo(12) = 0                                               ' --> SA NDiff OT
'            End If
            aTimeInfo(1) = Round((aTimeInfo(1) + aTimeInfo(2)) / 3600, 4)   ' --> Reg OT
            aTimeInfo(4) = Round((aTimeInfo(4) + aTimeInfo(12)) / 3600, 4)  ' --> NDiff OT
            aTimeInfo(2) = 0                                                ' --> SA Reg OT
            aTimeInfo(12) = 0
        End If
    End If
    
    aTimeInfo(5) = Round(aTimeInfo(5) / 3600, 4)         ' --> Sunday
    aTimeInfo(6) = Round(aTimeInfo(6) / 3600, 4)         ' --> Sunday OT
    
    ' --> 20070629
    aTimeInfo(13) = Round(aTimeInfo(13) / 3600, 4)         ' --> Sunday ND
    aTimeInfo(14) = Round(aTimeInfo(14) / 3600, 4)         ' --> Sunday ND OT
    
    ComputeDays = aTimeInfo
    
    Exit Function

ErrCompute:
    ErrorMsg Err.Number, Err.Description, "Compute Days Worked", "ComputeDays"
    
    Resume EndCompute
End Function


Public Function ChkPersonnel(ByVal oCombo As Control, Optional ByVal cMessage As String = "missing entry detected.") As Boolean
    If TypeName(oCombo) = "ComboBox" Then
        If (oCombo.ListIndex = -1) Or Trim(oCombo.Text) = "" Then
            MsgBox "Advisory!" & vbCrLf & "Cannot save record, " & cMessage, vbCritical, App.Title
            oCombo.SetFocus
            ChkPersonnel = False
        Else
            ChkPersonnel = True
        End If
    ElseIf TypeName(oCombo) = "TextBox" Then
        If Trim(oCombo.Text) = "" Then
            MsgBox "Advisory!" & vbCrLf & "Cannot save record, " & cMessage, vbCritical, App.Title
            oCombo.SetFocus
            ChkPersonnel = False
        Else
            ChkPersonnel = True
        End If
    End If
End Function


' + -->
' |     Procedure Name  :   ComputeTax
' |     Description     :   Compute Annual Withholding Tax
' |     Date Created    :   05 jan 2007
' + -->
Public Function ComputeTax(ByVal cPeriodID As String, _
                    ByVal cEmpID As String, _
                    ByVal cDedID As String, _
                    ByVal nYear As Integer, _
                    ByVal nCurGross As Double, _
                    ByVal nCurExempt As Double, _
                    Optional ByVal lUpdate As Boolean = True) As Variant
    Dim cParam As String, _
        cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        nDedExempt As Double, _
        aTaxable As Variant
    
    aTaxable = Array(0#, 0#, 0#, 0#, 0#)
'    aTaxable(0)     Taxable Amount
'    aTaxable(1)     Net Taxable
'    aTaxable(2)     excess
'    aTaxable(3)     undefined
'    aTaxable(4)     undefined

    cSqlStmt = "select dedid, sum(ded_amt) as tot_ded " & _
               "From pah87263 " & _
               "where (empid=" & cQuote & cEmpID & cQuote & ") and (dedid in (" & cDedID & ")) " & _
               " and (periodid in (select periodid from pa7730 where year(date_end)=" & nYear & "))" & _
               "group by empid "
    OpenQueryDNS cSqlStmt, objdbRs, False
    nDedExempt = IIf(objdbRs.RecordCount > 0, objdbRs("tot_ded"), 0)
    
    cSqlStmt = "select sum(leave_pay) as leave_pay " & _
               "From pah87260 " & _
               "where (periodid in (select periodid from pa7730 where year(date_end)=" & nYear & ")) and (empid=" & cQuote & cEmpID & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, False
    aTaxable(3) = IIf(objdbRs.RecordCount > 0, objdbRs("leave_pay"), 0)
    
'    cSqlStmt = "select a.empid, " & _
'               "  a.ytd_wtax, " & _
'               "  a.ytd_gross, " & _
'               "  b.amount as tax_amt, " & _
'               "  b.percent, b.range1, " & _
'               "  if(instr(ifnull(c.taxcode,a.taxcode),'S')>0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H')>0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'M')>0,b.m_amt,0))) as amt_exempt, " & _
'               "  if(right(ifnull(c.taxcode,a.taxcode),1)>4,4,right(ifnull(c.taxcode,a.taxcode),1)) * b.ex_amt as add_ex_amt " & _
'               "from di3670 a " & _
'               "  left join pa8290 c on a.taxid=c.taxid" & _
'               "  left join PA4870 b on (a.ytd_gross+" & nCurGross - nCurExempt - nDedExempt - aTaxable(3) & ")- " & _
'               "                        (if(instr(ifnull(c.taxcode,a.taxcode),'S') > 0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H') > 0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'ME') > 0,b.m_amt,0))) + " & _
'               "                         if(if(instr(ifnull(c.taxcode,a.taxcode),'S') > 0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H') > 0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'ME') > 0,b.m_amt,0))) > 0,if(right(ifnull(c.taxcode,a.taxcode),1)>4,4,right(ifnull(c.taxcode,a.taxcode),1)) * b.ex_amt,0)) " & _
'               "                        between b.range1 and b.range2 " & _
'               "where a.empid=" & cQuote & cEmpID & cQuote
    
    'for K1 20090107
    cSqlStmt = "select a.empid, " & _
               "  a.ytd_wtax, " & _
               "  a.ytd_gross, " & _
               "  a.ytd_gross_sa, " & _
               "  b.amount as tax_amt, " & _
               "  b.percent, b.range1, " & _
               "  if(instr(ifnull(c.taxcode,a.taxcode),'S')>0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H')>0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'M')>0,b.m_amt,0))) as amt_exempt, " & _
               "  if(right(ifnull(c.taxcode,a.taxcode),1)>4,4,right(ifnull(c.taxcode,a.taxcode),1)) * b.ex_amt as add_ex_amt " & _
               "from di3670 a " & _
               "  left join pa8290 c on a.taxid=c.taxid" & _
               "  left join PA4870 b on (a.ytd_gross+" & nCurGross - nCurExempt - nDedExempt - aTaxable(3) & ")- " & _
               "                        (if(instr(ifnull(c.taxcode,a.taxcode),'S') > 0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H') > 0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'ME') > 0,b.m_amt,0))) + " & _
               "                         if(if(instr(ifnull(c.taxcode,a.taxcode),'S') > 0,b.s_amt,if(instr(ifnull(c.taxcode,a.taxcode),'H') > 0,b.h_amt,if(instr(ifnull(c.taxcode,a.taxcode),'ME') > 0,b.m_amt,0))) > 0,if(right(ifnull(c.taxcode,a.taxcode),1)>4,4,right(ifnull(c.taxcode,a.taxcode),1)) * b.ex_amt,0)) " & _
               "                        between b.range1 and b.range2 " & _
               "where a.empid=" & cQuote & cEmpID & cQuote
    
    
    Script2File cSqlStmt
    
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        
        If oRecordSet("amt_exempt") > 0 Then
        
            cSqlStmt = " select sum(SUN_PAY)+sum(SUN_OT_PAY)+sum(SUN_COLA)+ sum(SUN_ND_PAY)+sum(SUN_ND_OT_PAY) as suntot " & _
                       " From pah87260 " & _
                       " where periodid in ( " & _
                       " select periodid from pa7730 " & _
                       " where year(date_start)=2008 and 13month <> 1) and empid = " & cQuote & oRecordSet("empid") & cQuote
            OpenQueryDNS cSqlStmt, oRSet2, False
        
            
            '    aTaxable(0)     Taxable Amount
            '    aTaxable(1)     Net Taxable
            '    aTaxable(2)     excess
            '    aTaxable(3)     undefined
            '    aTaxable(4)     undefined
            
'            If gCompanyID <> "0002" Then
'                aTaxable(0) = Round(oRecordSet("ytd_gross") - nDedExempt + (nCurGross - nCurExempt) - aTaxable(3), 2)
'            Else
'                aTaxable(0) = Round((oRecordSet("ytd_gross") + (oRecordSet("ytd_gross_sa") - oRSet2("suntot"))) - nDedExempt + (nCurGross - nCurExempt) - aTaxable(3), 2)
'            End If

            aTaxable(0) = Round(oRecordSet("ytd_gross") - nDedExempt + (nCurGross - nCurExempt) - aTaxable(3), 2)
            
            aTaxable(1) = Round(aTaxable(0) - oRecordSet("amt_exempt") - oRecordSet("add_ex_amt"), 2)
            aTaxable(2) = Round((aTaxable(1) - oRecordSet("range1")) * (oRecordSet("percent") / 100), 2)
            
            If lUpdate Then
                If IfExists("pa87263", "(periodid=" & cQuote & cPeriodID & cQuote & ") and (empid=" & cQuote & cEmpID & cQuote & ") and (dedid='006')") Then
                    cSqlStmt = "update pa87263 set ded_amt = " & Round(oRecordSet("tax_amt") + aTaxable(2) - oRecordSet("ytd_wtax"), 2) & "," & _
                               "                   ded_amt2 = " & Round(oRecordSet("tax_amt") + aTaxable(2) - oRecordSet("ytd_wtax"), 2) & "," & _
                               "                   ded_amt3 = " & Round(nCurExempt + nDedExempt + aTaxable(3), 2) & _
                               " where (periodid=" & cQuote & cPeriodID & cQuote & ") and (empid=" & cQuote & cEmpID & cQuote & ") and (dedid='006')"
                Else
                    cSqlStmt = "INSERT INTO PA87263(PERIODID, EMPID, DEDID, DED_AMT, DED_AMT2, DED_AMT3, COMPUTED)VALUES(" & _
                               cQuote & cPeriodID & cQuote & "," & _
                               cQuote & cEmpID & cQuote & "," & _
                               cQuote & "006" & cQuote & "," & _
                               Round(oRecordSet("tax_amt") + aTaxable(2) - oRecordSet("ytd_wtax"), 2) & "," & _
                               Round(oRecordSet("tax_amt") + aTaxable(2) - oRecordSet("ytd_wtax"), 2) & "," & _
                               Round(nCurExempt + nDedExempt + aTaxable(3), 2) & "," & _
                               "0)"
                End If
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            End If
            
            ComputeTax = Array(Round(oRecordSet("tax_amt") + aTaxable(2) - oRecordSet("ytd_wtax"), 2), Round(nCurExempt + nDedExempt + aTaxable(3), 2))
        Else
            ComputeTax = Array(0#, 0#)
        End If
    End If
    
    Set oRecordSet = Nothing
    Set oRSet2 = Nothing
End Function

'Purpose: Unicode aware MkDir
Public Function CreateDirectory(ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Boolean
   CreateDirectory = CreateDirectoryW(StrPtr("\\?\" & lpPathName), ByVal lpSecurityAttributes) <> 0
End Function

