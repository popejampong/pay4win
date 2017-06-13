Attribute VB_Name = "uTable"
' project name  :   Dong-in Payroll & Time Management System
' module        :   uTable
' programmer    :   _-=[ srm ]=-_
' date          :   7 oct 2005

Option Explicit

Public Sub SaveTable(ByVal cTableName As String, ByVal aTableDef As Variant, Optional ByVal lBackup As Boolean = False)
    Dim aFieldStru As Variant, _
        nCtr As Integer, _
        cSqlStmt As String
    
    If Not IfExists("DI82250", "tbl_name=" & cQuote & UCase(cTableName) & cQuote) Then
        cSqlStmt = "INSERT INTO DI82250(tbl_name,tbl_desc,sysname)VALUES(" & _
                   cQuote & UCase(cTableName) & cQuote & ",'','PAYROLL')"
        OpenQueryDNS cSqlStmt, objdbRs, True
    End If
    
    ShowProgress 0
    
    For nCtr = 0 To UBound(aTableDef)
        aFieldStru = aTableDef(nCtr)
        If UBound(aFieldStru) < 0 Then Exit For
        
        If Not IfExists("DI82253", "tbl_name=" & cQuote & UCase(cTableName) & cQuote & " and fld_name=" & cQuote & UCase(aFieldStru(0)) & cQuote) Then
            
            ShowProgress 2, (nCtr / (UBound(aTableDef)) * 100), , , "saving " & UCase(aFieldStru(0)) & " field to " & cTableName
            
            cSqlStmt = "INSERT INTO DI82253(tbl_name,fld_name,fld_type,fld_null,fld_default,fld_index,seq_no,sysname)values(" & _
                       cQuote & UCase(cTableName) & cQuote & "," & cQuote & UCase(aFieldStru(0)) & cQuote & "," & _
                       cQuote & UCase(aFieldStru(1)) & cQuote & "," & aFieldStru(2) & "," & _
                       cQuote & UCase(aFieldStru(3)) & cQuote & ",0," & _
                       nCtr + 1 & "," & _
                       cQuote & "PAYROLL" & cQuote & ")"
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, True
        Else
            
        End If
        
    Next nCtr
    
    ShowProgress 4
    
End Sub

Public Sub ChkTable()
    Add2List "Please wait checking table integrity..."
    
    ChkStatus "DI82250", chkDI82250
    ChkStatus "DI82253", chkDI82253
    
    ChkStatus "PA7287", chksCounter     ' --> counter table
    ChkStatus "PA28348", chkPA28348     ' --> Payroll audit trail
    ChkStatus "DI28348", chkDI28348     ' --> Dicas and PPC audit trail
    
    ChkStatus "PA2360", chkPA2360       ' --> admin table
    ChkStatus "DI5463", chkDI5463       ' --> line/department table
    ChkStatus "DI3670", chkDI3670       ' --> employee table
    ChkStatus "DI7670", chkDI7670       ' --> employee position
    ChkStatus "DI7673", chkDI7673       ' --> employee position designation (201704-07)(TLC)
    
    ' --> for access right and system menu
    ChkStatus "PA2798", chkPA2798
    ChkStatus "PA6368", chkPA6368
    ChkStatus "PA7668", chkPA7668
    
    ' --> table/record lock in edit mode...
    ChkStatus "PA5625", chkPA5625
    
    ' --> 20051007 - c/o ryan 'oinki-doink' castro
    ChkStatus "PA7730", chkPA7730       ' --> Period Table
    ChkStatus "PA7770", chkPA7770       ' --> SSS Table
    
    ' --> 20051016
    ChkStatus "PA7454", chkPA7454       ' --> Philhealth Table
    ChkStatus "PA4329", chkPA4329       ' --> Holiday
    
    ' --> 20051017
    ChkStatus "PA3330", chkPA3330       ' --> Deduction Table
    
    ChkStatus "DI3673", chkDI3673       ' --> Employee Deduction Table
    
    ' --> 20051023
    ChkStatus "PA87260", chkPA87260     ' --> Transaction for Payroll Processing (days worked & adjustment)
    ChkStatus "PA87263", chkPA87263     ' --> Transaction for Payroll Processing (deduction)
    
    ' --> 20060224
    ChkStatus "PA8290", chkPA8290       ' --> Employee Deduction Table
    ChkStatus "PA8293", chkPA8293       ' --> Employee Deduction Table

    ChkStatus "PA74380", chkPA74380     ' --> Shift Info
    
    ChkStatus "PA84650", chkPA84650     ' --> Daily Attendance Table (DTR)
    
    ChkStatus "log", ChkLog             ' --> temporary log file for other bio-clock
    
    ChkStatus "DI36770", chkDI36770     ' --> Shifting Schedule by Employee
    ChkStatus "DIH36770", chkDI36770    ' --> Shifting Schedule by Employee (20061218)
    ChkStatus "DIHH36770", chkDI36770   ' --> for history (20090926)
    
    ChkStatus "DI36770A", chkDI36770A   ' --> For alternate shift schedule (20162610)
    
    ChkStatus "DI546370", chkDI546370   ' --> Shifting Schedule by Line
    ChkStatus "DI546373", chkDI546373   ' --> Shifting Schedule by Line (detail)
    
    ' --> 20060324 - transaction history file...
    ChkStatus "PAH87260", chkPA87260
    ChkStatus "PAHH87260", chkPA87260   ' --> for history (20090926)
    ChkStatus "PAH87263", chkPA87263
    ChkStatus "PAHH87263", chkPA87263   ' --> for history (20090926)
    
    ' --> 20060324 - DTR history file...
    ChkStatus "PAH84650", chkPA84650
    ChkStatus "PAHH84650", chkPA84650 ' ---> for history 20090926

    ' --> 20060328
    ChkStatus "PA53283", chkPA53283     ' --> Incentive Leave

    ChkStatus "PA3674", chkPA3674       ' --> employment history
    
    ' --> 20060526 - Employee Leave
    ChkStatus "PA367580", chkPA367580
    ChkStatus "PA367583", chkPA367583
    
    ChkStatus "PA73887", chkPA73887
    
    ChkStatus "PA7927", chkPA7927       ' --> Swap Day...
    
    ChkStatus "PA13667", chkPA13667     ' --> 13th month pay...
    
    ChkStatus "PA4870", chkPA4870       ' --> Annual Withholding Tax...
    
    ChkStatus "PAH4870", chkPAH4870       ' --> Annual Withholding Tax old...
    
  
    ' --> 20070616 - data from bioclock, automation of downloading...
    ChkStatus "att2000", chkAtt2000
    ChkStatus "att2000h", chkAtt2000
    ChkStatus "att2000hh", chkAtt2000      ' for history
    
    ChkStatus "AlarmLog", chkAlarmLog
    ChkStatus "AlarmLogH", chkAlarmLog

    ' --> Shifting Sched by Employee
    ChkStatus "PA3740", chkPA3740       ' --> Header
    ChkStatus "PA3743", chkPA3743       ' --> Detail - Employee
    ChkStatus "PA3747", chkPA3747       ' --> Detail - Shift
    
    ' --> Incentive by Employee         ' --> 20080313
    ChkStatus "PA4620", chkPA4620       ' --> Header
    ChkStatus "PA4623", chkPA4623       ' --> Detail
    
    
    ' --> Shifting Schedule Checking    ' --> 20080313
    ChkStatus "PA7720", chkPA7720       ' --> Header
    ChkStatus "PA7723", chkPA7723       ' --> Detail
    
    ' --> 20120125
    ChkStatus "PA5380", chkPA5380       ' --> Level Table
    
    ' --> Level by Employee             ' --> 20120125
    ChkStatus "PA35380", chkPA35380     ' --> Header"
    ChkStatus "PA35383", chkPA35383     ' --> Detail
    
    'ERP                                ' --> 20120712
    ChkStatus "PA37722", chkPA37722     ' --> Cost Center table
    ChkStatus "PA97722", chkPA97722     ' --> Work Center table
    
    ChkStatus "PA2660", chkPA2660       ' --> Company for ERP
    
    ChkStatus "PA66220", chkPA66220     ' --> ODBC Server Connector
    
    ChkStatus "PA255578", chkPA255578   ' --> Block List Table
    
    ChkStatus "PA7250", chkPA7250       ' --> Salary Increase Table
    ChkStatus "DIH3670", chkDIH3670     ' --> Employee History Table
    
End Sub


Private Sub ChkStatus(ByVal cTblName As String, ByVal aTable As Variant)
'    Add2List "Please wait checking status of " & cTblName
    
    If Not TableExist(cTblName) Then
        CreateTable cTblName, aTable
    Else
        AlterTable cTblName, aTable
    End If

    If Not ((cTblName = "DI82250") Or (cTblName = "DI82253")) Then SaveTable cTblName, aTable
End Sub


' --> create table automatically with the supplied listing from table definition below...
Private Sub CreateTable(ByVal cTblName As String, ByVal aTableDef As Variant)
    Dim aFieldStru As Variant, _
        nCtr As Integer, _
        cSqlStmt As String
        
    DoEvents
    
    cTblName = "`" & cTblName & "`"
    For nCtr = 0 To UBound(aTableDef)
        aFieldStru = aTableDef(nCtr)
        If UBound(aFieldStru) < 0 Then Exit For
        cSqlStmt = cSqlStmt & _
                   "`" & aFieldStru(0) & "` " & _
                   aFieldStru(1) & " " & _
                   IIf(aFieldStru(2) = 0, "NOT NULL ", "") & _
                   IIf(aFieldStru(2) = 0, "DEFAULT " & aFieldStru(3), "") & _
                   IIf(nCtr <> UBound(aTableDef), ",", "")
    Next nCtr
    
    If Trim(cSqlStmt) <> "" Then
        Add2List "Creating " & cTblName
        cSqlStmt = "CREATE TABLE " & cTblName & "(" & cSqlStmt & ")"
'        MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, objdbRs, True
        Script2File cSqlStmt
        Log2Audit "uTable", "Create table " & cTblName
    End If
End Sub


' --> alter table automatically with the supplied listing from table definition below...
Private Sub AlterTable(ByVal cTblName As String, ByVal aTableDef As Variant)
    Dim aFieldStru As Variant, _
        nCtr As Integer, _
        cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset
        
    DoEvents
    
    cTblName = "`" & cTblName & "`"
    OpenQueryDNS "SHOW COLUMNS FROM " & cTblName, oRecordSet, False
    
    For nCtr = 0 To UBound(aTableDef)
        aFieldStru = aTableDef(nCtr)
        
        If UBound(aFieldStru) < 0 Then Exit For
        
        oRecordSet.Requery adAsyncFetch
        oRecordSet.Find "FIELD='" & PadStr(aFieldStru(0), " ", oRecordSet.Fields.Item("FIELD").DefinedSize, PadRight) & "'"
        If Not oRecordSet.EOF Then
            If Trim(UCase(oRecordSet("TYPE"))) <> Trim(UCase(aFieldStru(1))) Then
                cSqlStmt = cSqlStmt & _
                           "MODIFY `" & aFieldStru(0) & "` " & _
                           aFieldStru(1) & " " & _
                           IIf(aFieldStru(2) = 0, "NOT NULL ", "") & _
                           IIf(aFieldStru(2) = 0, "DEFAULT " & aFieldStru(3), "") & _
                           ","
            End If
        Else
            cSqlStmt = cSqlStmt & _
                       "ADD `" & aFieldStru(0) & "` " & _
                       aFieldStru(1) & " " & _
                       IIf(aFieldStru(2) = 0, "NOT NULL ", "") & _
                       IIf(aFieldStru(2) = 0, "DEFAULT " & aFieldStru(3), "") & _
                       ","
        End If
        
    Next nCtr
    
    If Trim(cSqlStmt) <> "" Then
        Add2List "Altering table structure of " & cTblName
        cSqlStmt = "ALTER TABLE " & cTblName & " " & left(cSqlStmt, Len(cSqlStmt) - 1)
        OpenQueryDNS cSqlStmt, objdbRs, True
        Script2File cSqlStmt
        Log2Audit "uTable", "Altering table " & cTblName
    End If
    
'    MsgBox cSqlStmt
    Set oRecordSet = Nothing
End Sub


Public Function chksCounter() As Variant
    chksCounter = Array(Array("wsID", "char(4)", 0, "''"), _
                        Array("ctrID", "char(10)", 0, "''"), _
                        Array("counter", "int(5)", 0, "'0'"), _
                        Array("unused", "char(200)", 0, "''"), _
                        Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA28348() As Variant
    chkPA28348 = Array(Array("WSID", "char(3)", 0, "''"), _
                       Array("USERID", "char(6)", 0, "''"), _
                       Array("USERNAME", "char(50)", 0, "''"), _
                       Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("TIME", "char(10)", 0, "''"), _
                       Array("MODULE", "char(50)", 0, "''"), _
                       Array("ACTIVITY", "char(250)", 0, "''"), _
                       Array("MGR_CODE", "char(6)", 0, "''"), _
                       Array("MGR_NAME", "char(50)", 0, "''"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI28348() As Variant
    chkDI28348 = Array(Array("WSID", "char(3)", 0, "''"), _
                       Array("USERID", "char(6)", 0, "''"), _
                       Array("USERNAME", "char(50)", 0, "''"), _
                       Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("TIME", "char(10)", 0, "''"), _
                       Array("MODULE", "char(50)", 0, "''"), _
                       Array("ACTIVITY", "char(250)", 0, "''"), _
                       Array("MGR_CODE", "char(6)", 0, "''"), _
                       Array("MGR_NAME", "char(50)", 0, "''"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA2360() As Variant
    chkPA2360 = Array(Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("USERID", "char(6)", 0, "''"), _
                      Array("FIRSTNAME", "char(20)", 0, "''"), _
                      Array("MNAME", "char(20)", 0, "''"), _
                      Array("LASTNAME", "char(20)", 0, "''"), _
                      Array("PASSWORD", "char(16)", 0, "''"), _
                      Array("USERLEVEL", "int(1)", 0, "'0'"), _
                      Array("DEPID", "char(3)", 0, "''"), _
                      Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("TIME", "char(10)", 0, "''"), _
                      Array("WSID", "char(3)", 0, "''"), _
                      Array("GROUPID", "int(1)", 0, "'0'"), _
                      Array("DATE_LOG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("SYSUSER", "int(1)", 0, "'0'"), _
                      Array("POSITION", "char(40)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI5463() As Variant
    chkDI5463 = Array(Array("LINEID", "char(3)", 0, "''"), _
                      Array("LINENAME", "char(100)", 0, "''"), _
                      Array("production", "int(1)", 0, "'0'"), _
                      Array("P1000", "int(5)", 0, "'0'"), _
                      Array("P500", "int(5)", 0, "'0'"), _
                      Array("P100", "int(5)", 0, "'0'"), _
                      Array("P50", "int(5)", 0, "'0'"), _
                      Array("P20", "int(5)", 0, "'0'"), _
                      Array("P10", "int(5)", 0, "'0'"), _
                      Array("P5", "int(5)", 0, "'0'"), _
                      Array("P1", "int(5)", 0, "'0'"), _
                      Array("PCOIN", "int(5)", 0, "'0'"), _
                      Array("COSTCENTERID", "char(10)", 0, "''"), _
                      Array("WORKCENTERID", "char(10)", 0, "''"), _
                      Array("ERPPOSCODE", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI3670() As Variant
    chkDI3670 = Array(Array("EMPID", "char(6)", 0, "''"), Array("TCID", "char(5)", 0, "''"), Array("BCID", "char(2)", 0, "''"), Array("BACCNTNO", "char(16)", 0, "''"), _
                      Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("FIRSTNAME", "char(50)", 0, "''"), Array("MNAME", "char(50)", 0, "''"), Array("LASTNAME", "char(50)", 0, "''"), _
                      Array("BIRTHDAY", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("ADD_NO", "char(230)", 0, "''"), Array("ADD_BRGY", "char(230)", 0, "''"), Array("ADD_CITY", "char(230)", 0, "''"), Array("TEL_NUM", "char(15)", 0, "''"), _
                      Array("DEPID", "char(3)", 0, "''"), Array("SHIFTID", "char(5)", 0, "''"), _
                      Array("POSID", "char(3)", 0, "''"), Array("POSITION", "char(40)", 0, "''"), Array("POS_ALLOW", "decimal(18,4)", 0, "0.0000"), _
                      Array("PHEALTHNUM", "char(15)", 0, "''"), Array("PAGIBIGNO", "char(15)", 0, "''"), Array("SSNUM", "char(15)", 0, "''"), Array("TIN", "char(15)", 0, "''"), Array("TAXCODE", "char(15)", 0, "''"), Array("TAXID", "char(3)", 0, "''"), _
                      Array("ISUNION", "int(1)", 0, "'0'"), Array("SEX", "int(1)", 0, "'0'"), Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("EMP_STAT", "int(1)", 0, "'0'"), Array("WAP", "int(1)", 0, "'0'"), Array("ACTIVE", "int(1)", 0, "'1'"), Array("ERP_ACTIVE", "int(1)", 0, "'1'"), _
                      Array("PAYSTATUS", "int(1)", 0, "'0'"), _
                      Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("SSER1215", "decimal(18,4)", 0, "0.0000"), Array("SSPREM1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("PS1215", "decimal(18,4)", 0, "0.0000"), Array("ES1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("MTD_GROSS", "decimal(18,4)", 0, "0.0000"), Array("MTD_BASIC", "decimal(18,4)", 0, "0.0000"), Array("MTD_TAXABLE", "decimal(18,4)", 0, "0.0000"), _
                      Array("YTD_GROSS", "decimal(18,4)", 0, "0.0000"), Array("YTD_GROSS_SA", "decimal(18,4)", 0, "0.0000"), _
                      Array("YTD_BASIC", "decimal(18,4)", 0, "0.0000"), Array("YTD_WTAX", "decimal(18,4)", 0, "0.0000"), Array("YTD_COLA", "decimal(18,4)", 0, "0.0000"), _
                      Array("SL_AVAIL", "int(2)", 0, "'0'"), Array("VL_AVAIL", "int(2)", 0, "'0'"), Array("UL_AVAIL", "int(3)", 0, "'0'"), _
                      Array("SL_USE", "int(2)", 0, "'0'"), Array("VL_USE", "int(2)", 0, "'0'"), Array("UL_USE", "int(3)", 0, "'0'"), _
                      Array("DATE_HIRE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_FIN", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("DATE_RES", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("OLD_RATE", "decimal(18,4)", 0, "0.0000"), Array("OLD_COLA", "decimal(18,4)", 0, "0.0000"), _
                      Array("REF_EMPID", "char(6)", 0, "''"), Array("LVLCode", "char(2)", 0, "''"), _
                      Array("LABORTYPE", "int(1)", 0, "'0'"), Array("COSTCENTERID", "char(16)", 0, "''"), Array("WORKCENTERID", "char(16)", 0, "''"), Array("BEPWORKCENTERID", "char(16)", 0, "''"), Array("REMARK", "char(100)", 0, "''"), Array("S_REMARK", "char(100)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"), Array("TERM", "int(1)", 0, "'0'"), Array("DEPID2", "char(3)", 0, "''"))
                      
                      
End Function

Public Function chkPA2798() As Variant
    chkPA2798 = Array(Array("userID", "char(6)", 0, "''"), _
                      Array("mnuName", "char(100)", 0, "''"), _
                      Array("BIT1", "int(1)", 0, "'0'"), _
                      Array("BIT2", "int(1)", 0, "'0'"), _
                      Array("BIT3", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA6368() As Variant
    chkPA6368 = Array(Array("mnuName", "char(100)", 0, "''"), _
                      Array("Caption", "char(100)", 0, "''"), _
                      Array("MenuID", "int(5)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA7668() As Variant
    chkPA7668 = Array(Array("mnuName", "char(100)", 0, "''"), _
                      Array("Parent", "char(100)", 0, "''"), _
                      Array("Caption", "char(100)", 0, "''"), _
                      Array("Avail", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA5625() As Variant
    chkPA5625 = Array(Array("WSID", "char(3)", 0, "''"), _
                      Array("userid", "char(6)", 0, "''"), _
                      Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("TIME", "char(10)", 0, "''"), _
                      Array("MODULE", "char(50)", 0, "''"), _
                      Array("FIELD", "char(20)", 0, "''"), _
                      Array("VALUE", "char(20)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI7670()
    chkDI7670 = Array(Array("POSID", "char(3)", 0, "''"), _
                      Array("POSNAME", "char(50)", 0, "''"), _
                      Array("STAFF", "int(1)", 0, "'0'"), _
                      Array("ALLOWANCE", "decimal(18,4)", 0, "0.0000"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI7673()
    chkDI7673 = Array(Array("POSID", "char(3)", 0, "''"), _
                      Array("DESIGNATION", "int(1)", 0, "'0'"))
End Function


Public Function chkPA7730() As Variant
    chkPA7730 = Array(Array("PERIODID", "char(5)", 0, "''"), _
                      Array("DATE_START", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_END", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DURATION", "char(50)", 0, "''"), _
                      Array("PCLOSE", "int(1)", 0, "'0'"), _
                      Array("DATE_CLOSE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("WORKINDAYS", "int(2)", 0, "'0'"), _
                      Array("HOLIDAYS", "int(2)", 0, "'0'"), _
                      Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("13month", "int(1)", 0, "'0'"), _
                      Array("wtax", "int(1)", 0, "'0'"), _
                      Array("isprocess", "int(1)", 0, "'0'"), _
                      Array("date_process", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA7770() As Variant
    chkPA7770 = Array(Array("RANGE1", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("RANGE2", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("SALCRED", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("ER_SS", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("EE_SS", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("SS_TOT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("ER_EC", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("ER_TOT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("EE_TOT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CON_TOT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA3330() As Variant
    chkPA3330 = Array(Array("DEDID", "char(3)", 0, "''"), _
                      Array("DEDNAME", "char(50)", 0, "''"), _
                      Array("DEDNAME2", "char(50)", 0, "''"), _
                      Array("SHORT_DESC", "char(10)", 0, "''"), _
                      Array("DEF_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CUT_OFF_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("FIX_DED", "int(1)", 0, "'0'"), _
                      Array("PERIOD1", "int(1)", 0, "'1'"), _
                      Array("PERIOD2", "int(1)", 0, "'1'"), _
                      Array("AUTO_DED", "int(1)", 0, "'0'"), _
                      Array("DEDTAG", "int(1)", 0, "'0'"), _
                      Array("DEDTYPE", "int(1)", 0, "'2'"), _
                      Array("DEDERPID", "char(7)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA7454() As Variant
    chkPA7454 = Array(Array("MSAL_BRAC", "int(1)", 0, "'0'"), _
                      Array("RANGE1", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("RANGE2", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("SAL_BASE", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("MTOT_CONT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("PS", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("ES", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

'Public Function chkPA4329() As Variant
'    chkPA4329 = Array(Array("HOLIDAYID", "char(3)", 0, "''"), _
'                      Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
'                      Array("DESCRIPTION", "char(100)", 0, "''"), _
'                      Array("FIX_DAY", "int(1)", 0, "'0'"), _
'                      Array("WITHPAY", "int(1)", 0, "'1'"), _
'                      Array("TAG", "int(1)", 0, "'0'"), _
'                      Array("TAG1", "int(1)", 0, "'0'"), _
'                      Array("TAG2", "int(1)", 0, "'0'"), _
'                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
'End Function

Public Function chkPA4329() As Variant
    chkPA4329 = Array(Array("HOLIDAYID", "char(3)", 0, "''"), _
                      Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DESCRIPTION", "char(100)", 0, "''"), _
                      Array("FIX_DAY", "int(1)", 0, "'0'"), _
                      Array("WITHPAY", "int(1)", 0, "'1'"), _
                      Array("TAG", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function
Public Function chkDI3673() As Variant
    chkDI3673 = Array(Array("EMPID", "char(6)", 0, "''"), _
                      Array("DEDID", "char(3)", 0, "''"), _
                      Array("DEF_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CUT_OFF_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("PERIOD1", "int(1)", 0, "'1'"), _
                      Array("PERIOD2", "int(1)", 0, "'1'"), _
                      Array("ACC_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("LOAN_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CTRL_NO", "char(10)", 0, "''"), _
                      Array("REF_NO", "char(20)", 0, "''"), _
                      Array("REMARK", "char(100)", 0, "''"), _
                      Array("DATE_GRANT", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_START", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_END", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_FIN", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA87260() As Variant
    chkPA87260 = Array(Array("PERIODID", "char(5)", 0, "''"), Array("PERIOD_STAT", "int(1)", 0, "'1'"), _
                       Array("P_DAY", "decimal(18,4)", 0, "0.0000"), Array("P_HOLIDAY", "decimal(18,4)", 0, "0.0000"), Array("SEQ_NO", "int(4)", 0, "'0'"), _
                       Array("EMPID", "char(6)", 0, "''"), Array("DEPID", "char(3)", 0, "''"), Array("POSID", "char(3)", 0, "''"), Array("TAXID", "char(3)", 0, "''"), Array("ACTIVE", "int(1)", 0, "'0'"), Array("EMP_STAT", "int(1)", 0, "'0'"), Array("PAYSTATUS", "int(1)", 0, "'0'"), Array("WAP", "int(1)", 0, "'0'"), _
                       Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA", "decimal(18,4)", 0, "0.0000"), Array("SUN_COLA", "decimal(18,4)", 0, "0.0000"), Array("COLA1215", "decimal(18,4)", 0, "0.0000"), Array("POS_ALLOW", "decimal(18,4)", 0, "0.0000"), _
                       Array("REG_DAY", "decimal(18,4)", 0, "0.0000"), Array("REG_PAY", "decimal(18,4)", 0, "0.0000"), Array("REG_OT_HR", "decimal(18,4)", 0, "0.0000"), Array("REG_OT_PAY", "decimal(18,4)", 0, "0.0000"), Array("SA_REG_OT", "decimal(18,4)", 0, "0.0000"), Array("SA_REG_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("NDIFF_DAY", "decimal(18,4)", 0, "0.0000"), Array("NDIFF_PAY", "decimal(18,4)", 0, "0.0000"), Array("NDIFF_OT_HR", "decimal(18,4)", 0, "0.0000"), Array("NDIFF_OT_PAY", "decimal(18,4)", 0, "0.0000"), Array("SA_NDIFF_OT", "decimal(18,4)", 0, "0.0000"), Array("SA_NDIFF_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("HOLIDAY", "decimal(18,4)", 0, "0.0000"), Array("HOL_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("SUN_HR", "decimal(18,4)", 0, "0.0000"), Array("SUN_PAY", "decimal(18,4)", 0, "0.0000"), Array("SUN_OT", "decimal(18,4)", 0, "0.0000"), Array("SUN_OT_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("SUN_ND", "decimal(18,4)", 0, "0.0000"), Array("SUN_ND_PAY", "decimal(18,4)", 0, "0.0000"), Array("SUN_ND_OT", "decimal(18,4)", 0, "0.0000"), Array("SUN_ND_OT_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("ADJ_PAY", "decimal(18,4)", 0, "0.0000"), Array("SA_ADJ_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("OTHER_PAY", "decimal(18,4)", 0, "0.0000"), Array("LEAVE_PAY", "decimal(18,4)", 0, "0.0000"), Array("M13PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("GROSS_PAY", "decimal(18,4)", 0, "0.0000"), Array("GROSS16231", "decimal(18,4)", 0, "0.0000"), Array("BASICPAY", "decimal(18,4)", 0, "0.0000"), Array("BASIC1215", "decimal(18,4)", 0, "0.0000"), _
                       Array("NET_PAY", "decimal(18,4)", 0, "0.0000"), Array("SA_NET_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("SSSNUM", "char(15)", 0, "''"), Array("SSER", "decimal(18,4)", 0, "0.0000"), Array("SSPREM", "decimal(18,4)", 0, "0.0000"), Array("SSS01", "decimal(18,4)", 0, "0.0000"), Array("EC001", "decimal(18,4)", 0, "0.0000"), Array("SSER1215", "decimal(18,4)", 0, "0.0000"), Array("SSPREM1215", "decimal(18,4)", 0, "0.0000"), _
                       Array("MEDICARE", "decimal(18,4)", 0, "0.0000"), Array("MEDICARE2", "decimal(18,4)", 0, "0.0000"), Array("MED01", "decimal(18,4)", 0, "0.0000"), Array("PS1215", "decimal(18,4)", 0, "0.0000"), Array("ES1215", "decimal(18,4)", 0, "0.0000"), _
                       Array("TINNUM", "char(15)", 0, "''"), Array("WTAX", "decimal(18,4)", 0, "0.0000"), Array("TAXABLE", "decimal(18,4)", 0, "0.0000"), Array("TAX1215", "decimal(18,4)", 0, "0.0000"), _
                       Array("DED_AMT", "decimal(18,4)", 0, "0.0000"), Array("INC_HR", "decimal(18,4)", 0, "0.0000"), Array("INC_PAY", "decimal(18,4)", 0, "0.0000"), _
                       Array("PAGIBIGNO", "char(15)", 0, "''"), Array("PHEALTHNUM", "char(15)", 0, "''"), _
                       Array("FULLNAME", "char(100)", 0, "''"), Array("FIRSTNAME", "char(25)", 0, "''"), Array("MNAME", "char(25)", 0, "''"), Array("LASTNAME", "char(25)", 0, "''"), _
                       Array("DATE_HIRE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("DATE_RES", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("SP_DAY", "decimal(18,4)", 0, "0.0000"), Array("SP_HOLIDAY", "decimal(18,4)", 0, "0.0000"), Array("SPOT_HOLIDAY", "decimal(18,4)", 0, "0.0000"), Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"), Array("BACCNTNO", "char(16)", 0, "''"), Array("COSTCENTERID", "char(10)", 0, "''"), Array("WORKCENTERID", "char(10)", 0, "''"))
End Function

Public Function chkPA87263() As Variant
    chkPA87263 = Array(Array("PERIODID", "char(5)", 0, "''"), _
                       Array("PERIOD_STAT", "int(1)", 0, "'1'"), _
                       Array("EMPID", "char(6)", 0, "''"), _
                       Array("DEDID", "char(3)", 0, "''"), _
                       Array("CTRL_NO", "char(10)", 0, "''"), _
                       Array("COMPUTED", "int(1)", 0, "'0'"), _
                       Array("DED_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                       Array("DED_AMT2", "decimal(18,4)", 0, "'0.0000'"), _
                       Array("DED_AMT3", "decimal(18,4)", 0, "'0.0000'"))
End Function

Public Function chkPA8290() As Variant
    chkPA8290 = Array(Array("TAXID", "char(3)", 0, "''"), _
                      Array("TAXCODE", "char(5)", 0, "''"), _
                      Array("TAXNAME", "char(100)", 0, "''"))
End Function

Public Function chkPA8293() As Variant
    chkPA8293 = Array(Array("TAXID", "char(3)", 0, "''"), _
                      Array("SEQ_NO", "int(2)", 0, "'0'"), _
                      Array("DED_PCT", "decimal(18,4)", 0, "0.0000"), _
                      Array("DED_AMT", "decimal(18,4)", 0, "0.0000"), _
                      Array("DED_AMT2", "decimal(18,4)", 0, "0.0000"))
End Function


Public Function chkPA74380() As Variant
    chkPA74380 = Array(Array("SHIFTID", "char(5)", 0, "''"), _
                       Array("DESCRIPTION", "char(100)", 0, "''"), _
                       Array("TIME1", "char(10)", 0, "''"), _
                       Array("TIME2", "char(10)", 0, "''"), _
                       Array("REMARK", "char(50)", 0, "''"), _
                       Array("NDIFF", "int(1)", 0, "'0'"), _
                       Array("DEFAULT", "int(1)", 0, "'0'"), _
                       Array("ALLOWANCE", "decimal(18,4)", 0, "'0'"), _
                       Array("REG_HR", "decimal(18,4)", 0, "'0'"), _
                       Array("BTIME", "decimal(18,4)", 0, "'0'"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA84650() As Variant
    chkPA84650 = Array(Array("TRAN_NO", "char(10)", 0, "''"), _
                       Array("BCID", "char(2)", 0, "''"), _
                       Array("TCID", "char(5)", 0, "''"), _
                       Array("EMPID", "char(6)", 0, "''"), _
                       Array("SHIFTID", "char(5)", 0, "''"), _
                       Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("TRANSDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("TRANTIME", "char(10)", 0, "''"), _
                       Array("TRANTYPE", "int(1)", 0, "'0'"), _
                       Array("TAG", "int(1)", 0, "'0'"), _
                       Array("SWAPDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function ChkLog() As Variant
    ChkLog = Array(Array("BCID", "char(2)", 0, "''"), _
                   Array("TCID", "char(5)", 0, "''"), _
                   Array("dat1", "char(7)", 0, "''"), _
                   Array("dat2", "char(15)", 0, "''"), _
                   Array("TRANSDATE", "char(10)", 0, "''"), _
                   Array("TRANTIME", "char(15)", 0, "''"), _
                   Array("TAG", "int(1)", 0, "'0'"))
End Function


Public Function chkDI36770() As Variant
    chkDI36770 = Array(Array("EMPID", "char(6)", 0, "''"), _
                       Array("PERIODID", "char(5)", 0, "''"), _
                       Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("SHIFTID", "char(5)", 0, "''"), _
                       Array("DESCRIPTION", "char(100)", 0, "''"), _
                       Array("TIME1", "char(10)", 0, "''"), Array("TIME2", "char(10)", 0, "''"), Array("allowance", "double", 0, "'5'"), _
                       Array("reg_hr", "double", 0, "'0'"), Array("reg_ot_hr", "double", 0, "'0'"), Array("sa_reg_ot", "double", 0, "'0'"), Array("tot_ot", "double", 0, "'0'"), _
                       Array("nd_hr", "double", 0, "'0'"), Array("nd_ot_hr", "double", 0, "'0'"), Array("sa_nd_ot", "double", 0, "'0'"), Array("nd_tot_ot", "double", 0, "'0'"), _
                       Array("sun_hr", "double", 0, "'0'"), Array("sun_ot_hr", "double", 0, "'0'"), _
                       Array("sun_nd", "double", 0, "'0'"), Array("sun_nd_ot", "double", 0, "'0'"), _
                       Array("Inc_hr", "double", 0, "'0'"), _
                       Array("REMARK", "char(50)", 0, "''"), _
                       Array("TAG", "int(1)", 0, "'0'"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function
'----------------> (201703-02)
Public Function chkDI36770A() As Variant
    chkDI36770A = Array(Array("EMPID", "char(6)", 0, "''"), _
                       Array("PERIODID", "char(5)", 0, "''"), _
                       Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("SHIFTID", "char(5)", 0, "''"), _
                       Array("DESCRIPTION", "char(100)", 0, "''"), _
                       Array("TIME1", "char(10)", 0, "''"), Array("TIME2", "char(10)", 0, "''"), Array("allowance", "double", 0, "'5'"), _
                       Array("reg_hr", "double", 0, "'0'"), Array("reg_ot_hr", "double", 0, "'0'"), Array("sa_reg_ot", "double", 0, "'0'"), Array("tot_ot", "double", 0, "'0'"), _
                       Array("nd_hr", "double", 0, "'0'"), Array("nd_ot_hr", "double", 0, "'0'"), Array("sa_nd_ot", "double", 0, "'0'"), Array("nd_tot_ot", "double", 0, "'0'"), _
                       Array("sun_hr", "double", 0, "'0'"), Array("sun_ot_hr", "double", 0, "'0'"), _
                       Array("sun_nd", "double", 0, "'0'"), Array("sun_nd_ot", "double", 0, "'0'"), _
                       Array("Inc_hr", "double", 0, "'0'"), _
                       Array("REMARK", "char(50)", 0, "''"), _
                       Array("TAG", "int(1)", 0, "'0'"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function



Public Function chkDI546370() As Variant
    chkDI546370 = Array(Array("SCHED_NO", "char(10)", 0, "''"), _
                        Array("DATE_SCHED", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("PERIODID", "char(5)", 0, "''"), _
                        Array("DEPID", "char(3)", 0, "''"), _
                        Array("STATUS", "int(1)", 0, "'0'"), _
                        Array("DATE_POST", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkDI546373() As Variant
    chkDI546373 = Array(Array("SCHED_NO", "char(10)", 0, "''"), _
                        Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("SHIFTID", "char(5)", 0, "''"), _
                        Array("DESCRIPTION", "char(100)", 0, "''"), _
                        Array("TIME1", "char(10)", 0, "''"), _
                        Array("TIME2", "char(10)", 0, "''"), _
                        Array("REMARK", "char(50)", 0, "''"), _
                        Array("SEQ_NO", "int(4)", 0, "'0'"), _
                        Array("STATUS", "int(1)", 0, "'0'"), _
                        Array("DATE_POST", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("A_SHIFTID", "char(5)", 0, "''"), _
                        Array("A_DESC", "char(100)", 0, "''"), _
                        Array("A_TIME1", "char(10)", 0, "''"), _
                        Array("A_TIME2", "char(10)", 0, "''"), _
                        Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA53283() As Variant
    chkPA53283 = Array(Array("RANGE1", "int(4)", 0, "'0'"), _
                       Array("RANGE2", "int(4)", 0, "'0'"), _
                       Array("SL", "decimal(18,4)", 0, "'0'"), _
                       Array("VL", "decimal(18,4)", 0, "'0'"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA3674() As Variant
    chkPA3674 = Array(Array("EMPID", "char(6)", 0, "''"), _
                      Array("TIME_HISTORY", "char(10)", 0, "''"), _
                      Array("DATE_HISTORY", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_HIRE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_FIN", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_RES", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("DEPID", "char(3)", 0, "''"), _
                      Array("POSID", "char(3)", 0, "''"), _
                      Array("RATE_AMT", "decimal(18,4)", 0, "'0'"), _
                      Array("ACTIVE", "int(1)", 0, "'0'"), _
                      Array("EMP_STAT", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkDI82250() As Variant
    chkDI82250 = Array(Array("tbl_name", "char(15)", 0, "''"), _
                       Array("tbl_desc", "char(100)", 0, "''"), _
                       Array("sysname", "char(50)", 0, "''"))
End Function

Public Function chkDI82253() As Variant
    chkDI82253 = Array(Array("tbl_name", "char(15)", 0, "''"), _
                       Array("fld_name", "char(15)", 0, "''"), _
                       Array("fld_type", "char(15)", 0, "''"), _
                       Array("fld_null", "int(1)", 0, "'0'"), _
                       Array("fld_default", "char(50)", 0, "''"), _
                       Array("fld_index", "int(2)", 0, "'0'"), _
                       Array("fld_desc", "char(100)", 0, "''"), _
                       Array("seq_no", "int(3)", 0, "'0'"), _
                       Array("sysname", "char(50)", 0, "''"))
End Function


Public Function chkPA367580() As Variant
    chkPA367580 = Array(Array("leave_no", "char(10)", 0, "''"), _
                        Array("date_leave", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("prep_by", "char(6)", 0, "''"), _
                        Array("rec_by", "char(6)", 0, "''"), _
                        Array("chk_by", "char(6)", 0, "''"), _
                        Array("noted_by", "char(6)", 0, "''"), _
                        Array("appr_by", "char(6)", 0, "''"), _
                        Array("status", "int(1)", 0, "'0'"), _
                        Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA367583() As Variant
    chkPA367583 = Array(Array("leave_no", "char(10)", 0, "''"), _
                        Array("date_leave", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("empid", "char(6)", 0, "''"), _
                        Array("sl_avail", "int(2)", 0, "'0'"), _
                        Array("vl_avail", "int(2)", 0, "'0'"), _
                        Array("ul_avail", "int(3)", 0, "'0'"), _
                        Array("leave_cnt", "int(3)", 0, "'0'"), _
                        Array("tag", "int(1)", 0, "'0'"), _
                        Array("paytag", "int(1)", 0, "'0'"), _
                        Array("date_start", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("date_end", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("remark", "char(50)", 0, "''"), _
                        Array("seq_no", "int(3)", 0, "'0'"), _
                        Array("status", "int(1)", 0, "'0'"), _
                        Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                        Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA73887() As Variant
    chkPA73887 = Array(Array("UL_AVAIL", "int(3)", 0, "'0'"), _
                       Array("UL_USE", "int(3)", 0, "'0'"), _
                       Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), _
                       Array("COLA_AMT", "decimal(18,4)", 0, "0.0000"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA7927() As Variant
    chkPA7927 = Array(Array("ctrl_no", "char(5)", 0, "''"), _
                      Array("date1", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("date2", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"))
End Function


Public Function chkPA13667() As Variant
    chkPA13667 = Array(Array("year", "char(4)", 0, "''"), _
                       Array("DEPID", "char(3)", 0, "''"), _
                       Array("day_cnt", "int(1)", 0, "'0'"), Array("reg_ot", "int(1)", 0, "'0'"), Array("sa_ot", "int(1)", 0, "'0'"), _
                       Array("BACCNTNO", "char(16)", 0, "''"), _
                       Array("empid", "char(6)", 0, "''"), Array("emp_stat", "int(1)", 0, "'0'"), _
                       Array("firstname", "char(50)", 0, "''"), Array("mname", "char(50)", 0, "''"), Array("lastname", "char(50)", 0, "''"), Array("fullname", "char(100)", 0, "''"), _
                       Array("rate_amt", "double", 0, "'0'"), Array("cola_amt", "double", 0, "'0'"), Array("pos_allow", "double", 0, "'0'"), _
                       Array("taxcode", "char(10)", 0, "''"), Array("tin", "char(15)", 0, "''"), _
                       Array("gr01", "double", 0, "'0'"), Array("gr02", "double", 0, "'0'"), Array("gr03", "double", 0, "'0'"), Array("gr04", "double", 0, "'0'"), Array("gr05", "double", 0, "'0'"), Array("gr06", "double", 0, "'0'"), Array("gr07", "double", 0, "'0'"), Array("gr08", "double", 0, "'0'"), Array("gr09", "double", 0, "'0'"), Array("gr10", "double", 0, "'0'"), Array("gr11", "double", 0, "'0'"), Array("gr12", "double", 0, "'0'"), Array("gr13", "double", 0, "'0'"), Array("gr14", "double", 0, "'0'"), Array("gr15", "double", 0, "'0'"), Array("gr16", "double", 0, "'0'"), Array("gr17", "double", 0, "'0'"), Array("gr18", "double", 0, "'0'"), Array("gr19", "double", 0, "'0'"), Array("gr20", "double", 0, "'0'"), Array("gr21", "double", 0, "'0'"), Array("gr22", "double", 0, "'0'"), Array("gr23", "double", 0, "'0'"), Array("gr24", "double", 0, "'0'"), _
                       Array("sa01", "double", 0, "'0'"), Array("sa02", "double", 0, "'0'"), Array("sa03", "double", 0, "'0'"), Array("sa04", "double", 0, "'0'"), Array("sa05", "double", 0, "'0'"), Array("sa06", "double", 0, "'0'"), Array("sa07", "double", 0, "'0'"), Array("sa08", "double", 0, "'0'"), Array("sa09", "double", 0, "'0'"), Array("sa10", "double", 0, "'0'"), Array("sa11", "double", 0, "'0'"), Array("sa12", "double", 0, "'0'"), Array("sa13", "double", 0, "'0'"), Array("sa14", "double", 0, "'0'"), Array("sa15", "double", 0, "'0'"), Array("sa16", "double", 0, "'0'"), Array("sa17", "double", 0, "'0'"), Array("sa18", "double", 0, "'0'"), Array("sa19", "double", 0, "'0'"), Array("sa20", "double", 0, "'0'"), Array("sa21", "double", 0, "'0'"), Array("sa22", "double", 0, "'0'"), Array("sa23", "double", 0, "'0'"), Array("sa24", "double", 0, "'0'"), _
                       Array("bas01", "double", 0, "'0'"), Array("bas02", "double", 0, "'0'"), Array("bas03", "double", 0, "'0'"), Array("bas04", "double", 0, "'0'"), Array("bas05", "double", 0, "'0'"), Array("bas06", "double", 0, "'0'"), Array("bas07", "double", 0, "'0'"), Array("bas08", "double", 0, "'0'"), Array("bas09", "double", 0, "'0'"), Array("bas10", "double", 0, "'0'"), Array("bas11", "double", 0, "'0'"), Array("bas12", "double", 0, "'0'"), Array("bas13", "double", 0, "'0'"), Array("bas14", "double", 0, "'0'"), Array("bas15", "double", 0, "'0'"), Array("bas16", "double", 0, "'0'"), Array("bas17", "double", 0, "'0'"), Array("bas18", "double", 0, "'0'"), Array("bas19", "double", 0, "'0'"), Array("bas20", "double", 0, "'0'"), Array("bas21", "double", 0, "'0'"), Array("bas22", "double", 0, "'0'"), Array("bas23", "double", 0, "'0'"), Array("bas24", "double", 0, "'0'"), _
                       Array("totgross", "double", 0, "'0'"), Array("totsa", "double", 0, "'0'"), Array("totbasic", "double", 0, "'0'"), _
                       Array("leave_cnt", "double", 0, "'0'"), Array("leave_pay", "double", 0, "'0'"), _
                       Array("13mopay", "double", 0, "'0'"), _
                       Array("cash_adv", "double", 0, "'0'"), _
                       Array("ytd_gross", "double", 0, "'0'"), Array("ytd_basic", "double", 0, "'0'"), Array("ytd_cola", "double", 0, "'0'"), Array("ytd_gross_sa", "double", 0, "'0'"), _
                       Array("seq_no", "int(4)", 0, "'0'"), _
                       Array("date_process", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("date_hire", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("date_fin", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                       Array("date_res", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"))
End Function


' --> 20070105
Public Function chkPA4870() As Variant
    chkPA4870 = Array(Array("RANGE1", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("RANGE2", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("AMOUNT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("PERCENT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("S_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("H_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("M_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("EX_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

' --> 20090108
Public Function chkPAH4870() As Variant
   chkPAH4870 = Array(Array("RANGE1", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("RANGE2", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("AMOUNT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("PERCENT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("S_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("H_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("M_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("EX_AMT", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


' --> 20070616, automatic downloading of dtr...
Public Function chkAtt2000()
    chkAtt2000 = Array(Array("empid", "char(6)", 0, "''"), _
                       Array("BCID", "char(2)", 0, "''"), _
                       Array("TCID", "char(5)", 0, "''"), _
                       Array("TRANSDATE", "char(10)", 0, "''"), _
                       Array("TRANTIME", "char(15)", 0, "''"), _
                       Array("trantype", "int(1)", 0, "'0'"), _
                       Array("logid", "char(10)", 0, "''"), _
                       Array("TAG", "int(1)", 0, "'0'"))
End Function

' --> error in using bioclock...
Public Function chkAlarmLog()
    chkAlarmLog = Array(Array("id", "char(10)", 0, "''"), _
                       Array("operator", "char(20)", 0, "''"), _
                       Array("EnrollNum", "char(30)", 0, "''"), _
                       Array("TRANSDATE", "char(10)", 0, "''"), _
                       Array("TRANTIME", "char(15)", 0, "''"), _
                       Array("bcname", "char(20)", 0, "''"), _
                       Array("TAG", "int(1)", 0, "'0'"))
End Function


' --> Shifting Sched by Employee
Public Function chkPA3740() As Variant
    chkPA3740 = Array(Array("SHIFT_NO", "char(10)", 0, "''"), _
                      Array("SHIFT_DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("PERIODID", "char(5)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA3743() As Variant
    chkPA3743 = Array(Array("SHIFT_NO", "char(10)", 0, "''"), _
                      Array("SHIFT_DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("EMPID", "char(6)", 0, "''"), _
                      Array("SEQ_NO", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))

End Function

Public Function chkPA3747() As Variant
    chkPA3747 = Array(Array("SHIFT_NO", "char(10)", 0, "''"), _
                      Array("SHIFT_DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("SHIFTID", "char(5)", 0, "''"), _
                      Array("REMARK", "char(50)", 0, "''"), _
                      Array("A_SHIFTID", "char(5)", 0, "''"), _
                      Array("SEQ_NO", "int(1)", 0, "'0'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))

End Function

' --> Employee Incentive
Public Function chkPA4620() As Variant
    chkPA4620 = Array(Array("INC_NO", "char(10)", 0, "''"), _
                      Array("INC_DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("PERIODID", "char(5)", 0, "''"), _
                      Array("DEPID", "char(3)", 0, "''"), _
                      Array("SHIFTID", "char(5)", 0, "''"), _
                      Array("prep_by", "char(6)", 0, "''"), _
                      Array("chk_by", "char(6)", 0, "''"), _
                      Array("noted_by", "char(6)", 0, "''"), _
                      Array("appr_by", "char(6)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA4623() As Variant
    chkPA4623 = Array(Array("INC_NO", "char(10)", 0, "''"), _
                      Array("INC_DATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("empid", "char(6)", 0, "''"), _
                      Array("Inc_hr", "double", 0, "'0'"), _
                      Array("SEQ_NO", "int(1)", 0, "'0'"), _
                      Array("REMARK", "char(50)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA7720() As Variant
    chkPA7720 = Array(Array("SSCheckID", "char(10)", 0, "''"), _
                      Array("SSCheckDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("PERIODID", "char(5)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA7723() As Variant
    chkPA7723 = Array(Array("SSCheckID", "char(10)", 0, "''"), _
                      Array("SSCheckDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("TRAN_NO", "char(10)", 0, "''"), _
                      Array("EMPID", "char(6)", 0, "''"), _
                      Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("TRANSDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("SHIFTID", "char(5)", 0, "''"), _
                      Array("TRANTIME", "char(10)", 0, "''"), _
                      Array("TRANTYPE", "int(1)", 0, "'0'"), _
                      Array("SHIFTID2", "char(5)", 0, "''"), _
                      Array("SEQ_NO", "int(1)", 0, "'0'"), _
                      Array("REMARK", "char(50)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("TAG", "int(1)", 0, "'0'"), _
                      Array("STAG", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA5380() As Variant
    chkPA5380 = Array(Array("LVLCode", "char(2)", 0, "''"), _
                      Array("Rate", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("Cola", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA35380() As Variant
    chkPA35380 = Array(Array("LVLTran", "char(10)", 0, "''"), _
                      Array("LVLDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("prep_by", "char(6)", 0, "''"), _
                      Array("chk_by", "char(6)", 0, "''"), _
                      Array("noted_by", "char(6)", 0, "''"), _
                      Array("appr_by", "char(6)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA35383() As Variant
    chkPA35383 = Array(Array("LVLTran", "char(10)", 0, "''"), _
                      Array("LVLDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("EMPID", "char(6)", 0, "''"), _
                      Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("TRANSDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("LVLCode", "char(10)", 0, "''"), _
                      Array("SEQ_NO", "int(1)", 0, "'0'"), _
                      Array("REMARK", "char(50)", 0, "''"), _
                      Array("status", "int(1)", 0, "'0'"), _
                      Array("TAG", "int(1)", 0, "'0'"), _
                      Array("date_post", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function



'Public Function chkPA37722() As Variant
'    chkPA37722 = Array(Array("CREF_NO", "char(10)", 0, "''"), _
'                      Array("COSTCENTERID", "char(10)", 0, "''"), _
'                      Array("DESCRIPTION", "char(100)", 0, "''"), _
'                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
'End Function

Public Function chkPA37722() As Variant
    chkPA37722 = Array(Array("COSTCENTERID", "char(10)", 0, "''"), _
                      Array("DESCRIPTION", "char(100)", 0, "''"), _
                      Array("COMPCODE", "char(4)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

Public Function chkPA97722() As Variant
    chkPA97722 = Array(Array("WORKCENTERID", "char(10)", 0, "''"), _
                      Array("DESCRIPTION", "char(100)", 0, "''"), _
                      Array("COSTCENTERID", "char(10)", 0, "''"), _
                      Array("COMPCODE", "char(4)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

' ---> Company for ERP 2012-08-31
Public Function chkPA2660() As Variant
    chkPA2660 = Array(Array("COMPCODE", "char(4)", 0, "''"), _
                      Array("COMPName", "char(50)", 0, "''"), _
                      Array("COMPAddress1", "char(50)", 0, "''"), _
                      Array("COMPAddress2", "char(50)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function


Public Function chkPA66220() As Variant
    chkPA66220 = Array(Array("ODBCCODE", "char(4)", 0, "''"), _
                       Array("DSOURCENAME", "char(50)", 0, "''"), _
                       Array("DESCRIPTION", "char(50)", 0, "''"), _
                       Array("ODBCSERVER", "char(50)", 0, "''"), _
                       Array("ODBCUSER", "char(50)", 0, "''"), _
                       Array("ODBCPASSWORD", "char(50)", 0, "''"), _
                       Array("ODBCDATABASE", "char(50)", 0, "''"), _
                       Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

''---> Employee Block List 201311-18
'Public Function chkPA255578() As Variant
'    chkPA255578 = Array(Array("BLKID", "char(6)", 0, "''"), _
'                      Array("EMPID", "char(6)", 0, "''"), _
'                      Array("SSNUM", "char(15)", 0, "''"), _
'                      Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
'                      Array("FIRSTNAME", "char(50)", 0, "''"), Array("MNAME", "char(50)", 0, "''"), Array("LASTNAME", "char(50)", 0, "''"), _
'                      Array("BIRTHDAY", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
'                      Array("ADD_NO", "char(230)", 0, "''"), Array("ADD_BRGY", "char(230)", 0, "''"), Array("ADD_CITY", "char(230)", 0, "''"), Array("TEL_NUM", "char(15)", 0, "''"), _
'                      Array("DEPID", "char(3)", 0, "''"), _
'                      Array("POSID", "char(3)", 0, "''"), _
'                      Array("SEX", "int(1)", 0, "'0'"), Array("STATUS", "int(1)", 0, "'0'"), _
'                      Array("EMP_STAT", "int(1)", 0, "'0'"), Array("WAP", "int(1)", 0, "'0'"), Array("ACTIVE", "int(1)", 0, "'1'"), _
'                      Array("PAYSTATUS", "int(1)", 0, "'0'"), _
'                      Array("DATE_HIRE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
'                      Array("DATE_FIN", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("DATE_RES", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("S_REMARK", "char(100)", 0, "''"), _
'                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
'End Function

'---> Employee Block List 20131118
Public Function chkPA255578() As Variant
    chkPA255578 = Array(Array("BLKID", "char(6)", 0, "''"), _
                      Array("EMPID", "char(6)", 0, "''"), _
                      Array("SSNUM", "char(15)", 0, "''"), _
                      Array("S_REMARK", "char(200)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

'---> Salariy Increase List 20141129
Public Function chkPA7250() As Variant
    chkPA7250 = Array(Array("SALIN", "char(6)", 0, "''"), _
                      Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATECON", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("RATE_ADJ", "decimal(18,4)", 0, "'0.0000'"), _
                      Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), _
                      Array("REMARK", "char(200)", 0, "''"), _
                      Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("DATE_POST", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
End Function

'---> Employee History List 20141129
Public Function chkDIH3670() As Variant
    chkDIH3670 = Array(Array("SALIN", "char(6)", 0, "''"), Array("DATECON", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("DATEBAK", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("EMPID", "char(6)", 0, "''"), Array("TCID", "char(5)", 0, "''"), Array("BCID", "char(2)", 0, "''"), Array("BACCNTNO", "char(16)", 0, "''"), _
                      Array("DATEREG", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("LOGDATE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("FIRSTNAME", "char(50)", 0, "''"), Array("MNAME", "char(50)", 0, "''"), Array("LASTNAME", "char(50)", 0, "''"), _
                      Array("BIRTHDAY", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("ADD_NO", "char(230)", 0, "''"), Array("ADD_BRGY", "char(230)", 0, "''"), Array("ADD_CITY", "char(230)", 0, "''"), Array("TEL_NUM", "char(15)", 0, "''"), _
                      Array("DEPID", "char(3)", 0, "''"), Array("SHIFTID", "char(5)", 0, "''"), _
                      Array("POSID", "char(3)", 0, "''"), Array("POSITION", "char(40)", 0, "''"), Array("POS_ALLOW", "decimal(18,4)", 0, "0.0000"), _
                      Array("PHEALTHNUM", "char(15)", 0, "''"), Array("PAGIBIGNO", "char(15)", 0, "''"), Array("SSNUM", "char(15)", 0, "''"), Array("TIN", "char(15)", 0, "''"), Array("TAXCODE", "char(15)", 0, "''"), Array("TAXID", "char(3)", 0, "''"), _
                      Array("ISUNION", "int(1)", 0, "'0'"), Array("SEX", "int(1)", 0, "'0'"), Array("STATUS", "int(1)", 0, "'0'"), _
                      Array("EMP_STAT", "int(1)", 0, "'0'"), Array("WAP", "int(1)", 0, "'0'"), Array("ACTIVE", "int(1)", 0, "'1'"), Array("ERP_ACTIVE", "int(1)", 0, "'1'"), _
                      Array("PAYSTATUS", "int(1)", 0, "'0'"), _
                      Array("RATE_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA_AMT", "decimal(18,4)", 0, "0.0000"), Array("COLA1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("SSER1215", "decimal(18,4)", 0, "0.0000"), Array("SSPREM1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("PS1215", "decimal(18,4)", 0, "0.0000"), Array("ES1215", "decimal(18,4)", 0, "0.0000"), _
                      Array("MTD_GROSS", "decimal(18,4)", 0, "0.0000"), Array("MTD_BASIC", "decimal(18,4)", 0, "0.0000"), Array("MTD_TAXABLE", "decimal(18,4)", 0, "0.0000"), _
                      Array("YTD_GROSS", "decimal(18,4)", 0, "0.0000"), Array("YTD_GROSS_SA", "decimal(18,4)", 0, "0.0000"), _
                      Array("YTD_BASIC", "decimal(18,4)", 0, "0.0000"), Array("YTD_WTAX", "decimal(18,4)", 0, "0.0000"), Array("YTD_COLA", "decimal(18,4)", 0, "0.0000"), _
                      Array("SL_AVAIL", "int(2)", 0, "'0'"), Array("VL_AVAIL", "int(2)", 0, "'0'"), Array("UL_AVAIL", "int(3)", 0, "'0'"), _
                      Array("SL_USE", "int(2)", 0, "'0'"), Array("VL_USE", "int(2)", 0, "'0'"), Array("UL_USE", "int(3)", 0, "'0'"), Array("DATE_HIRE", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("DATE_FIN", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), Array("DATE_RES", "date", 0, "'" & Format(Now, "yyyy-mm-dd") & "'"), _
                      Array("OLD_RATE", "decimal(18,4)", 0, "0.0000"), Array("OLD_COLA", "decimal(18,4)", 0, "0.0000"), Array("REF_EMPID", "char(6)", 0, "''"), Array("LVLCode", "char(2)", 0, "''"), _
                      Array("LABORTYPE", "int(1)", 0, "'0'"), Array("COSTCENTERID", "char(16)", 0, "''"), Array("WORKCENTERID", "char(16)", 0, "''"), Array("BEPWORKCENTERID", "char(16)", 0, "''"), Array("REMARK", "char(100)", 0, "''"), Array("S_REMARK", "char(100)", 0, "''"), _
                      Array("CMPID", "char(4)", 0, "'" & gCompanyID & "'"))
                      
                      
End Function


