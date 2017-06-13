VERSION 5.00
Begin VB.Form frmAlphalist2 
   Caption         =   "Genarate Alpha"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   3405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Genarate"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2625
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2220
      Width           =   2475
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2475
   End
End
Attribute VB_Name = "frmAlphalist2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDBFConn As New ADODB.Connection

Function DetectDBF(cDBFPath As String) As Boolean
    On Error GoTo ErrDetect
    Dim cString As String
    
    DoEvents

    If oDBFConn.State = adStateOpen Then oDBFConn.Close
    With oDBFConn
        .CursorLocation = adUseClient
        cString = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;" & _
                   "SourceDB=" & cDBFPath & ";" & _
                   "Exclusive=No"
'        MsgBox cString
        .ConnectionString = cString
        .Open
    End With

    DetectDBF = True

    Exit Function

ErrDetect:
    MsgBox "Error retrieving DBF file", vbCritical
    DetectDBF = False
End Function

Sub QueryDBF(ByVal cSqlStmt As String, oADORSet As ADODB.Recordset, ByVal lState As Boolean)
    On Error GoTo ErrQuery
    
    DoEvents
    If Not lState Then
        Set oADORSet = oDBFConn.Execute(cSqlStmt)
    Else
        oDBFConn.Execute (cSqlStmt)
        While oDBFConn.State = adStateExecuting
            DoEvents
        Wend
    End If
    Exit Sub
    
ErrQuery:
    ErrorMsg Err.Number, Err.Description, "Open DBF Query", "uSRM"
End Sub

Private Sub Command1_Click()
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        aOtherInfo As Variant
    
    DetectDBF CheckPath(Dir1.Path)
    
    
' sched 7.1

    aOtherInfo = Array("", "", "", 0, 0, "", "")
    
    cSqlStmt = "select sched1,sched2,sched3a,sched3b,sched3c,sched3d,sched3e,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g,sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12 from ALPHA7_1 "
    
    QueryDBF cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        cSqlStmt = "delete from alphadtl where retrn_period={^2008/12/31} and schedule_num=" & cQuote & "D7.3" & cQuote
        QueryDBF cSqlStmt, objdbRs, True

'        cSqlStmt = "delete from alphadtl where retrn_period={^2006/12/31} and schedule_num='D7.1'"
'        QueryDBF cSqlStmt, objdbRs, True
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
                aOtherInfo(0) = Trim(oRecordSet("sched3b"))
                aOtherInfo(1) = left(Trim(oRecordSet("sched3c")), 1)
                aOtherInfo(2) = Trim(oRecordSet("sched3a"))
                aOtherInfo(5) = "{^" & Trim(oRecordSet("sched3d")) & "}"
                aOtherInfo(6) = "{^" & Trim(oRecordSet("sched3e")) & "}"
            
            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^2008/12/31} and schedule_num=" & cQuote & "D7.1" & cQuote & " order by sequence_num desc"
            QueryDBF cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherInfo(4) = objdbRs("sequence_num") + 1
            Else
                aOtherInfo(4) = 1
            End If
            
            cSqlStmt = "insert into alphadtl(monetary_value,fringe_benefit,heath_premium,status_code,atc_code,branch_code," & _
                       "tin,retrn_period,registered_name,first_name,last_name,middle_name,employment_from,employment_to,employer_tin,employer_branch_code," & _
                       "form_type,schedule_num,sequence_num," & _
                       "tax_rate,prev_nontax_sss_gsis_oth_cont,prev_nontax_13th_month,prev_nontax_salaries,prev_tax_wthld,prev_taxable_13th_month,prev_taxable_salaries," & _
                       "pres_nontax_sss_gsis_oth_cont,pres_nontax_13th_month,pres_nontax_salaries,pres_tax_wthld," & _
                       "actual_amt_wthld,income_payment,pres_taxable_salaries,pres_taxable_13th_month,amt_wthld_dec,over_wthld,exmpn_amt,tax_due)values(" & _
                        "0,0,0,'','','000'," & cQuote & Replace(Trim(oRecordSet("sched2")), "-", "") & cQuote & "," & _
                       "{^2008/12/31}," & cQuote & EncodeStr2(DecodeStr(cCompany)) & cQuote & "," & _
                       cQuote & aOtherInfo(0) & cQuote & "," & _
                       cQuote & aOtherInfo(2) & cQuote & "," & _
                       cQuote & aOtherInfo(1) & cQuote & "," & _
                       aOtherInfo(5) & "," & _
                       aOtherInfo(6) & "," & _
                       cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
                       cQuote & "020" & cQuote & "," & _
                       cQuote & "1604CF" & cQuote & "," & _
                       cQuote & "D7.1" & cQuote & "," & _
                       aOtherInfo(4) & "," & _
                       "0,0,0,0,0,0,0," & oRecordSet("sched4d") & "," & oRecordSet("sched4b") & ",0," & oRecordSet("sched9") & "," & oRecordSet("sched8") & ",0," & _
                       oRecordSet("sched4a") - oRecordSet("sched4d") & ",0," & oRecordSet("sched11") & "," & oRecordSet("sched10b") & "," & oRecordSet("sched5b") & "," & oRecordSet("sched8") & ")"
'            MsgBox cSqlStmt
                       
                       
            QueryDBF cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 4
    End If
' sched 7.3

    aOtherInfo = Array("", "", "", 0, 0, "", "")
    
    cSqlStmt = "select sched1,sched2,sched3a,sched3b,sched3c,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g,sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12 from ALPHA7_3 "
    
    QueryDBF cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        cSqlStmt = "delete from alphadtl where retrn_period={^2007/12/31} and schedule_num=" & cQuote & "D7.3" & cQuote
        QueryDBF cSqlStmt, objdbRs, True

'        cSqlStmt = "delete from alphadtl where retrn_period={^2006/12/31} and schedule_num='D7.1'"
'        QueryDBF cSqlStmt, objdbRs, True
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
                aOtherInfo(0) = Trim(oRecordSet("sched3b"))
                aOtherInfo(1) = left(Trim(oRecordSet("sched3c")), 1)
                aOtherInfo(2) = Trim(oRecordSet("sched3a"))
                aOtherInfo(5) = "{^1975/11/28}"
                aOtherInfo(6) = "{^2008/12/31}"
            
            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^2008/12/31} and schedule_num=" & cQuote & "D7.3" & cQuote & " order by sequence_num desc"
            QueryDBF cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherInfo(4) = objdbRs("sequence_num") + 1
            Else
                aOtherInfo(4) = 1
            End If
            
            cSqlStmt = "insert into alphadtl(monetary_value,fringe_benefit,heath_premium,status_code,atc_code,branch_code," & _
                       "tin,retrn_period,registered_name,first_name,last_name,middle_name,employment_from,employment_to,employer_tin,employer_branch_code," & _
                       "form_type,schedule_num,sequence_num," & _
                       "tax_rate,prev_nontax_sss_gsis_oth_cont,prev_nontax_13th_month,prev_nontax_salaries,prev_tax_wthld,prev_taxable_13th_month,prev_taxable_salaries," & _
                       "pres_nontax_sss_gsis_oth_cont,pres_nontax_13th_month,pres_nontax_salaries,pres_tax_wthld," & _
                       "actual_amt_wthld,income_payment,pres_taxable_salaries,pres_taxable_13th_month,amt_wthld_dec,over_wthld,exmpn_amt,tax_due)values(" & _
                        "0,0,0,'','','000'," & cQuote & Replace(Trim(oRecordSet("sched2")), "-", "") & cQuote & "," & _
                       "{^2008/12/31}," & cQuote & EncodeStr2(DecodeStr(cCompany)) & cQuote & "," & _
                       cQuote & aOtherInfo(0) & cQuote & "," & _
                       cQuote & aOtherInfo(2) & cQuote & "," & _
                       cQuote & aOtherInfo(1) & cQuote & "," & _
                       aOtherInfo(5) & "," & _
                       aOtherInfo(6) & "," & _
                       cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
                       cQuote & "020" & cQuote & "," & _
                       cQuote & "1604CF" & cQuote & "," & _
                       cQuote & "D7.3" & cQuote & "," & _
                       aOtherInfo(4) & "," & _
                       "0,0,0,0,0,0,0," & oRecordSet("sched4d") & "," & oRecordSet("sched4b") & ",0," & oRecordSet("sched9") & "," & oRecordSet("sched8") & ",0," & _
                       oRecordSet("sched4a") - oRecordSet("sched4d") & ",0," & oRecordSet("sched11") & "," & oRecordSet("sched10b") & "," & oRecordSet("sched5b") & "," & oRecordSet("sched8") & ")"
'            MsgBox cSqlStmt
                       
                       
            QueryDBF cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 4
    End If
        
End Sub

'Private Sub Command2_Click()
'    Dim cSqlStmt As String, _
'        oRecordSet As New ADODB.Recordset, _
'        aOtherInfo As Variant
'
''sched1=seq_no
''sched2=tin
''sched3a=lastname
''sched3b=firtname
''sched3c=midlename
''sched4a=gross
''sched4b=13month
''sched4c=de minimis
''sched4d= non tax
''sched4e=salaries
''sched4f=total non taxable
''sched4g=basic
''sched4h=salaries
''sched5a=code
''sched5b=amount
''sched6=premium paid
''sched7=net taxable
''sched8=tax due
'
'
'' sched 7.2
'
'    aOtherInfo = Array("", "", "", 0, 0, "", "")
'
'    cSqlStmt = "select sched1,sched2,sched3a,sched3b,sched3c,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g,sched4h,sched5a,sched5b,sched6,sched7,sched8 from ALPHA7_2 "
'
'    QueryDBF cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'
'        ShowProgress 0
'
'        cSqlStmt = "delete from alphadtl where retrn_period={^2008/12/31} and schedule_num=" & cQuote & "D7.2" & cQuote
'        QueryDBF cSqlStmt, objdbRs, True
'
''        cSqlStmt = "delete from alphadtl where retrn_period={^2006/12/31} and schedule_num='D7.1'"
''        QueryDBF cSqlStmt, objdbRs, True
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
'
'            aOtherInfo(0) = Trim(oRecordSet("sched3b"))
'            aOtherInfo(1) = left(Trim(oRecordSet("sched3c")), 1)
'            aOtherInfo(2) = Trim(oRecordSet("sched3a"))
'            aOtherInfo(5) = "{^1975/11/28}"
'            aOtherInfo(6) = "{^2008/12/31}"
'
'            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^2008/12/31} and schedule_num=" & cQuote & "D7.2" & cQuote & " order by sequence_num desc"
'            QueryDBF cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aOtherInfo(4) = objdbRs("sequence_num") + 1
'            Else
'                aOtherInfo(4) = 1
'            End If
'
'            cSqlStmt = "insert into alphadtl(monetary_value,fringe_benefit,heath_premium,status_code,atc_code,branch_code," & _
'                       "tin,retrn_period,registered_name,first_name,last_name,middle_name,employment_from,employment_to,employer_tin,employer_branch_code," & _
'                       "form_type,schedule_num,sequence_num," & _
'                       "tax_rate,prev_nontax_sss_gsis_oth_cont,prev_nontax_13th_month,prev_nontax_salaries,prev_tax_wthld,prev_taxable_13th_month,prev_taxable_salaries," & _
'                       "pres_nontax_sss_gsis_oth_cont,pres_nontax_13th_month,pres_nontax_salaries,pres_tax_wthld," & _
'                       "actual_amt_wthld,income_payment,pres_taxable_salaries,pres_taxable_13th_month,amt_wthld_dec,over_wthld,exmpn_amt,tax_due)values(" & _
'                       "0,0,0,'','','000'," & cQuote & Replace(oRecordSet("sched2"), "-", "") & cQuote & "," & _
'                       "{^2008/12/31}," & cQuote & cCompany & cQuote & "," & _
'                       cQuote & aOtherInfo(0) & cQuote & "," & _
'                       cQuote & aOtherInfo(2) & cQuote & "," & _
'                       cQuote & aOtherInfo(1) & cQuote & "," & _
'                       aOtherInfo(5) & "," & _
'                       aOtherInfo(6) & "," & _
'                       cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
'                       cQuote & "020" & cQuote & "," & _
'                       cQuote & "1604CF" & cQuote & "," & _
'                       cQuote & "D7.2" & cQuote & "," & _
'                       aOtherInfo(4) & "," & _
'                       "0,0,0,0,0,0,0," & Trim(oRecordSet("sched4d")) & "," & Trim(oRecordSet("sched4b")) & ",0," & "0" & "," & Trim(oRecordSet("sched8")) & ",0," & Trim(oRecordSet("sched4a")) - Trim(oRecordSet("sched4d")) & ",0," & "0" & "," & "0" & "," & Trim(oRecordSet("sched5b")) & "," & Trim(oRecordSet("sched8")) & ")"
'
''            MsgBox cSqlStmt
'            QueryDBF cSqlStmt, objdbRs, True
'
'            oRecordSet.MoveNext
'
'        Wend
'
'        ShowProgress 4
'    End If
'
'
'
''    aOtherInfo = Array("", "", "", 0, 0, "", "")
''
''    cSqlStmt = "select empid, tin, fullname, m13pay, non_tax, grosspay, " & _
''               " ex_amt, tax_due, tax_wheld, adj_tax, over_wheld, taxcode from c_alpha order by fullname"
''    QueryDBF cSqlStmt, oRecordSet, False
''    If oRecordSet.RecordCount > 0 Then
''
''        ShowProgress 0
''
''        cSqlStmt = "delete from alphadtl where retrn_period={^2007/12/31} and schedule_num='D7.2'"
''        QueryDBF cSqlStmt, objdbRs, True
''        While Not oRecordSet.EOF
''
''            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
''
''            cSqlStmt = "select firstname,mname,lastname,if(active>0,if(active=1 and year(date_res)=2007,1,if(active=2 and year(date_fin)=2006,2,0)),0) as active," & _
''                       "if(active>0,if(active=1 and year(date_res)=2007,date_res,if(active=2 and year(date_fin)=2007,date_fin,'1975/11/28')),'1975/11/28') as date_resfc from di3670 where empid=" & cQuote & oRecordSet("empid") & cQuote
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                aOtherInfo(0) = objdbRs("firstname")
''                aOtherInfo(1) = left(objdbRs("mname"), 1)
''                aOtherInfo(2) = objdbRs("lastname")
''                aOtherInfo(3) = objdbRs("active")
''                aOtherInfo(5) = "{^" & Format(objdbRs("date_resfc"), "yyyy/mm/dd") & "}"
''            Else
''                If InStr(1, oRecordSet("fullname"), ",") > 0 Then
''                    aOtherInfo(2) = Trim(left(oRecordSet("fullname"), InStr(1, oRecordSet("fullname"), ",") - 1))
''                    aOtherInfo(1) = Mid(oRecordSet("fullname"), Len(Trim(oRecordSet("fullname"))) - 1, 1)
''                    aOtherInfo(0) = Trim(Mid(oRecordSet("fullname"), InStr(1, oRecordSet("fullname"), ",") + 1, Len(Trim(oRecordSet("fullname"))) - Len(aOtherInfo(2)) - 3))
''                End If
''                aOtherInfo(3) = 0
''                aOtherInfo(5) = "{^1975/11/28}"
''            End If
''
''            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^2007/12/31} and schedule_num='D7.2' order by sequence_num desc"
''            QueryDBF cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                aOtherInfo(4) = objdbRs("sequence_num") + 1
''            Else
''                aOtherInfo(4) = 1
''            End If
''
''            cSqlStmt = "insert into alphadtl(monetary_value,fringe_benefit,heath_premium,status_code,atc_code,branch_code," & _
''                       "tin,retrn_period,registered_name,first_name,last_name,middle_name,employment_from,employment_to,employer_tin,employer_branch_code," & _
''                       "form_type,schedule_num,sequence_num," & _
''                       "tax_rate,prev_nontax_sss_gsis_oth_cont,prev_nontax_13th_month,prev_nontax_salaries,prev_tax_wthld,prev_taxable_13th_month,prev_taxable_salaries," & _
''                       "pres_nontax_sss_gsis_oth_cont,pres_nontax_13th_month,pres_nontax_salaries,pres_tax_wthld," & _
''                       "actual_amt_wthld,income_payment,pres_taxable_salaries,pres_taxable_13th_month,amt_wthld_dec,over_wthld,exmpn_amt,tax_due)values(" & _
''                       "0,0,0,'','','000'," & cQuote & Replace(oRecordSet("tin"), "-", "") & cQuote & "," & _
''                       "{^2007/12/31}," & cQuote & cCompany & cQuote & "," & _
''                       cQuote & aOtherInfo(0) & cQuote & "," & _
''                       cQuote & aOtherInfo(2) & cQuote & "," & _
''                       cQuote & aOtherInfo(1) & cQuote & "," & _
''                       IIf(aOtherInfo(3) = 0, aOtherInfo(5), "{^2006/01/01}") & "," & aOtherInfo(5) & "," & _
''                       cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
''                       cQuote & "020" & cQuote & "," & _
''                       cQuote & "1604CF" & cQuote & "," & _
''                       cQuote & "D7.2" & cQuote & "," & _
''                       aOtherInfo(4) & "," & _
''                       "0,0,0,0,0,0,0," & oRecordSet("non_tax") & "," & oRecordSet("m13pay") & ",0," & oRecordSet("tax_wheld") & "," & oRecordSet("tax_due") & ",0," & oRecordSet("grosspay") - oRecordSet("non_tax") & ",0," & oRecordSet("adj_tax") & "," & oRecordSet("over_wheld") & "," & oRecordSet("ex_amt") & "," & oRecordSet("tax_due") & ")"
''            QueryDBF cSqlStmt, objdbRs, True
''
''            oRecordSet.MoveNext
''
''        Wend
''
''        ShowProgress 4
''
''    End If
'End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    If InStr(1, UCase(File1.FileName), ".DBF") > 0 Then
'        MsgBox CheckPath(Dir1.Path) & File1.Filename
        DetectDBF CheckPath(Dir1.Path)
    End If
End Sub


'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        QueryDBF Text1.Text, objdbRs, False
''        MsgBox objdbRs.RecordCount
'        MSHFlexGrid1.Clear
'        Set MSHFlexGrid1.Recordset = objdbRs
'    End If
'End Sub

