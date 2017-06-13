VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlphalist 
   Caption         =   "Upload to BIR's Alphalist System Utility"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Upload"
      Height          =   495
      Left            =   6255
      TabIndex        =   5
      Top             =   255
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2625
      TabIndex        =   3
      Top             =   90
      Width           =   3405
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2580
      Width           =   10650
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4590
      Left            =   120
      TabIndex        =   4
      Top             =   3135
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   8096
      _Version        =   393216
      RowHeightMin    =   285
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      GridColor       =   12632256
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmAlphalist"
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
'    MsgBox DetectDBF
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
    
' sched 7.1
    aOtherInfo = Array("", "", "", 0, 0, "", "")

    cSqlStmt = "select sched1,sched2,sched3a,sched3b,sched3c,sched3d,sched3e,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g,sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12 from ALPHA7_1 "
    QueryDBF cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        ShowProgress 0

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

                aOtherInfo(0) = Trim(oRecordSet("sched3b"))
                aOtherInfo(1) = Trim(oRecordSet("sched3c"))
                aOtherInfo(2) = Trim(oRecordSet("sched3a"))

'                UPDATE **201410-09**
                If Trim(Format(oRecordSet("sched3d"), "dd-mm-yyyy")) < ("01/01/" & Year(Now) - 1) Then
                    aOtherInfo(5) = "{^" & (Year(Now) - 1) & "/01/01}"
                Else
                    aOtherInfo(5) = "{^" & Format(Trim(oRecordSet("sched3d")), "dd-mm-yyyy") & "}"
                End If

'                UPDATE **201410-09**
                 If Trim(Format(oRecordSet("sched3e"), "dd-mm-yyyy")) > ("12/31/" & Year(Now) - 1) Then
                    aOtherInfo(6) = "{^" & (Year(Now) - 1) & "/12/31}"

                Else
                    aOtherInfo(6) = "{^" & Format(Trim(oRecordSet("sched3e")), "dd-mm-yyyy") & "}"
                End If

            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^" & (Year(Now) - 1) & "/12/31} and schedule_num=" & cQuote & "D7.1" & cQuote & " order by sequence_num desc"
            Script2File cSqlStmt

            QueryDBF cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherInfo(4) = objdbRs("sequence_num") + 1
            Else
                aOtherInfo(4) = 1
            End If

'           UPDATE **201410-09**
            cSqlStmt = " insert into alphadtl(form_type,employer_tin,employer_branch_code,retrn_period,schedule_num,sequence_num,registered_name,first_name,last_name,middle_name,tin,branch_code,employment_from,employment_to," & _
                       " subs_filing,exmpn_code,actual_amt_wthld,pres_taxable_salaries,pres_taxable_13th_month,pres_tax_wthld,pres_nontax_salaries,pres_nontax_13th_month,pres_nontax_sss_gsis_oth_cont, " & _
                       " over_wthld,amt_wthld_dec,exmpn_amt,tax_due,net_taxable_comp_income,gross_comp_income,pres_nontax_de_minimis,pres_taxable_basic_salary,total_nontax_comp_income,total_taxable_comp_income, " & _
                       " status_code,atc_code,region_num,factor_used,income_payment,prev_taxable_salaries,prev_taxable_13th_month,prev_tax_wthld,prev_nontax_salaries,prev_nontax_13th_month,prev_nontax_sss_gsis_oth_cont,tax_rate,heath_premium,fringe_benefit,monetary_value,prev_nontax_de_minimis,prev_total_nontax_comp_income,prev_taxable_basic_salary,pres_total_comp,prev_pres_total_taxable," & _
                       " pres_total_nontax_comp_income,prev_nontax_gross_comp_income,prev_nontax_basic_smw,prev_nontax_holiday_pay,prev_nontax_overtime_pay,prev_nontax_night_diff,prev_nontax_hazard_pay,pres_nontax_gross_comp_income,pres_nontax_basic_smw_day,pres_nontax_basic_smw_month,pres_nontax_basic_smw_year,pres_nontax_holiday_pay,pres_nontax_overtime_pay,pres_nontax_night_diff,prev_pres_total_comp_income, " & _
                       " pres_nontax_hazard_pay,prev_total_taxable,nontax_basic_sal,tax_basic_sal)values(" & _
                        cQuote & "1604CF" & cQuote & "," & cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
                         cQuote & "0000" & cQuote & "," & "{^" & (Year(Now) - 1) & "/12/31}," & _
                        cQuote & "D7.1" & cQuote & "," & aOtherInfo(4) & "," & _
                        cQuote & EncodeStr2(DecodeStr(cCompany)) & cQuote & "," & cQuote & aOtherInfo(0) & cQuote & "," & _
                        cQuote & aOtherInfo(2) & cQuote & "," & cQuote & aOtherInfo(1) & cQuote & "," & _
                        cQuote & Replace(Trim(oRecordSet("sched2")), "-", "") & cQuote & "," & cQuote & "0000" & cQuote & "," & _
                        aOtherInfo(5) & "," & aOtherInfo(6) & "," & _
                        cQuote & Trim(oRecordSet("sched12")) & cQuote & "," & cQuote & Trim(oRecordSet("sched5a")) & cQuote & "," & _
                        oRecordSet("sched11") & "," & oRecordSet("sched4i") & "," & _
                        oRecordSet("sched4h") & "," & oRecordSet("sched9") & "," & _
                        oRecordSet("sched4e") & "," & oRecordSet("sched4b") & "," & _
                        oRecordSet("sched4d") & "," & oRecordSet("sched10b") & "," & _
                        oRecordSet("sched10a") & "," & oRecordSet("sched5b") & "," & _
                        oRecordSet("sched8") & "," & oRecordSet("sched7") & "," & _
                        oRecordSet("sched4a") & "," & oRecordSet("sched4c") & "," & _
                        oRecordSet("sched4g") & "," & oRecordSet("sched4f") & "," & _
                        oRecordSet("sched4j") & ",'','','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0" & ")"
'            MsgBox cSqlStmt
            Script2File cSqlStmt
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

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            aOtherInfo(0) = Trim(oRecordSet("sched3b"))
            aOtherInfo(1) = Trim(oRecordSet("sched3c"))
            aOtherInfo(2) = Trim(oRecordSet("sched3a"))
            aOtherInfo(5) = "{^" & (Year(Now) - 1) & "/01/01}"
            aOtherInfo(6) = "{^" & (Year(Now) - 1) & "/12/31}"

            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^ " & (Year(Now) - 1) & " /12/31} and schedule_num=" & cQuote & "D7.3" & cQuote & " order by sequence_num desc"
            QueryDBF cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherInfo(4) = objdbRs("sequence_num") + 1
            Else
                aOtherInfo(4) = 1
            End If

            cSqlStmt = " insert into alphadtl(form_type,employer_tin,employer_branch_code,retrn_period,schedule_num,sequence_num,registered_name,first_name,last_name,middle_name,tin,branch_code,employment_from,employment_to," & _
                       " subs_filing,exmpn_code,actual_amt_wthld,pres_taxable_salaries,pres_taxable_13th_month,pres_tax_wthld,pres_nontax_salaries,pres_nontax_13th_month,pres_nontax_sss_gsis_oth_cont, " & _
                       " over_wthld,amt_wthld_dec,exmpn_amt,tax_due,net_taxable_comp_income,gross_comp_income,pres_nontax_de_minimis,pres_taxable_basic_salary,total_nontax_comp_income,total_taxable_comp_income, " & _
                       " status_code,atc_code,region_num,factor_used,income_payment,prev_taxable_salaries,prev_taxable_13th_month,prev_tax_wthld,prev_nontax_salaries,prev_nontax_13th_month,prev_nontax_sss_gsis_oth_cont,tax_rate,heath_premium,fringe_benefit,monetary_value,prev_nontax_de_minimis,prev_total_nontax_comp_income,prev_taxable_basic_salary,pres_total_comp,prev_pres_total_taxable," & _
                       " pres_total_nontax_comp_income,prev_nontax_gross_comp_income,prev_nontax_basic_smw,prev_nontax_holiday_pay,prev_nontax_overtime_pay,prev_nontax_night_diff,prev_nontax_hazard_pay,pres_nontax_gross_comp_income,pres_nontax_basic_smw_day,pres_nontax_basic_smw_month,pres_nontax_basic_smw_year,pres_nontax_holiday_pay,pres_nontax_overtime_pay,pres_nontax_night_diff,prev_pres_total_comp_income, " & _
                       " pres_nontax_hazard_pay,prev_total_taxable,nontax_basic_sal,tax_basic_sal)values(" & _
                        cQuote & "1604CF" & cQuote & "," & cQuote & Replace(gTINNum, "-", "") & cQuote & "," & _
                        cQuote & "0000" & cQuote & "," & "{^" & (Year(Now) - 1) & "/12/31}," & _
                        cQuote & "D7.3" & cQuote & "," & aOtherInfo(4) & "," & _
                        cQuote & EncodeStr2(DecodeStr(cCompany)) & cQuote & "," & cQuote & aOtherInfo(0) & cQuote & "," & _
                        cQuote & aOtherInfo(2) & cQuote & "," & cQuote & aOtherInfo(1) & cQuote & "," & _
                        cQuote & Replace(Trim(oRecordSet("sched2")), "-", "") & cQuote & "," & cQuote & "0000" & cQuote & "," & _
                        aOtherInfo(5) & "," & aOtherInfo(6) & "," & _
                        cQuote & Trim(oRecordSet("sched12")) & cQuote & "," & cQuote & Trim(oRecordSet("sched5a")) & cQuote & "," & _
                        oRecordSet("sched11") & "," & oRecordSet("sched4i") & "," & _
                        oRecordSet("sched4h") & "," & oRecordSet("sched9") & "," & _
                        oRecordSet("sched4e") & "," & oRecordSet("sched4b") & "," & _
                        oRecordSet("sched4d") & "," & oRecordSet("sched10b") & "," & _
                        oRecordSet("sched10a") & "," & oRecordSet("sched5b") & "," & _
                        oRecordSet("sched8") & "," & oRecordSet("sched7") & "," & _
                        oRecordSet("sched4a") & "," & oRecordSet("sched4c") & "," & _
                        oRecordSet("sched4g") & "," & oRecordSet("sched4f") & "," & _
                        oRecordSet("sched4j") & ",'','','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0" & ")"

'            MsgBox cSqlStmt
'            Script2File cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            oRecordSet.MoveNext
        Wend
        ShowProgress 4
    End If

' sched 7.5
    aOtherInfo = Array("", "", "", 0#, "", "", "", 0#, 0#)
    If gCompanyID = "0003" Then
        cSqlStmt = " select sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h, " & _
                   " sched5i,sched5j,sched5k,sched5l,sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v, " & _
                   " sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad,sched5ae,sched5af,sched5ag,sched6,sched6b, " & _
                   " sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12,sched13 from ALPHA7_5 "
    Else
        cSqlStmt = " select sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h, " & _
                   " sched5i,sched5j,sched5k,sched5l,sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v, " & _
                   " sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad,sched5ae,sched5af,sched5ag,sched6,sched6b, " & _
                   " sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12 from ALPHA7_5 "
    End If

'    MsgBox cSqlStmt
    QueryDBF cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        ShowProgress 0

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            aOtherInfo(0) = Trim(oRecordSet("sched3b"))
            aOtherInfo(1) = Trim(oRecordSet("sched3c"))
            aOtherInfo(2) = Trim(oRecordSet("sched3a"))
            
            'pang mico to
            aOtherInfo(5) = "{^" & Format(oRecordSet("sched5o"), "yyyy/mm/dd") & "}"
            aOtherInfo(6) = "{^" & Format(oRecordSet("sched5p"), "yyyy/mm/dd") & "}"

            cSqlStmt = "select sequence_num from alphadtl where retrn_period={^" & Year(Now) - 1 & " /12/31} and schedule_num=" & cQuote & "D7.5" & cQuote & " order by sequence_num desc"
            QueryDBF cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherInfo(4) = objdbRs("sequence_num") + 1
            Else
                aOtherInfo(4) = 1
            End If
            If Trim(oRecordSet("sched4")) = "" Then
                aOtherInfo(3) = "III"
            Else
                aOtherInfo(3) = oRecordSet("sched4")
            End If

            'total non taxable compensation Income
            If gCompanyID = "0003" Then
                aOtherInfo(7) = oRecordSet("SCHED5V") + oRecordSet("SCHED5W") + oRecordSet("SCHED5X") + oRecordSet("SCHED5Y") + oRecordSet("SCHED5Z") + oRecordSet("SCHED5AA") + oRecordSet("SCHED5AB") + oRecordSet("SCHED5AC") + oRecordSet("sched13")
            Else
                aOtherInfo(7) = oRecordSet("SCHED5V") + oRecordSet("SCHED5W") + oRecordSet("SCHED5X") + oRecordSet("SCHED5Y") + oRecordSet("SCHED5Z") + oRecordSet("SCHED5AA") + oRecordSet("SCHED5AB") + oRecordSet("SCHED5AC")
            End If

            aOtherInfo(8) = Round(oRecordSet("SCHED5AD") + oRecordSet("SCHED5AE"), 2)

            cSqlStmt = " insert into alphadtl(form_type,employer_tin,employer_branch_code,retrn_period,schedule_num,sequence_num,registered_name,first_name,last_name,middle_name,tin,branch_code,employment_from,employment_to, " & _
                       " atc_code,status_code,region_num,subs_filing,exmpn_code,factor_used,actual_amt_wthld,income_payment, " & _
                       " pres_taxable_salaries,pres_taxable_13th_month,pres_tax_wthld,pres_nontax_salaries,pres_nontax_13th_month, " & _
                       " prev_taxable_salaries,prev_taxable_13th_month,prev_tax_wthld,prev_nontax_salaries,prev_nontax_13th_month,pres_nontax_sss_gsis_oth_cont,prev_nontax_sss_gsis_oth_cont,tax_rate,over_wthld,amt_wthld_dec,exmpn_amt, " & _
                       " tax_due,heath_premium,fringe_benefit,monetary_value,net_taxable_comp_income,gross_comp_income,prev_nontax_de_minimis,prev_total_nontax_comp_income,prev_taxable_basic_salary,pres_nontax_de_minimis, " & _
                       " pres_taxable_basic_salary,pres_total_comp,prev_pres_total_taxable,pres_total_nontax_comp_income,prev_nontax_gross_comp_income,prev_nontax_basic_smw,prev_nontax_holiday_pay,prev_nontax_overtime_pay,prev_nontax_night_diff,prev_nontax_hazard_pay,pres_nontax_gross_comp_income,pres_nontax_basic_smw_day, " & _
                       " pres_nontax_basic_smw_month,pres_nontax_basic_smw_year,pres_nontax_holiday_pay,pres_nontax_overtime_pay,pres_nontax_night_diff,prev_pres_total_comp_income,pres_nontax_hazard_pay, " & _
                       " total_nontax_comp_income,total_taxable_comp_income , prev_total_taxable, nontax_basic_sal, tax_basic_sal)values(" & _
                        cQuote & "1604CF" & cQuote & "," & cQuote & Replace(gTINNum, "-", "") & cQuote & "," & cQuote & "0000" & cQuote & "," & "{^" & (Year(Now) - 1) & "/12/31}," & cQuote & "D7.5" & cQuote & "," & aOtherInfo(4) & "," & cQuote & EncodeStr2(DecodeStr(cCompany)) & cQuote & "," & cQuote & aOtherInfo(0) & cQuote & "," & cQuote & aOtherInfo(2) & cQuote & "," & cQuote & aOtherInfo(1) & cQuote & "," & cQuote & Replace(Trim(oRecordSet("sched2")), "-", "") & cQuote & "," & cQuote & "0000" & cQuote & "," & _
                        IIf(Trim(oRecordSet("sched5o")) <> "", aOtherInfo(5), "{^" & (Year(Now) - 1) & "/01/01}") & "," & IIf(Trim(oRecordSet("sched5p")) <> "", aOtherInfo(6), "{^" & (Year(Now) - 1) & "/01/01}") & "," & _
                        "'',''," & cQuote & aOtherInfo(3) & cQuote & ",''," & cQuote & Trim(oRecordSet("sched6")) & cQuote & "," & oRecordSet("sched5u") & "," & oRecordSet("sched12") & ",0," & _
                        oRecordSet("sched5ae") & "," & oRecordSet("sched5ad") & "," & oRecordSet("sched10b") & "," & oRecordSet("sched5ac") & "," & oRecordSet("sched5Z") & "," & _
                        "0,0,0,0,0," & oRecordSet("sched5ab") & ",0,0," & oRecordSet("sched11b") & "," & oRecordSet("sched11a") & "," & oRecordSet("sched6b") & "," & _
                        oRecordSet("sched9") & "," & oRecordSet("sched7") & ",0,0," & oRecordSet("sched8") & "," & aOtherInfo(8) & ",0,0,0," & oRecordSet("sched5aa") & "," & _
                        "0,0,0," & Round(aOtherInfo(7), 2) & "0,0,0,0,0,0,0," & oRecordSet("sched5q") & "," & oRecordSet("sched5r") & "," & oRecordSet("sched5s") & "," & _
                        oRecordSet("sched5t") & "," & oRecordSet("sched5v") & "," & oRecordSet("sched5w") & "," & oRecordSet("sched5x") & ",0," & oRecordSet("sched5y") & "," & _
                        "0,0,0,0,0" & ")"
                       

'            MsgBox cSqlStmt
'            Script2File cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            oRecordSet.MoveNext
        Wend
        ShowProgress 4
    End If

    Set oRecordSet = Nothing
    Set oDBFConn = Nothing
    
    MsgBox "Trasport of alphalist to the BIR Software is Done!!!"

End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    If InStr(1, UCase(File1.FileName), ".DBC") > 0 Then
'        MsgBox CheckPath(Dir1.Path) & File1.Filename
        DetectDBF CheckPath(Dir1.Path)
    End If
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        QueryDBF Text1.Text, objdbRs, False
'        MsgBox objdbRs.RecordCount
        MSHFlexGrid1.Clear
        Set MSHFlexGrid1.Recordset = objdbRs
    End If
End Sub


