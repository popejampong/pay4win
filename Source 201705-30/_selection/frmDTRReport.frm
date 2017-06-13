VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDTRReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Manpower Report"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   720
      TabIndex        =   9
      Top             =   1485
      Width           =   375
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
      Left            =   120
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "TXT:PERIODID"
      Top             =   1485
      Width           =   585
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Combine"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1890
      Width           =   2625
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Range"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   855
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Preview"
      Height          =   795
      Left            =   5535
      Picture         =   "frmDTRReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   795
      Left            =   5535
      Picture         =   "frmDTRReport.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1155
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_ISS"
      Top             =   105
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   609
      _Version        =   393216
      Format          =   57081856
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   720
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_ISS"
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   609
      _Version        =   393216
      Format          =   57081856
      CurrentDate     =   38623
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   135
      TabIndex        =   11
      Top             =   1230
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1140
      TabIndex        =   10
      Top             =   1515
      Width           =   4005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   5925
      Left            =   5205
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmDTRReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmDTRReport
' programmer    :   _-=[ srm ]=-_
' date          :   15 May 2009

Option Explicit
    Dim oTempADO As New ADODB.Recordset

Private Sub Check1_Click()
    DTPicker2.Visible = Check1.Value = vbChecked
    Label1.Visible = Check1.Value = vbChecked
End Sub

Private Sub Command11_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    ShowProgress 0
    frmLookup.showPopup 3
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "select empid,concat(lastname,'. ',firstname,' ',left(mname,1),'.') as fullname from di3670 where empid = " & cQuote & cResult & cQuote, objdbRs, False
        Text2.Text = cResult
        Label4.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("fullname"), "")
    End If
End Sub
Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTableName As String
        
        cSqlStmt = " CREATE TABLE tmpDTRD(  [paystatus] integer, " & _
                   " [emp_stat] integer,    [wap] integer," & _
                   " [EMPID] char(6),       [TRAN_NO] char(10), " & _
                   " [FULLNAME] char(100),  [DEPTNAME] char(100), " & _
                   " [DAY_DATE] date,       [DAY_NAME] char(20), " & _
                   " [RegHour] double,      [OTHour] double, " & _
                   " [SAOT] double,         [NDiff] double, " & _
                   " [NDiffOT] double,      [SANDOT] double, " & _
                   " [SUN] double,          [SUNOT] double, " & _
                   " [SUN_ND] double,       [SUN_ND_OT] double, " & _
                   " [LOGDATE] date,        [TRANSDATE] date," & _
                   " [outtrantime] char(15),[intrantime] char(15), " & _
                   " [SHIFTDESC] char(100), [REMARK] char(100)," & _
                   " [TIME1] char(15),      [TIME2] char(15)," & _
                   " [SEQ_NO] integer,      [tag] integer, " & _
                   " [periodid] char(5),    [Duration] char(100)," & _
                   " [TOT_OT] double,       [ND_TOT_OT] double )"

        cTableName = "tmpDTRD"
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM " & cTableName, oTempADO, True
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        aTimeInfo As Variant, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        dLogDate As Date, _
        cParam As String, _
        oRecordSet As New Recordset, _
        oRSet As New Recordset, _
        oRSet2 As New Recordset, _
        oRSet3 As New Recordset, _
        lWap As Boolean, _
        lPclose As Boolean
        
    CreateTemp
    
    ShowProgress 0
    
    cParam = IIf(Check1.Value = 1, cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote, cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote)
    
    cSqlStmt = "select pclose,periodid from pa7730 where date_start between " & cParam
    OpenQueryDNS cSqlStmt, oRSet3, False
    If oRSet3.RecordCount > 0 Then
        While Not oRSet3.EOF
            ShowProgress 2, (oRSet3.AbsolutePosition / oRSet3.RecordCount) * 100
            OpenQueryDNS " select lastname,firstname,mname from di3670 " & _
                         " where empid =" & cQuote & Text2.Text & cQuote, objdbRs, False
        
            aShiftInfo = Array("", "", "", "")
            aTrantype = Array("", "", "", "")
            
            If oRSet3("pclose") = 1 Then
                lPclose = True
            Else
                lPclose = False
            End If
            
            cSqlStmt = " select a.EMPID, concat(a.LASTNAME,', ', a.FIRSTNAME,' ', left(a.MNAME,1),'.') as fullname, a.POSID, " & _
                       " a.EMP_STAT, a.ACTIVE, a.PAYSTATUS, a.DATE_RES, a.WAP, b.PERIODID, b.DATE, c.date_start,c.date_end,duration, " & _
                       " a.DEPID,ifnull(d.linename,'') as linename " & _
                       " from di3670 a " & _
                       " left join " & IIf(lPclose, "dih36770", "di36770") & " b on a.empid=b.empid " & _
                       " left join pa7730 c on b.periodid=c.periodid " & _
                       " left join di5463 d on a.depid=d.lineid " & _
                       " where a.lastname=" & cQuote & objdbRs("lastname") & cQuote & _
                       " and a.firstname=" & cQuote & objdbRs("Firstname") & cQuote & _
                       " and a.emp_stat <> 0 and a.wap=0 and a.paystatus<>2 " & _
                       " and c.periodid = " & cQuote & oRSet3("periodid") & cQuote & _
                       " group by b.periodid " & _
                       " order by c.date_start desc, b.periodid, b.date "
            OpenQueryDNS cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
                While Not oRecordSet.EOF
                    
                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
                    
                    cSqlStmt = " select EMPID, PERIODID, DATE, SHIFTID, DESCRIPTION, TIME1, TIME2, reg_hr, reg_ot_hr, " & _
                               " sa_reg_ot, nd_hr, nd_ot_hr, sa_nd_ot, sun_hr, sun_ot_hr, REMARK, CMPID, allowance, " & _
                               " TAG, sun_nd, sun_nd_ot, Inc_hr, tot_ot, nd_tot_ot from " & IIf(lPclose, "dih36770", "di36770") & _
                               " where empid = " & cQuote & oRecordSet("empid") & cQuote & _
                               " and periodid = " & cQuote & oRecordSet("periodid") & cQuote & _
                               " order by date desc "
                               
                    OpenQueryDNS cSqlStmt, oRSet, False
                    If oRSet.RecordCount > 0 Then
                        While Not oRSet.EOF
                            ShowProgress 2, (oRSet.AbsolutePosition / oRSet.RecordCount) * 100
                            aShiftInfo = Array("", "", "", "")
                            aTrantype = Array("", "", "", "")
                            
                            cSqlStmt = " SELECT a.empid, a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2, " & _
                                       " a.tran_no,a.transdate,date_format(a.transdate,'%a - %b %e, %Y') as `day`,trantype,if(a.trantype=0,'In','Out') as trn_type,a.trantime " & _
                                       " FROM " & IIf(lPclose, "pah84650", "pa84650") & " a " & _
                                       " left join pa74380 b on a.shiftid = b.shiftid " & _
                                       " Where a.empid = " & cQuote & oRecordSet("empid") & cQuote & _
                                       " And a.logdate = " & cQuote & Format(oRSet("date"), "yyyy-mm-dd)") & cQuote & _
                                       " order by a.logdate,a.transdate, a.trantime"
                            OpenQueryDNS cSqlStmt, oRSet2, False
                            If oRSet2.RecordCount > 0 Then
                                While Not oRSet2.EOF
                                
                                    aTrantype(3) = oRSet2("TRANSDATE")
                                    If oRSet2("trantype") = 0 Then
                                        If Trim(aTrantype(1)) = "" Then
                                            aTrantype(0) = oRSet2("trantype")
                                            aTrantype(1) = oRSet2("trantime")
                                            dLogDate = oRSet2("logdate")
                                        End If
                                    Else
                                        aTrantype(0) = oRSet2("trantype")
                                        aTrantype(2) = oRSet2("trantime")
                                        
                                        aShiftInfo(0) = oRSet2("description")
                                        If gCompanyID = "0002" Then
                                            If objdbRs.RecordCount > 0 Then
                                                aShiftInfo(1) = Format(oRSet2("time1"), "hh:mm AMPM")
                                                aShiftInfo(2) = Format(oRSet2("time2"), "hh:mm AMPM")
                                            Else
                                                aShiftInfo(1) = Format(oRSet("time1"), "hh:mm AMPM")
                                                aShiftInfo(2) = Format(oRSet("time2"), "hh:mm AMPM")
                                            End If
                                        Else
                                            aShiftInfo(1) = Format(oRSet("time1"), "hh:mm AMPM")
                                            aShiftInfo(2) = Format(oRSet("time2"), "hh:mm AMPM")
                                        End If
                                        dLogDate = oRSet2("logdate")
                                    End If
                                                                
                                    oRSet2.MoveNext
                                    
                                    If oRecordSet("EMP_STAT") = 1 Then
                                        If oRecordSet("WAP") = 1 Then
                                            lWap = True
                                        Else
                                            lWap = False
                                        End If
                                    Else
                                        lWap = False
                                    End If
                                    
                                    If Not oRSet2.EOF Then
                                        If dLogDate = oRSet2("logdate") Then
                                            If (oRSet2("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                                If Check2 <> 1 Then
                                                    If aTrantype(2) <> "" Then
                                                        If aTrantype(1) <> "" Then
                                                            If oRSet("nd_hr") <> 0 Then
                                                                'for nDIFF
                                                                If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                                    If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                        aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                        aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                        aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                                    End If
                                                                Else
                                                                    If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                        aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 1, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                        aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                        aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                                    End If
                                                                End If
                                                            
                                                            Else
                                                                If oRSet("reg_hr") >= 8 Then
                                                                    'for regular
                                                                    If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                                        If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                            End If
                                                                        Else
                                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                            End If
                                                                        Else
                                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Else
                                                                    If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                                        aTrantype(2) = ""
                                                                        aTrantype(1) = ""
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                           " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                           " LOGDATE,TRANSDATE, " & _
                                                           " intrantime,outtrantime," & _
                                                           " SHIFTDESC,REMARK," & _
                                                           " TIME1,TIME2," & _
                                                           " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                                           cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                                           cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                           oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                                           oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                                           oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                                           oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                           cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                           cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                           cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                                           cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                           oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                                           cQuote & oRecordSet("periodid") & cQuote & "," & _
                                                           cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
                                                QueryTemp cSqlStmt, objdbRs, True
                                                    
                                                aTrantype = Array("", "", "", "")
                                            End If
                                        Else
                                            If Check2 <> 1 Then
                                                If aTrantype(2) <> "" Then
                                                    If aTrantype(1) <> "" Then
                                                        If oRSet("nd_hr") <> 0 Then
                                                            'for nDIFF
                                                            If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                                If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                    aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                                End If
                                                            Else
                                                                If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 1, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                    aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                                End If
                                                            End If
                                                        
                                                        Else
                                                            
                                                            If oRSet("reg_hr") >= 8 Then
                                                                'for regular
                                                                If lWap Then
                                                                    aTrantype(2) = Hour(DateAdd("h", 0, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", aTimeInfo(1), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                Else
                                                                    If Minute(Format(Hour(aShiftInfo(1)) & ":" & Minute(aTrantype(1)), "hh:mm AMPM")) > 5 Then
            '                                                                MsgBox Hour(aTrantype(1)) + 1
                                                                        If Hour(aTrantype(1)) + 1 > Hour(aShiftInfo(1)) Then
                                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr") + 1, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                        Else
                                                                            If Hour(aTrantype(1)) + 1 > Hour(aShiftInfo(1)) - 1 Then
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                            Else
                                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr") - 1, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                    End If
                                                                End If
                                                            Else
                                                                If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                                    aTrantype(2) = ""
                                                                    aTrantype(1) = ""
                                                                End If
                                                            
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                       " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                       " LOGDATE,TRANSDATE, " & _
                                                       " intrantime,outtrantime," & _
                                                       " SHIFTDESC,REMARK," & _
                                                       " TIME1,TIME2," & _
                                                       " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                                       cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                                       cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                       oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                                       oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                                       oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                                       oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                       cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                       cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                       cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                                       cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                       oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                                       cQuote & oRecordSet("periodid") & cQuote & "," & _
                                                       cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
                                            QueryTemp cSqlStmt, objdbRs, True
                                            
                                            aTrantype = Array("", "", "", "")
                                        End If
                                    Else
                                        If Check2 <> 1 Then
                                            If aTrantype(2) <> "" Then
                                                If aTrantype(1) <> "" Then
                                                    If oRSet("nd_hr") <> 0 Then
                                                        'for nDIFF
                                                        If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                            If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                            End If
                                                        Else
                                                            If aTrantype(2) > DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("nd_hr") + oRSet("reg_ot_hr") + 1, Hour(DateAdd("h", 1, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                            End If
                                                        End If
                                                    
                                                    Else
                                                        If oRSet("reg_hr") >= 8 Then
                                                            'for regular
                                                            If lWap Then
                                                                aTrantype(2) = Hour(DateAdd("h", 0, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", aTimeInfo(1), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                            Else
                                                                If Minute(Format(Hour(aShiftInfo(1)) & ":" & Minute(aTrantype(1)), "hh:mm AMPM")) > 5 Then
                                                                    If Hour(aTrantype(1)) + 1 > Hour(aShiftInfo(1)) Then
                                                                        aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr") + 1, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                    Else
                                                                        If Hour(aTrantype(1)) + 1 > Hour(aShiftInfo(1)) - 1 Then
                                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                        Else
                                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr") - 1, Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                        End If
                                                                    End If
                                                                Else
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))) & ":" & left(Minute(DateAdd("h", oRSet("reg_ot_hr"), Format(aShiftInfo(2), "hh:mm AMPM"))), 1) & left(Second(aTrantype(2)), 1)
                                                                End If
                                                            End If
                                                        Else
                                                            If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                                aTrantype(2) = ""
                                                                aTrantype(1) = ""
                                                            End If
                                                        
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                   " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                   " LOGDATE,TRANSDATE, " & _
                                                   " intrantime,outtrantime," & _
                                                   " SHIFTDESC,REMARK," & _
                                                   " TIME1,TIME2," & _
                                                   " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                                   cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                                   cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                                   cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                   oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                                   oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                                   oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                                   oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                                   cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                   cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                                   cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                   oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                                   cQuote & oRecordSet("periodid") & cQuote & "," & _
                                                   cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
                                        QueryTemp cSqlStmt, objdbRs, True
                                        aTrantype = Array("", "", "", "")
                                    End If
                                    
                                Wend
                            End If
                            
                            oRSet.MoveNext
                        Wend
                    End If
                    oRecordSet.MoveNext
                Wend
            End If
            oRSet3.MoveNext
        Wend
        ShowProgress 3
    
        GenerateReport "Daily Time Report", IIf(Check2.Value = 1, "prv376AR2.rpt", "prv376a2.rpt")
    
        ShowProgress 4
    Else
        ShowProgress 4
        
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim cParam As String, _
        nCtr As Integer
    
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    
    Check1_Click
    
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
