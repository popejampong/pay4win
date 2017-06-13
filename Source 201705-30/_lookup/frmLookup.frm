VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLookup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   11580
   ControlBox      =   0   'False
   Icon            =   "frmLookup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10140
      TabIndex        =   5
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   8700
      TabIndex        =   4
      Top             =   7740
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3855
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
      Left            =   1290
      TabIndex        =   0
      Top             =   60
      Width           =   5535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6870
      Left            =   75
      TabIndex        =   6
      Top             =   735
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12118
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   105
      Width           =   1575
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmLookup
' description   :   Module for Lookup
' programmer    :   _-=[ srm ]=-_
' date created  :   04 January 2005

Option Explicit
    Dim oTempADO As ADODB.Recordset
    Dim nLookup As Integer

Private m_cMiddleScroller As New cMiddleButtonScroller

Sub showPopup(ByVal nMode As Integer, Optional ByVal cCondition As String, Optional ByVal lProgress As Boolean = True)
    Dim cTitle As String, cSqlStmt As String, _
        myArray As Variant, nCtr As Integer, _
        nPos As Integer
    
    DoEvents
    
    nLookup = nMode
    
    nPos = 0
    
    Select Case nMode
        Case 1
            nPos = 1
            cTitle = "User Lookup"
            myArray = Array("TXT:[User ID]:10:True", _
                            "TXT:[User Name]:50:True", _
                            "TXT:[SYSUSER]:20:TRUE")
                            
                            
            cSqlStmt = "SELECT USERID, " & _
                       " CONCAT(FIRSTNAME,' ',LASTNAME) AS Username, " & _
                       " IF(SYSUSER = 0,'Non System User','System User') AS SYSUSER FROM PA2360 "
            
            
        Case 2      ' --> Line table
            nPos = 1
            cTitle = "Line Lookup"
            myArray = Array("TXT:[Code]:10:True", _
                            "TXT:[Line Name]:30:True")
            cSqlStmt = "select lineid,linename from di5463"
            
        Case 3      ' --> Employee module... 20050627
            nPos = 1
            cTitle = "Employee Lookup"
              myArray = Array("TXT:[Emp ID]:8:True", _
                            "TXT:[TC ID]:6:True", _
                            "TXT:[Bank Accnt No]:15:True", _
                            "TXT:[Fullname]:28:True", _
                            "TXT:[Department]:20:True", _
                            "TXT:[Position]:15:True", _
                            "TXT:[Status]:13:True", _
                            "TXT:[Contract Description]:35:True", _
                            "TXT:[Contract Remark]:50:True", _
                            "TXT:[Special Remarks]:50:True")
          
            cSqlStmt = " SELECT a.empid, " & _
                       " a.tcid,a.BACCNTNO, concat(a.lastname, ', ', a.firstname) as fullname," & _
                       " ifnull(b.linename,'Undefined Department') as department, c.posname, " & _
                       " if(a.paystatus=2,'Emergency',if(a.emp_stat=0,'WAP',if(a.emp_stat=1,if(a.wap=1,'WAP','Contractual'),if(a.emp_stat=2,'Regular','Tesda')))) as status, " & _
                       " if(active=0, '',concat(if (active=1, 'Resigned', if(active=2, 'Finished' ,'Terminated')), ' as of ', date_format(if(active=1, date_res, if(active = 2, date_fin, date_res)), '%M %d, %Y'))) as remarks, " & _
                       " a.remark, a.s_remark FROM di3670 a " & _
                       " left join di5463 b on a.depid=b.lineid " & _
                       " left join di7670 c on a.posid=c.posid"


        Case 4      ' --> Position table
            nPos = 1
            cTitle = "Position Lookup"
            myArray = Array("TXT:[ID]:10:True", _
                            "TXT:[Description]:50:True")
            cSqlStmt = "select posid, posname from DI7670"
            
            
        Case 5
            nPos = 0
            cTitle = "Period Lookup"
            myArray = Array("TXT:[Period ID]:10:True", _
                            "DAT:[Date Start]:12:True", _
                            "DAT:[Date end]:12:True", _
                            "TXT:[Duration]:25:True", _
                            "TXT:[Period Status]:13:True", _
                            "NUM:[Working days]:13:True", _
                            "NUM:[Holiday]:13:True")
            cSqlStmt = " SELECT periodid,date_start,date_end,duration,if(pclose=1,'Close','Active') as pclose,workindays,holidays FROM pa7730 "
                       
                       
        Case 6
            nPos = 1
            cTitle = "Holiday Lookup"
            myArray = Array("TXT:[Holiday ID]:10:True", _
                            "DAT:[Date]:20:True", _
                            "TXT:[Description]:30:True")
            cSqlStmt = " SELECT holidayid, date_format(pa4329.date,if(fix_day=1,'%M %D','%M %d, %Y')), description FROM pa4329 "
        
        
        Case 7
            nPos = 0
            cTitle = "Deduction Lookup"
            myArray = Array("TXT:[Ded ID]:7:True", _
                            "TXT:[Description]:30:True")
            cSqlStmt = "SELECT DEDID, DEDNAME FROM PA3330"
            
            
        Case 8
            nPos = 0
            cTitle = "Tax Code Lookup"
            myArray = Array("TXT:[Tax ID]:7:True", _
                            "TXT:[Code]:7:True", _
                            "TXT:[Description]:50:True")
            cSqlStmt = "SELECT TAXID, TAXCODE, TAXNAME FROM PA8290"
            
        Case 9
            nPos = 0
            cTitle = "Shift Lookup"
            myArray = Array("TXT:[Shift ID]:10:True", _
                            "TXT:[Description]:30:True", _
                            "TXT:[Start Time]:15:True", _
                            "TXT:[End Time]:15:True", _
                            "TXT:[Remarks]:30:True")
            cSqlStmt = " SELECT SHIFTID,DESCRIPTION,TIME_FORMAT(TIME1,'%h:%i %p'),TIME_FORMAT(TIME2,'%h:%i %p'),REMARK FROM PA74380"
            
        Case 10
            nPos = 0
            cTitle = "Shifting Schedule Lookup"
            myArray = Array("TXT:[Schedule #]:13:True", _
                            "TXT:[Schedule Date]:25:True", _
                            "TXT:[Duration]:40:True", _
                            "TXT:[Department]:30:True")
            cSqlStmt = "select a.sched_no, date_format(a.date_sched,'%M %d, %Y') as date_sched," & _
                       " ifnull(b.duration,'') as duration, ifnull(c.linename,'') as linename" & _
                       " from di546370 a left join pa7730 b on a.periodid=b.periodid" & _
                       " left join di5463 c on a.depid=c.lineid"
                       
        Case 11
            cTitle = "Leave Entry lookup"
            myArray = Array("TXT:[Leave #]:13:True", _
                            "TXT:[Date]:25:True", _
                            "TXT:[Status]:40:True")
            cSqlStmt = "SELECT leave_no, " & _
                       "  date_format(date_leave,'%M %e, %Y') as leave_date, " & _
                       "  if(status=1,concat('Posted as of ',date_format(date_post,'%M %e, %Y')),'') as remark " & _
                       "FROM pa367580"
            
        Case 12
            nPos = 2
            cTitle = "Payroll Entry lookup"
            myArray = Array("TXT:[Emp ID]:8:True", _
                            "TXT:[Bank Account No]:15:True", _
                            "TXT:[Fullname]:40:True", _
                            "TXT:[Position]:30:True", _
                            "TXT:[Department]:30:True", _
                            "TXT:[Status]:13:True", _
                            "TXT:[Remark]:50:True")
            cSqlStmt = "select a.empid, " & _
                       "  d.BACCNTNO, " & _
                       "  a.fullname, " & _
                       "  ifnull(c.posname,'') as position, " & _
                       "  ifnull(b.linename,'') as department, " & _
                       " if(d.paystatus=2,'Emergency',if(d.emp_stat=0,'WAP',if(d.emp_stat=1,if(d.wap=1,'WAP','Contractual'),if(d.emp_stat=2,'Regular','')))) as status, " & _
                       "  if(a.active>0,concat(if(a.active=1,'Resigned',if(a.active=2,'Finished Contract','Terminated')),' as of ',date_format(a.date_res,'%b %e, %Y')),'') as remark " & _
                       "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
                       "  left join di7670 c on a.posid=c.posid " & _
                       "  left join di3670 d on a.empid=d.empid "
                       
        Case 13
            nPos = 1
            cTitle = "Employee Loan lookup"
            myArray = Array("TXT:[Emp ID]:8:True", _
                            "TXT:[Fullname]:40:True", _
                            "TXT:[Position]:25:True", _
                            "TXT:[Department]:25:True", _
                            "TXT:[Remark]:50:True")
            cSqlStmt = "select distinct a.empid, " & _
                       "  concat(b.firstname,' ',if(trim(b.mname)='','',concat(left(b.mname,1),'. ')),b.lastname) as fullname, " & _
                       "  ifnull(d.posname,'') as position, " & _
                       "  ifnull(c.linename,'') as linename, " & _
                       "  if(b.active>0,concat(if(b.active=1,'Resigned ',if(b.active=2,'Finished Contract','Terminated')),' as of ',date_format(if(b.active=1,b.date_res,b.date_fin),'%b %d, %Y')),'') as status " & _
                       "from di3673 a left join di3670 b on a.empid=b.empid " & _
                       "  left join di5463 c on b.depid=c.lineid " & _
                       "  left join di7670 d on b.posid=d.posid"
                       
        Case 14
            nPos = 1
            cTitle = "Custom Shifting Schedule"
            myArray = Array("TXT:[Shift No.]:12:True", _
                            "TXT:[Date Created]:30:True", _
                            "TXT:[Remark]:40:True")
            cSqlStmt = "Select shift_no, date_format(shift_date,'%M %d, %Y') as shift_date, if(status=0,'',concat('Posted as of ',date_format(date_post,'%M %d, %Y'))) as remark " & _
                       "from pa3740"
                       
        Case 15
            cTitle = "Swap Date"
            myArray = Array("TXT:[Ctrl No]:10:True", _
                            "TXT:[Date to Swap]:30:True", _
                            "TXT:[Swap Date]:30:True", _
                            "TXT:[Remark]:50:True")
            cSqlStmt = "select ctrl_no, " & _
                       "       date_format(date1,'%M %d, %Y') as date1," & _
                       "       date_format(date2,'%M %d, %Y') as date2," & _
                       "       if(status=1,concat('Applied as of ',date_format(date_post,'%M %d, %Y')),'') as remark " & _
                       "from pa7927"
        Case 16
            cTitle = "Employee Incentive lookup"
            myArray = Array("TXT:[Incentive #]:18:True", _
                            "TXT:[Date]:30:True", _
                            "TXT:[DURATION]:50:True", _
                            "TXT:[Department]:25:True", _
                            "TXT:[Status]:40:false")
            cSqlStmt = " SELECT a.INC_NO,date_format(a.INC_DATE,'%M %e, %Y') as inc_date, " & _
                       " b.duration, c.linename,if(a.status=1,concat('Posted as of ',date_format(date_post,'%M %e, %Y')),'') as remark " & _
                       " FROM pa4620 a " & _
                       " left join pa7730 b on a.periodid=b.periodid " & _
                       " left join di5463 c on a.depid=c.lineid "
    
        Case 17 '---> for SSCHECK 20080515
            cTitle = "Employee Incentive lookup"
            myArray = Array("TXT:[SSCheck ID]:13:True", _
                            "TXT:[Date]:30:True", _
                            "TXT:[DURATION]:50:True", _
                            "TXT:[Status]:40:false")
            cSqlStmt = " SELECT a.SSCheckID, date_format(a.SSCheckDATE, '%M %d, %Y') as SSCheckDATE, " & _
                       " b.DURATION, a.status FROM pa7720 a left join pa7730 b on a.periodid=b.periodid "


        Case 18 '---> for employee level 20120125
            cTitle = "Level lookup"
            myArray = Array("TXT:[Level Code]:13:True", _
                            "NUM:[Rate]:13:True", _
                            "NUM:[Cola]:13:True")
            cSqlStmt = " SELECT LVLCode,Rate,Cola FROM pa5380 "

        Case 19 '---> for employee Entry level 20120125
            cTitle = "Employee Level lookup"
            myArray = Array("TXT:[Level Transaction]:13:True", _
                            "TXT:[Level Date]:30:True", _
                            "NUM:[Remarks]:13:True")
            cSqlStmt = " select LVLTran, date_format(LVLDATE, '%M %d, %Y') as LVLDATE, if(status=1,concat('Applied as of ',date_format(date_post,'%M %d, %Y')),'') as remark from PA35380 "

        Case 20 '---> for ERP - Cost Center 20120712
            nPos = 1
            cTitle = "Cost Center lookup"
            myArray = Array("TXT:[Cost Center ID]:30:True", _
                            "TXT:[Description]:100:True")
            cSqlStmt = " SELECT COSTCENTERID, DESCRIPTION FROM pa37722 "

        Case 21 '---> for ERP - Work Center 20120720
            nPos = 1
            cTitle = "Work Center lookup"
            myArray = Array("TXT:[Work Center Code]:20:True", _
                            "TXT:[Work Center Description]:25:True", _
                            "TXT:[Cost Center Code]:20:True", _
                            "TXT:[Cost Center Description]:25:True", _
                            "TXT:[Company Code]:15:True", _
                            "TXT:[Company Name]:25:True")
                            
            cSqlStmt = " select a.WORKCENTERID, a.DESCRIPTION, ifnull(b.COSTCENTERID,'') as COSTCENTERID, ifnull(b.DESCRIPTION,'') as DESCRIPTION," & _
                       " ifnull(c.COMPCODE,'') as COMPCODE, ifnull(c.COMPName,'') as COMPName from pa97722 a " & _
                       " left join pa37722 b on a.costcenterid=b.costcenterid left join pa2660 c on a.compcode=c.compcode "

        Case 22 '---> for ERP - Company 20120831
            nPos = 1
            cTitle = "ERP Company lookup"
            myArray = Array("TXT:[Company ID]:5:True", _
                            "TXT:[Company Name]:100:True")
            cSqlStmt = " SELECT COMPCODE, COMPName FROM pa2660 "

        Case 23 '---> for ODBC COnfig 20131109
            nPos = 1
            cTitle = "ODBC Configuration"
            myArray = Array("TXT:[ID]:5:True", _
                            "TXT:[Data Source Name]:20:True", _
                            "TXT:[Description]:30:False", _
                            "TXT:[Server]:20:False", _
                            "TXT:[User]:10:False", _
                            "TXT:[Password]:20:False", _
                            "TXT:[Database]:20:False")
                            
            cSqlStmt = " SELECT ODBCCODE,DSOURCENAME,DESCRIPTION,ODBCSERVER,ODBCUSER,ODBCPASSWORD,ODBCDATABASE FROM pa66220 "
            
        Case 24      ' --> Employee module... 20050627
            nPos = 4
            cTitle = "Employee Import Lookup"
              myArray = Array("TXT:[Emp ID]:8:True", _
                            "TXT:[SSS NO]:15:True", _
                            "TXT:[TC ID]:6:True", _
                            "TXT:[Bank Accnt No]:15:True", _
                            "TXT:[Fullname]:28:True", _
                            "TXT:[Department]:20:True", _
                            "TXT:[Position]:15:True", _
                            "TXT:[Status]:13:True", _
                            "TXT:[Contract Description]:35:True", _
                            "TXT:[Contract Remark]:50:True", _
                            "TXT:[Special Remarks]:50:True")
          
            cSqlStmt = " SELECT a.empid, " & _
                       " a.ssnum, a.tcid,a.BACCNTNO, concat(a.lastname, ', ', a.firstname) as fullname," & _
                       " ifnull(b.linename,'Undefined Department') as department, c.posname, " & _
                       " if(a.paystatus=2,'Emergency',if(a.emp_stat=0,'WAP',if(a.emp_stat=1,if(a.wap=1,'WAP','Contractual'),if(a.emp_stat=2,'Regular','Tesda')))) as status, " & _
                       " if(active=0, '',concat(if (active=1, 'Resigned', if(active=2, 'Finished' ,'Terminated')), ' as of ', date_format(if(active=1, date_res, if(active = 2, date_fin, date_res)), '%M %d, %Y'))) as remarks, " & _
                       " a.remark, a.s_remark FROM di3670 a " & _
                       " left join di5463 b on a.depid=b.lineid " & _
                       " left join di7670 c on a.posid=c.posid"

        Case 25      ' --> Employee module... 20050627
            nPos = 0
            cTitle = "Employee Import Lookup Based on SSS Number"
              myArray = Array("TXT:[SSS NO]:15:True", _
                            "TXT:[Emp ID]:8:True", _
                            "TXT:[TC ID]:6:True", _
                            "TXT:[Bank Accnt No]:15:True", _
                            "TXT:[Fullname]:28:True", _
                            "TXT:[Department]:20:True", _
                            "TXT:[Position]:15:True", _
                            "TXT:[Status]:13:True", _
                            "TXT:[Contract Description]:35:True", _
                            "TXT:[Contract Remark]:50:True", _
                            "TXT:[Special Remarks]:50:True")
          
            cSqlStmt = " SELECT a.ssnum, a.empid, " & _
                       " a.tcid,a.BACCNTNO, concat(a.lastname, ', ', a.firstname) as fullname," & _
                       " ifnull(b.linename,'Undefined Department') as department, c.posname, " & _
                       " if(a.paystatus=2,'Emergency',if(a.emp_stat=0,'WAP',if(a.emp_stat=1,if(a.wap=1,'WAP','Contractual'),if(a.emp_stat=2,'Regular','Tesda')))) as status, " & _
                       " if(active=0, '',concat(if (active=1, 'Resigned', if(active=2, 'Finished' ,'Terminated')), ' as of ', date_format(if(active=1, date_res, if(active = 2, date_fin, date_res)), '%M %d, %Y'))) as remarks, " & _
                       " a.remark, a.s_remark FROM di3670 a " & _
                       " left join di5463 b on a.depid=b.lineid " & _
                       " left join di7670 c on a.posid=c.posid"
                       
        Case 26      ' --> Employee module... 20050627
            nPos = 0
            cTitle = "Employee Import Lookup Based on SSS Number"
              myArray = Array("TXT:[BLKID]:8:True", _
                              "TXT:[Emp ID]:8:True", _
                              "TXT:[SSS NO]:15:True", _
                              "TXT:[Fullname]:28:True", _
                              "TXT:[Department]:20:True", _
                              "TXT:[Position]:15:True", _
                              "TXT:[Status]:13:True", _
                              "TXT:[Contract Description]:35:True")
          
            cSqlStmt = " select a.blkid,a.EMPID,a.SSNUM,concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'. ') as fullname, " & _
                       " ifnull(c.linename,'') as linename,ifnull(d.posname,'') as posname,if(b.EMP_STAT=0,'Wap',if(b.EMP_STAT=1,'Conmtractual','Regular')) as emp_stat, " & _
                       " concat(if(b.ACTIVE=0,'Active',if(b.ACTIVE=1,'Resigned',if(b.ACTIVE=2,'Finished','Terminated'))), ' - ' , b.DATE_RES ) as date_Term " & _
                       " from PA255578 a left join di3670 b on a.empid=b.empid left join di5463 c on b.depid=c.lineid left join di7670 d on b.posid=d.posid "
    
        Case 27      ' --> Employee Salary Increase 20141129
            nPos = 0
            cTitle = "Employee Salary Increase Lookup"
              myArray = Array("TXT:[Salary Increase]:16:True", _
                              "TXT:[Date Registered]:25:True", _
                              "TXT:[Date Confirm]:25:True", _
                              "NUM:[Adjusted Rate]:13:True", _
                              "TXT:[Remarks]:40:True")
                            
            cSqlStmt = " select SALIN, " & _
                       " date_format(DATEREG,'%M %e, %Y') as DATEREG, " & _
                       " date_format(DATECON,'%M %e, %Y') as DATECON, " & _
                       " RATE_ADJ, REMARK from PA7250 "
    
    End Select
    
    If Trim(cCondition) <> "" Then
        cSqlStmt = cSqlStmt & " " & cCondition
    End If
    
'    MsgBox cSqlStmt
    Script2File cSqlStmt
    
    Me.Caption = cTitle
    OpenQueryDNS cSqlStmt, oTempADO, False
    
    If oTempADO.RecordCount > 0 Then
        QueryAttach oTempADO, MSHFlexGrid1, myArray, lProgress, , , 1
        For nCtr = 1 To MSHFlexGrid1.Cols - 1
            If MSHFlexGrid1.ColWidth(nCtr) > 0 Then Combo1.AddItem MSHFlexGrid1.TextMatrix(0, nCtr)
        Next nCtr
        Combo1.ListIndex = nPos
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub Combo1_Click()
    MSHFlexGrid1.Redraw = False
    MSHFlexGrid1.Col = Combo1.ListIndex + 1
    MSHFlexGrid1.Sort = flexSortGenericAscending
    MSHFlexGrid1.Redraw = True
'    RefreshGrid MSHFlexGrid1, True
    Text1_Change
End Sub

Private Sub Command1_Click()
    cResult = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)
    
    ModalResult = mrOk
    Unload Me
End Sub

Private Sub Command2_Click()
    cResult = ""
    ModalResult = mrCancel
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If oTempADO.RecordCount > 0 Then Command1_Click
End Sub

Private Sub Form_Load()
    Set oTempADO = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTempADO = Nothing
End Sub

Private Sub MSHFlexGrid1_DblClick()
    Command1_Click
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbMiddleButton Then
        With m_cMiddleScroller
            .HorizontalMode = ePixelBased
            .VerticalMode = eLineBased
            .StartMiddleScroll MSHFlexGrid1.hwnd
        End With
    End If
End Sub

Private Sub Text1_Change()
    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        Do While .Row < .Rows - 1 And _
                 UCase(left(.TextMatrix(.Row, Combo1.ListIndex + 1), Len(Trim(Text1.Text)))) <> UCase(Trim(Text1.Text))
            .Row = .Row + 1
        Loop
        If .Row <> .Rows - 1 Then
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols
            .TopRow = .Row
            .RowSel = .Row
            .Refresh
        End If
        .Redraw = True
    End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40, 38
            MSHFlexGrid1.SetFocus
            MSHFlexGrid1.Col = 1
            MSHFlexGrid1.ColSel = MSHFlexGrid1.Cols - MSHFlexGrid1.FixedCols
            MSHFlexGrid1.RowSel = MSHFlexGrid1.Row
    End Select
End Sub
