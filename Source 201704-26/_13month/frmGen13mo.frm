VERSION 5.00
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Begin VB.Form frmGen13mo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "13 Month Pay Computation"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   705
      Left            =   3030
      Picture         =   "frmGen13mo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "21"
      Top             =   1665
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View"
      Height          =   705
      Left            =   3030
      Picture         =   "frmGen13mo.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   885
      Width           =   1185
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
      Left            =   2325
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   1740
      Visible         =   0   'False
      Width           =   750
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   450
      Left            =   75
      TabIndex        =   9
      Top             =   615
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   794
      BorderStyle     =   2
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "December Assumption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   180
         TabIndex        =   10
         Top             =   105
         Width           =   2460
      End
   End
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmGen13mo.frx":224C
      Left            =   1545
      List            =   "frmGen13mo.frx":225F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "1"
      Top             =   120
      Width           =   1200
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
      Left            =   1545
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   1920
      Width           =   750
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
      Left            =   1545
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   1620
      Width           =   750
   End
   Begin VB.TextBox Text10 
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
      Left            =   1545
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   1320
      Width           =   750
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Generate"
      Height          =   705
      Left            =   3030
      Picture         =   "frmGen13mo.frx":2281
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   165
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SA OT Hours"
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
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   1965
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reg OT Hours"
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
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1665
      Width           =   1380
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No of days"
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
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   1365
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   2490
      Left            =   0
      Top             =   0
      Width           =   1515
   End
End
Attribute VB_Name = "frmGen13mo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' project name  :   Dong-in Payroll System
' module        :   frmGen13mo
' description   :   module to generate 13 month pay
' programmer    :   _-=[ srm ]=-_
' date          :   09 Dec 2006

Option Explicit
    Dim oTempADO As New ADODB.Recordset, _
        aDateInfo As Variant

Private Sub Command1_Click()
    If Trim(Text3.Text) <> "" Then
        GetUserRights PadStr(frmMain.mnuTransaction.Name, " ", 100, PadRight), gUserID
        frmTransaction.lblPeriod.Caption = Text3.Text
        frmTransaction.lblDuration.Caption = "13th Month Pay " & cmbFlex.Text
        frmTransaction.Show
    Else
        If Command7.Enabled Then Command7_Click
    End If
End Sub

Private Sub Command11_Click()
    Unload Me
End Sub

Private Sub Command7_Click()
    Dim cSqlStmt As String, _
        cField As String, _
        cFldValue As String, _
        cParam As String, _
        cParam2 As String, _
        cString As String, _
        cDepid As String, _
        oRecordSet As New ADODB.Recordset, _
        nRecNo As Integer, nCtr As Integer, nFldCnt As Integer, _
        aOtherInfo As Variant, _
        nDedAmt As Double
        
    Dim ndecr As Integer
    
        
    cSqlStmt = "select * from pa13667 where year=" & cQuote & cmbFlex.Text & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cSqlStmt = "Warning!!!" & vbCrLf & _
                   "13 month pay had been generated..." & vbCrLf & _
                   "Would you like to overwrite?"
        If MsgBox(cSqlStmt, vbExclamation + vbYesNo, "System Advisory") <> vbYes Then Exit Sub
    End If

    ' --> clear transaction first...
    OpenQueryDNS "delete from pa13667 where year=" & cQuote & cmbFlex.Text & cQuote, objdbRs, True
    Script2File "delete from pa13667 where year=" & cQuote & cmbFlex.Text & cQuote
    
    ' --> clear payroll transaction as well...
    OpenQueryDNS "delete from pa87260 where periodid=" & cQuote & Text3.Text & cQuote, objdbRs, True
    Script2File "delete from pa87260 where periodid=" & cQuote & Text3.Text & cQuote

    OpenQueryDNS "delete from pa87263 where periodid=" & cQuote & Text3.Text & cQuote, objdbRs, True
    Script2File "delete from pa87263 where periodid=" & cQuote & Text3.Text & cQuote
    
    cSqlStmt = "select periodid, pclose " & _
              "From PA7730 " & _
              "Where (13month=0) and ((Year(date_start) = " & cmbFlex.Text & ") Or (Year(date_end) = " & cmbFlex.Text & "))" & _
              " and ((month(date_start)<12) or (month(date_end)<12)) order by date_start "

'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        nFldCnt = objdbRs.RecordCount
        cField = ""
        cFldValue = ""
        cParam = ""
        cParam2 = ""
        
        
        If gAgency = 3 Then
            If objdbRs.RecordCount < 22 Then
                ndecr = 22 - objdbRs.RecordCount
            Else
                ndecr = 0
            End If
        Else
            ndecr = 0
        End If
        
        
        While Not objdbRs.EOF
            cString = Chr$(97 + objdbRs.AbsolutePosition)
            'k1 computation 20081206
            cField = cField & _
                     "bas" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                     "gr" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                     "sa" & Format(objdbRs.AbsolutePosition + ndecr, "00") & ","
                     
            If gCompanyID <> "0002" Then
                If gCompanyID <> "0003" Then

                        cParam = cParam & _
                                 " ifnull((" & cString & ".REG_DAY * " & cString & ".RATE_AMT)+(" & cString & ".NDIFF_DAY * " & cString & ".RATE_AMT),0) as ba" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                                 " ifnull((" & cString & ".REG_DAY * " & cString & ".RATE_AMT)+(" & cString & ".NDIFF_DAY * " & cString & ".RATE_AMT),0) as gr" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                                 " ifnull(" & cString & ".sa_net_pay,0) as sa" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & vbCrLf
                Else
                    'Mico lang to...
                    cParam = cParam & _
                             " ifnull(" & cString & ".basicpay,0) as ba" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                             " ifnull(" & cString & ".basicpay + " & cString & ".COLA + " & cString & ".reg_ot_pay + " & cString & ".ndiff_ot_pay ,0) as gr" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                             " ifnull(" & cString & ".sa_net_pay ,0) as sa" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & vbCrLf

                End If
            Else
                'k1 lang to...
                cParam = cParam & _
                         " ifnull(" & cString & ".basicpay,0) as ba" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                         " ifnull(" & cString & ".basicpay + " & cString & ".COLA + " & cString & ".reg_ot_pay + " & cString & ".ndiff_ot_pay ,0) as gr" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & _
                         " ifnull(" & cString & ".sa_net_pay ,0) as sa" & Format(objdbRs.AbsolutePosition + ndecr, "00") & "," & vbCrLf
            End If

'            MsgBox cParam
            cParam2 = cParam2 & " left join " & IIf(objdbRs("pclose") = 1, "pah87260 ", "pa87260 ") & cString & _
                      " on a.empid=" & cString & ".empid and " & cString & ".periodid=" & cQuote & objdbRs("periodid") & cQuote & vbCrLf
            
            objdbRs.MoveNext
        Wend
    End If
    
    ShowProgress 0
    
    cSqlStmt = "select a.date_hire,date_fin,a.date_res,a.empid,a.depid, a.posid, " & _
               " a.firstname, a.mname, a.lastname, " & _
               " concat(a.lastname,', ',a.firstname,if(trim(a.mname)<>'',concat(' ',left(a.mname,1),'.'),'')) as fullname, a.BACCNTNO," & _
               cParam & _
               " a.emp_stat, a.active, " & _
               " a.tin, a.taxcode, a.pos_allow, a.rate_amt, a.cola_amt, " & _
               " a.paystatus, a.sl_avail + a.vl_avail as leave_cnt, " & _
               " (a.sl_avail - a.sl_use) + (a.vl_avail - a.vl_use) as leave_unuse, " & _
               " a.ytd_basic, ytd_gross, ytd_gross_sa, ytd_cola, " & _
               " a.COSTCENTERID, a.WORKCENTERID " & _
               "from di3670 a" & _
               cParam2 & _
               "where (a.paystatus=0) and ((a.emp_stat <> 0) and not ((a.wap=1) and (a.emp_stat=1))) and " & _
               "(((a.ACTIVE=0) and (a.date_hire<" & cQuote & Format("12/01/" & cmbFlex.Text, "yyyy-mm-dd") & cQuote & "))" & _
               " OR ((a.active=1) and (a.date_res between " & cQuote & Format("12/01/" & cmbFlex.Text, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format("12/31/" & cmbFlex.Text, "yyyy-mm-dd") & cQuote & ")) " & _
               " OR ((a.active=2) and (a.date_fin between " & cQuote & Format("12/01/" & cmbFlex.Text, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format("12/31/" & cmbFlex.Text, "yyyy-mm-dd") & cQuote & "))) " & _
               "order by a.depid, a.emp_stat desc, fullname "
    
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        While Not oRecordSet.EOF
        
'            If oRecordSet("empid") = "241588" Then MsgBox "stop"
            If oRecordSet("depid") <> cDepid Then
                cDepid = oRecordSet("depid")
                nRecNo = 0
            End If
            nRecNo = nRecNo + 1
            
            aOtherInfo = Array(0#, 0#, 0#, 0#)
            '   aOtherInfo(0)   Basic Pay
            '   aOtherInfo(1)   Gross Pay
            '   aOtherInfo(2)   SA Pay
            '   aOtherInfo(3)   Incentive Leave
            
            aOtherInfo(3) = Round(oRecordSet("leave_cnt") * oRecordSet("rate_amt"), 2)
            
            cFldValue = ""
            For nCtr = 1 To nFldCnt
                cFldValue = cFldValue & _
                            oRecordSet("ba" & Format(nCtr + ndecr, "00")) & "," & _
                            oRecordSet("gr" & Format(nCtr + ndecr, "00")) & "," & _
                            oRecordSet("sa" & Format(nCtr + ndecr, "00")) & ","
            Next nCtr
            
            If gCompanyID <> "0002" Then
                If gCompanyID <> "0003" Then
                
                        'revise 20111123
                        aOtherInfo(0) = Round(Val(Text10.Text) * oRecordSet("rate_amt"), 2)
                        'aOtherInfo(0) = aOtherInfo(0) * 2
                        aOtherInfo(1) = Round(Val(Text10.Text) * oRecordSet("rate_amt"), 2)
                        'aOtherInfo(1) = aOtherInfo(1) * 2
                        aOtherInfo(2) = Round(Val(Text2.Text) * oRecordSet("rate_amt"), 2)
                        'aOtherInfo(2) = aOtherInfo(2) * 2
                
                Else
            
                    'sa mico lang to
                
                        'revise 2011
                        aOtherInfo(0) = Round(Val(Text10.Text) * oRecordSet("rate_amt"), 2)
                        'aOtherInfo(0) = aOtherInfo(0) * 2
                        aOtherInfo(1) = Round((Val(Text10.Text) * oRecordSet("rate_amt")) + _
                                        (Val(Text10.Text) * oRecordSet("cola_amt")) + _
                                        (Val(Text10.Text) * (Val(Text1.Text) * (oRecordSet("rate_amt") / 8 * 1.25))), 2)
                        'aOtherInfo(1) = aOtherInfo(1) * 2
                        aOtherInfo(2) = Round((Val(Text10.Text) * (Val(Text2.Text) * (oRecordSet("rate_amt") / 8 * 1.25))), 2)
                        'aOtherInfo(2) = aOtherInfo(2) * 2
                End If
            Else
                'K1
                'revise 2011
                aOtherInfo(0) = Round(Val(Text10.Text) * oRecordSet("rate_amt"), 2)
                'aOtherInfo(0) = aOtherInfo(0) * 2
                aOtherInfo(1) = Round((Val(Text10.Text) * oRecordSet("rate_amt")) + _
                                (Val(Text10.Text) * (Val(Text1.Text) * (oRecordSet("rate_amt") / 8 * 1.25))), 2)
                'aOtherInfo(1) = aOtherInfo(1) * 2
                aOtherInfo(2) = Round((Val(Text10.Text) * (Val(Text2.Text) * (oRecordSet("rate_amt") / 8 * 1.25))), 2)
                'aOtherInfo(2) = aOtherInfo(2) * 2

            End If
                        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            'revise 20111123
            cSqlStmt = "insert into pa13667(`year`,day_cnt,reg_ot,sa_ot,depid," & _
                       "BACCNTNO,empid,firstname,mname,lastname,fullname,emp_stat,cola_amt,pos_allow,rate_amt,taxcode,tin," & _
                       cField & " bas23,gr23,sa23,bas24,gr24,sa24," & _
                       "leave_cnt,leave_pay,ytd_gross,ytd_basic,ytd_cola,ytd_gross_sa,date_process,seq_no,date_hire,date_fin,date_res)values(" & _
                       cQuote & cmbFlex.Text & cQuote & "," & Val(Text10.Text) & "," & Val(Text1.Text) & "," & Val(Text2.Text) & "," & _
                       cQuote & oRecordSet("depid") & cQuote & "," & _
                       cQuote & oRecordSet("BACCNTNO") & cQuote & "," & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("fullname") & cQuote & "," & _
                       oRecordSet("emp_stat") & "," & oRecordSet("cola_amt") & "," & oRecordSet("pos_allow") & "," & oRecordSet("rate_amt") & "," & cQuote & oRecordSet("taxcode") & cQuote & "," & _
                       cQuote & oRecordSet("tin") & cQuote & "," & _
                       cFldValue & _
                       Round((aOtherInfo(0) / 2), 2) & "," & Round((aOtherInfo(1) / 2), 2) & "," & Round((aOtherInfo(2) / 2), 2) & "," & _
                       Round((aOtherInfo(0) / 2), 2) & "," & Round((aOtherInfo(1) / 2), 2) & "," & Round((aOtherInfo(2) / 2), 2) & "," & _
                       oRecordSet("leave_cnt") & "," & _
                       aOtherInfo(3) & "," & _
                       oRecordSet("ytd_gross") & "," & oRecordSet("ytd_basic") & "," & oRecordSet("ytd_cola") & "," & oRecordSet("ytd_gross_sa") & "," & _
                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                       nRecNo & "," & _
                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_res"), "yyyy-mm-dd") & cQuote & ")"
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, True

            ' --> end of 13th month transaction
            
            
            ' --> insert into payroll
            aOtherInfo = Array(0#, 0#, 0#, 0#)
            '   aOtherInfo(0)   Basic Pay
            '   aOtherInfo(1)   Gross Pay
            '   aOtherInfo(2)   SA Pay
            '   aOtherInfo(3)   Incentive Leave
        
            For nCtr = 1 To nFldCnt
                aOtherInfo(0) = aOtherInfo(0) + oRecordSet("ba" & Format(nCtr + ndecr, "00"))
'                gCompanyID = "0004"

                If (gCompanyID = "0002") Or (gCompanyID = "0003") Then
                    aOtherInfo(1) = aOtherInfo(1) + oRecordSet("gr" & Format(nCtr + ndecr, "00")) + oRecordSet("sa" & Format(nCtr + ndecr, "00"))
                Else
                    aOtherInfo(1) = aOtherInfo(1) + oRecordSet("ba" & Format(nCtr + ndecr, "00"))
                End If

            Next nCtr
            
            aOtherInfo(2) = 0
            
'            If gCompanyID = "0002" Then
'                aOtherInfo(3) = Round(oRecordSet("leave_cnt") * oRecordSet("rate_amt"), 2)
'            Else
'                aOtherInfo(3) = 0
'            End If

'            ---> Revised 201612-10
            If (gCompanyID = "0002") Or (gCompanyID = "0003") Then
'            ---> For K1 & MICO
                aOtherInfo(3) = Round(oRecordSet("leave_cnt") * oRecordSet("rate_amt"), 2)
            Else
                aOtherInfo(3) = 0
            End If



            
'                   aOtherInfo(3) = Round(oRecordSet("leave_cnt") * oRecordSet("rate_amt"), 2)
            If gCompanyID <> "0002" Then
                If gCompanyID <> "0003" Then
                
                    '2011
                    
'                    aOtherInfo(0) = Round((aOtherInfo(0) / 12) + ((Val(Text10.Text) * oRecordSet("rate_amt")) * 2) / 12, 2)
'                    aOtherInfo(1) = Round((aOtherInfo(1) / 12) + ((Val(Text10.Text) * oRecordSet("rate_amt")) * 2) / 12, 2)
                    
                    aOtherInfo(0) = Round((aOtherInfo(0) / 12) + ((Val(Text10.Text) * oRecordSet("rate_amt"))) / 12, 2)
                    aOtherInfo(1) = Round((aOtherInfo(1) / 12) + ((Val(Text10.Text) * oRecordSet("rate_amt"))) / 12, 2)
                    
                Else
            
                    'sa mico at sa DIDP ito
'                    'revise 2011
''                    aOtherInfo(0) = Round(((aOtherInfo(0) / 12) + (Val(Text10.Text) * oRecordSet("rate_amt") * 2) / 12), 2)
'                    aOtherInfo(0) = Round((aOtherInfo(0) / 12) + (Val(Text10.Text) * oRecordSet("rate_amt") / 12), 2)
'                    aOtherInfo(1) = Round((aOtherInfo(1) + _
'                                    (Val(Text10.Text) * oRecordSet("rate_amt")) + _
'                                    (Val(Text10.Text) * oRecordSet("cola_amt")) + _
'                                    (Val(Text10.Text) * (Val(Text1.Text) * oRecordSet("rate_amt") / 8 * 1.25)) + _
'                                    (Val(Text10.Text) * (Val(Text2.Text) * oRecordSet("rate_amt") / 8 * 1.25))) / 12, 2)

'                   ---> Revised 201612-10
'                   ---> For MICO
                    aOtherInfo(0) = Round((aOtherInfo(0) / 12) + (Val(Text10.Text) * oRecordSet("rate_amt") / 12), 2)
                    aOtherInfo(1) = Round(((aOtherInfo(1) + _
                                    (Val(Text10.Text) * oRecordSet("rate_amt")) + _
                                    (Val(Text10.Text) * oRecordSet("cola_amt")) + _
                                    (Val(Text10.Text) * (Val(Text1.Text) * oRecordSet("rate_amt") / 8 * 1.25)) + _
                                    (Val(Text10.Text) * (Val(Text2.Text) * oRecordSet("rate_amt") / 8 * 1.25))) + _
                                    aOtherInfo(3)) / 12, 2)



    
                End If
            Else
                'K1
                aOtherInfo(0) = Round((aOtherInfo(0) / 12) + (Val(Text10.Text) * oRecordSet("rate_amt") / 12), 2)
                aOtherInfo(1) = Round(((aOtherInfo(1) + _
                                (Val(Text10.Text) * oRecordSet("rate_amt")) + _
                                (Val(Text10.Text) * (Val(Text1.Text) * oRecordSet("rate_amt") / 8 * 1.25)) + _
                                (Val(Text10.Text) * (Val(Text2.Text) * oRecordSet("rate_amt") / 8 * 1.25))) + _
                                aOtherInfo(3)) / 12, 2)
                                
            End If
                       
            '20131207
            cSqlStmt = "INSERT INTO PA87260(PERIODID,SEQ_NO,EMPID,FULLNAME,FIRSTNAME,MNAME,LASTNAME," & _
                       "EMP_STAT,ACTIVE,PAYSTATUS,DEPID,RATE_AMT,POSID," & _
                       "GROSS_PAY,LEAVE_PAY,M13PAY,NET_PAY,BACCNTNO,COSTCENTERID,WORKCENTERID)VALUES(" & _
                       cQuote & Text3.Text & cQuote & "," & _
                       nRecNo & "," & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & oRecordSet("fullname") & cQuote & "," & _
                       cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & oRecordSet("lastname") & cQuote & "," & _
                       oRecordSet("emp_stat") & "," & _
                       oRecordSet("active") & "," & _
                       oRecordSet("paystatus") & "," & _
                       cQuote & oRecordSet("depid") & cQuote & "," & _
                       oRecordSet("rate_amt") & "," & _
                       cQuote & oRecordSet("posid") & cQuote & "," & _
                       Round(IIf((oRecordSet("emp_stat") = 1) Or (g13Month And (oRecordSet("EMP_STAT") = 2)), aOtherInfo(0), aOtherInfo(1)), 2) & "," & _
                       aOtherInfo(2) & "," & _
                       Round(IIf((oRecordSet("emp_stat") = 1) Or (g13Month And (oRecordSet("EMP_STAT") = 2)), aOtherInfo(0), aOtherInfo(1)), 2) & "," & _
                       Round(IIf((oRecordSet("emp_stat") = 1) Or (g13Month And (oRecordSet("EMP_STAT") = 2)), aOtherInfo(0), aOtherInfo(1)), 2) & "," & _
                       cQuote & oRecordSet("BACCNTNO") & cQuote & "," & _
                       cQuote & oRecordSet("COSTCENTERID") & cQuote & "," & _
                       cQuote & oRecordSet("WORKCENTERID") & cQuote & ")"
                       
                       
'            Script2File cSqlStmt
             
            OpenQueryDNS cSqlStmt, objdbRs, True
            
            
            'cash advance
            If Trim(gCashAdvance) <> "" Then
                cSqlStmt = "select def_amt, cut_off_amt, acc_amt, ctrl_no from di3673 " & _
                           " where (empid=" & cQuote & oRecordSet("EMPID") & cQuote & ")" & _
                           " and (dedid=" & cQuote & gCashAdvance & cQuote & ")" & _
                           " and (status=0)"
                OpenQueryDNS cSqlStmt, oTempADO, False
                If oTempADO.RecordCount > 0 Then
                    nDedAmt = oTempADO("cut_off_amt") - oTempADO("acc_amt")
                    cSqlStmt = "INSERT INTO PA87263(PERIODID, EMPID, DEDID, CTRL_NO, DED_AMT)VALUES(" & _
                               cQuote & Text3.Text & cQuote & "," & _
                               cQuote & oRecordSet("empid") & cQuote & "," & _
                               cQuote & gCashAdvance & cQuote & "," & _
                               cQuote & oTempADO("ctrl_no") & cQuote & "," & _
                               nDedAmt & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt

                    cSqlStmt = "update pa87260 set ded_amt=" & nDedAmt & "," & _
                               "net_pay=net_pay-" & nDedAmt & _
                               " where (periodid=" & cQuote & Text3.Text & cQuote & ")" & _
                               " and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt

                    cSqlStmt = "update pa13667 set cash_adv = " & nDedAmt & _
                               " where (`year`=" & cQuote & cmbFlex.Text & cQuote & ")" & _
                               " and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
            End If
            'k1 20091210
            If gCompanyID = "0002" Then
                If Trim(gECA) <> "" Then
                    cSqlStmt = "select def_amt, cut_off_amt, acc_amt, ctrl_no from di3673 " & _
                               " where (empid=" & cQuote & oRecordSet("EMPID") & cQuote & ")" & _
                               " and (dedid=" & cQuote & gECA & cQuote & ")" & _
                               " and (status=0)"
                    OpenQueryDNS cSqlStmt, oTempADO, False
                    If oTempADO.RecordCount > 0 Then
                        nDedAmt = oTempADO("cut_off_amt") - oTempADO("acc_amt")
                        cSqlStmt = "INSERT INTO PA87263(PERIODID, EMPID, DEDID, CTRL_NO, DED_AMT)VALUES(" & _
                                   cQuote & Text3.Text & cQuote & "," & _
                                   cQuote & oRecordSet("empid") & cQuote & "," & _
                                   cQuote & gECA & cQuote & "," & _
                                   cQuote & oTempADO("ctrl_no") & cQuote & "," & _
                                   nDedAmt & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
    
                        cSqlStmt = "update pa87260 set ded_amt= ded_amt + " & nDedAmt & "," & _
                                   "net_pay=net_pay-" & nDedAmt & _
                                   " where (periodid=" & cQuote & Text3.Text & cQuote & ")" & _
                                   " and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
    
                        cSqlStmt = "update pa13667 set cash_adv = cash_adv + " & nDedAmt & _
                                   " where (`year`=" & cQuote & cmbFlex.Text & cQuote & ")" & _
                                   " and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                    End If
                End If
            End If
            
            oRecordSet.MoveNext
            
        Wend
        
        cField = ""
        cParam = ""
        cParam2 = ""
        For nCtr = 1 To 24
'            If nCtr = 23 Then Stop
            cField = cField & "bas" & Format(nCtr, "00") & IIf(nCtr = 24, "", "+")
            cParam = cParam & "gr" & Format(nCtr, "00") & IIf(nCtr = 24, "", "+")
            cParam2 = cParam2 & "sa" & Format(nCtr, "00") & IIf(nCtr = 24, "", "+")
        
        Next nCtr
        
        cSqlStmt = "update pa13667 set " & _
                   "totbasic=" & cField & "," & _
                   "totgross=" & cParam & "," & _
                   "totsa=" & cParam2 & "," & _
                   "13mopay=round((if(emp_stat=1," & cField & "," & cParam & "+" & cParam2 & ")/12),2)"
                   
'        MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, objdbRs, True
        Script2File cSqlStmt
        
    End If
    
    ShowProgress 4
    
'    MsgBox "Done"
    
    Set oRecordSet = Nothing
    
    GetUserRights PadStr(frmMain.mnuTransaction.Name, " ", 100, PadRight), gUserID
    frmTransaction.lblPeriod.Caption = Text3.Text
    frmTransaction.lblDuration.Caption = "13th Month Pay " & cmbFlex.Text
    frmTransaction.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    aDateInfo = Array("", "")
    
    OpenQueryDNS "select * from pa7730 where 13month=1 and ((year(date_start) = year(curdate())) or (year(date_end) = year(curdate())))", objdbRs, False
    If objdbRs.RecordCount = 0 Then
        MsgBox "Please define a 13th Month Pay Period first!!!", vbExclamation, "System Advisory!!!"
        Command7.Enabled = False
        Command1.Enabled = False
    Else
        aDateInfo = Array(objdbRs("date_start"), objdbRs("date_end"))
        Text3.Text = objdbRs("periodid")
        Text10.Text = objdbRs("workindays")
        Text1.Text = 2
    End If
    
    With cmbFlex
        .Clear
        .AddItem Year(Now) - 2
        .AddItem Year(Now) - 1
        .AddItem Year(Now)
        .AddItem Year(Now) + 1
        .AddItem Year(Now) + 2
        .AddItem Year(Now) + 3
        .ListIndex = 2
        .Visible = True
    End With
    
    MatchCombo Format(Now, "yyyy"), cmbFlex

    OpenQueryDNS "select * from pa13667 where year=" & cQuote & cmbFlex.Text & cQuote, objdbRs, False
    Command1.Enabled = objdbRs.RecordCount > 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log2Audit Name, "CLOSE"
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

