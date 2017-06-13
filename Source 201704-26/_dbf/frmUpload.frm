VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Old Payroll Data"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13560
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
      Left            =   1050
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:PERIODID"
      Top             =   75
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   60
      Width           =   450
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6630
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   11695
      _Version        =   393216
      RowHeightMin    =   285
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      GridColor       =   -2147483632
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
   Begin VB.Frame Frame2 
      Height          =   885
      Left            =   8595
      TabIndex        =   5
      Top             =   6975
      Width           =   4890
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   3930
         Picture         =   "frmUpload.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Refresh"
         Height          =   660
         Left            =   2970
         Picture         =   "frmUpload.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   2010
         Picture         =   "frmUpload.frx":224C
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Download"
         Height          =   660
         Left            =   1050
         Picture         =   "frmUpload.frx":3BCE
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Upload"
         Height          =   660
         Left            =   90
         Picture         =   "frmUpload.frx":5550
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5700
      Top             =   7215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2190
      TabIndex        =   3
      Top             =   135
      Width           =   4005
   End
End
Attribute VB_Name = "frmUpload"
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
' module        :   frmUpload
' programmer    :   _-=[ srm ]=-_
' date          :   23 june 2006

Option Explicit

Dim oTempADO As New ADODB.Recordset, _
    myArray As Variant, _
    nAdd As Integer, _
    oDBFConn As New ADODB.Connection        ' --> for DBF connection

Sub BtnEnable(ByVal nMode As Integer)
    Command1.Enabled = nMode = 0
    Command2.Enabled = nMode = 1
    Command5.Enabled = nMode = 1
    Command3.Enabled = nMode = 1
    Command11.Caption = IIf(nMode = 1, "Cancel", "Close")
End Sub

Sub ClearGrid(ByVal oFlexGrid As MSHFlexGrid, ByVal nColPos As Integer)
    Dim nCtr As Integer
    
    With oFlexGrid
        ShowProgress 0
        
'        .Visible = False
        .Redraw = False
        
        DoEvents
        
        For nCtr = 1 To (.Rows - 1)
            If nCtr > (.Rows - 1) Then Exit For
            
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            .Row = nCtr
        
ikot:
            If nCtr > (.Rows - 1) Then Exit For
            If Val(.TextMatrix(nCtr, nColPos)) = 1 Then
                If .Rows - 1 = 1 Then
                    .AddItem "", .Rows
                    .RowHeight(.RowSel + 1) = 285
                End If
                .RemoveItem nCtr
                GoTo ikot
            End If
            
        Next nCtr
        
        .Redraw = True
'        .Visible = True
        
        ShowProgress 4
        
    End With
End Sub


Sub CheckGrid()
    Dim nCtr As Integer, _
        cDateRes As String
    
    OpenQueryDNS "select date_end from pa7730 where periodid=" & cQuote & Text5.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cDateRes = Format(objdbRs("date_end"), "yyyy-mm-dd")
    Else
        cDateRes = Format(Now, "yyyy-mm-dd")
    End If
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        DoEvents
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            .TextMatrix(nCtr, 6) = PadStr(Trim(.TextMatrix(nCtr, 6)), "0", 3, PadLeft)
            
            .TextMatrix(nCtr, 8) = IIf(.TextMatrix(nCtr, 8) = "W", 0, IIf(.TextMatrix(nCtr, 8) = "C", 1, 2))
            .TextMatrix(nCtr, 9) = IIf(.TextMatrix(nCtr, 9) = "D", 0, 1)
            
            If (Trim(.TextMatrix(nCtr, 11)) <> "") Or (Val(.TextMatrix(nCtr, 6)) = 99) Then
                .TextMatrix(nCtr, 6) = PadStr(left(.TextMatrix(nCtr, 1), 2), "0", 3, PadLeft)
                
                OpenQueryDNS "select lineid, linename from di5463 where lineid=" & cQuote & .TextMatrix(nCtr, 6) & cQuote, objdbRs, False
                .TextMatrix(nCtr, 7) = IIf(objdbRs.RecordCount > 0, objdbRs("linename"), "")
                
                If Trim(.TextMatrix(nCtr, 11)) = "" Then .TextMatrix(nCtr, 11) = cDateRes
                
                .TextMatrix(nCtr, 12) = "Resigned/FC"
'            Else
'                .TextMatrix(nCtr, 12) = IIf(.TextMatrix(nCtr, 8) = 0, "WAP", IIf(.TextMatrix(nCtr, 8) = 1, "Contractual", "Regular"))
            End If
            
            OpenQueryDNS "select empid from di3670 where empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nCtr, 13) = 1
                .TextMatrix(nCtr, 57) = ""
                HiLyt2 nCtr, MSHFlexGrid1, vbBlack
            Else
                .TextMatrix(nCtr, 13) = 0
                .TextMatrix(nCtr, 57) = "Undefined Employee..."
                HiLyt2 nCtr, MSHFlexGrid1, vbRed
            End If
            
            If Trim(.TextMatrix(nCtr, 55)) <> "" Then
                OpenQueryDNS "select taxid from pa8290 where taxcode=" & cQuote & Trim(.TextMatrix(nCtr, 55)) & cQuote, objdbRs, False
                .TextMatrix(nCtr, 56) = IIf(objdbRs.RecordCount > 0, objdbRs("taxid"), "")
            End If
            
            .TextMatrix(nCtr, 60) = IIf(.TextMatrix(nCtr, 60) = "M", 0, 1)

            If Trim(.TextMatrix(nCtr, 2)) = "" Then .TextMatrix(nCtr, 2) = .TextMatrix(nCtr, 61)
            If Trim(.TextMatrix(nCtr, 3)) = "" Then .TextMatrix(nCtr, 3) = .TextMatrix(nCtr, 62)
            If Trim(.TextMatrix(nCtr, 4)) = "" Then .TextMatrix(nCtr, 4) = .TextMatrix(nCtr, 63)

'            .TextMatrix(nCtr, 61) = IIf(.TextMatrix(nCtr, 61) = "Y", 1, 0)
'            .TextMatrix(nCtr, 62) = PadStr(Trim(.TextMatrix(nCtr, 62)), "0", 3, PadLeft)
            
'                     "TXT:[9Birthday]:50:False", _
'                    6 "TXT:[0Sex]:1:False", _
'                     "TXT:[1Union]:1:False", _
'                     "TXT:[2Position]:3:False")
            
            .TopRow = nCtr
            
        Next nCtr
        
        .TopRow = 1
        
        ShowProgress 4
        
    End With
    
End Sub

Function DetectDBF(cDBFPath As String) As Boolean
    On Error GoTo ErrDetect
    Dim cString As String
    
    DoEvents

    If oDBFConn.State = adStateOpen Then oDBFConn.Close
    With oDBFConn
        .CursorLocation = adUseClient
        cString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & cDBFPath & ";" & _
                   "Extended Properties=" & cQuote & "DBASE IV;" & cQuote & ";"
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

Private Sub Command1_Click()
    On Error GoTo ErrLoad
    
    Dim cSqlStmt As String, oFile As File, oTextFile As New FileSystemObject
    
    CommonDialog1.CancelError = False

    CommonDialog1.InitDir = CheckPath(cUploadPath) & "payroll\"
    CommonDialog1.Filter = "Payroll Path File |*.dbf"
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    
    Set oFile = oTextFile.GetFile(CommonDialog1.FileName)
    If DetectDBF(left(oFile.Path, Len(oFile.Path) - Len(oFile.Name) - 1)) Then
        cSqlStmt = " select a.empno, b.lastname, b.firstname, b.minitial, a.employee, " & _
                   " a.deptcode, a.department, a.empstatus, " & _
                   " a.paystatus, b.datehired, b.resigned, '' as status, 0 as tag, " & _
                   " a.rate, a.cola, a.posallow, " & _
                   " a.regdays, a.regpay, a.regot, a.regotpay, a.saot, a.saotpay, " & _
                   " a.ndfdays, a.ndfpay, a.regnd, a.regndpay, a.sand, a.sandpay, " & _
                   " a.sunhrs, a.sunpay, a.sunot, a.sunotpay, " & _
                   " a.holdays, a.holpay, " & _
                   " a.adjustment, a.adjust2, " & _
                   " a.incentive, a.m13pay, a.grosspay, a.taxable, a.gross16231, " & _
                   " a.ssnum, a.sssloan, a.ssspremium, a.ssser, a.sss01, a.ec001, " & _
                   " a.medicare, a.med01, " & _
                   " a.pagibigno, a.pagibig, a.pagloan, " & _
                   " a.tin, taxwheld, a.taxcode, '' as taxid, '' as remark, 0 as isup, " & _
                   " b.birthday, b.sex, a.esurn, a.ename, a.minitial as minitial2 " & _
                   " from k4pay a " & _
                   " left join masterk4 b on a.empno=b.empno"
        Set oTempADO = oDBFConn.Execute(cSqlStmt)
        If oTempADO.RecordCount > 0 Then
            QueryAttach oTempADO, MSHFlexGrid1, myArray, , , , 1
            CheckGrid
        Else
            SetGridColumn myArray, MSHFlexGrid1
        End If
        nAdd = IIf(oTempADO.RecordCount > 0, 1, 0)
        BtnEnable nAdd
    End If

ErrLoad:
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            SetGridColumn myArray, MSHFlexGrid1
            
            nAdd = 0
            BtnEnable nAdd
        End If
    End If
End Sub

Function ChkDate(ByVal cDateStr As String, Optional ByVal dDate As Date) As String
    On Error GoTo ErrDate
    If Trim(cDateStr) = "" Then cDateStr = Format(Now, "yyyy-mm-dd")
    ChkDate = Format(cDateStr, "yyyy-mm-dd")
    Exit Function
ErrDate:
    ChkDate = Format(Now, "yyyy-mm-dd")
End Function

Private Sub Command2_Click()
    Dim cParam, _
        cSqlStmt As String, _
        nPeriod, _
        nActive, _
        nCtr As Integer, _
        nColaAmt As Double
        
    cSqlStmt = "Warning!!!" & vbCrLf & vbCrLf & _
               "By clicking the [YES] button, you are agreeing" & vbCrLf & _
               "to the condition that all the transactions for the" & vbCrLf & _
               "selected period will be replaced by this utility." & vbCrLf & _
               "Click the [YES] button to proceed now."
    If MsgBox(cSqlStmt, vbYesNo, "System Advisory!!!") = vbYes Then
    
        With MSHFlexGrid1
        
            DoEvents
            cParam = ""
            For nCtr = 1 To (.Rows - 1)
                If Val(.TextMatrix(nCtr, 13)) = 1 Then cParam = cParam & cQuote & Trim(.TextMatrix(nCtr, 1)) & cQuote & ","
            Next nCtr
            
            If Trim(cParam) <> "" Then cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
            
'    myArray = Array("TXT:[1Emp ID]:10:True", _
'                    "TXT:[2Last Name]:25:True", _
                     "TXT:[3First Name]:25:True", _
                     "TXT:[4Middle Name]:25:True", _
'                    "TXT:[5Full Name]:60:False", _
'                    "TXT:[6DEPID]:3:False", _
                     "TXT:[7Department]:20:True", _
                     "TXT:[8EMP_STAT]:1:False", _
                     "TXT:[9PAY_STAT]:1:False", _
'                   1 "TXT:[0DATEHIRE]:10:False", _
                     "TXT:[1DATERES]:10:False", _
                     "TXT:[2Status]:15:True", _
                     "NUM:[3Tag]:1:False", _
'                    "NUM:[4Rate Amt]:10.4:False", _
                     "NUM:[5COLA Amt]:10.4:False", _
                     "NUM:[6Pos Allow]:10.4:False",
'                    "NUM:[7Reg Day]:10.4:False", _
                     "NUM:[8Reg Pay]:10.4:False", _
                     "NUM:[9Reg OT Hr]:10.4:False",
'                   2 "NUM:[0Reg OT Pay]:10.4:False", _
                     "NUM:[1SA Reg OT Hr]:10.4:False", _
                     "NUM:[2SA Reg Pay]:10.4:False",
'                    "NUM:[3NDiff Day]:10.4:False", _
                     "NUM:[4NDiff Pay]:10.4:False", _
                     "NUM:[5NDiff OT Hr]:10.4:False", _
                     "NUM:[6NDiff OT Pay]:10.4:False", _
                     "NUM:[7SA NDiff OT Hr]:10.4:False", _
                     "NUM:[8SA NDiff Pay]:10.4:False",
'                    "NUM:[9Sun Day]:10.4:False", _
                    3 "NUM:[0Sun Pay]:10.4:False", _
                     "NUM:[1Sun OT Hr]:10.4:False", _
                     "NUM:[2Sun OT Pay]:10.4:False",
'                    "NUM:[3Hol Day]:10.4:False", _
                     "NUM:[4Hol Pay]:10.4:False",
'                    "NUM:[5Adj Pay]:10.4:False", _
                     "NUM:[6SA Adj Pay]:10.4:False",
'                    "NUM:[7Incentive]:10.4:False", _
                     "NUM:[8M13Pay]:10.4:False", _
                     "NUM:[9Gross Pay]:10.4:False", _
                    4 "NUM:[0Tax Amt]:10.4:False", _
                    "NUM:[1Gross1631]:10.4:False",
'                    "TXT:[2SSSNUM]:15:False", _
                     "NUM:[3SSSLoan]:10.4:False", _
                     "NUM:[4SSSPrem]:10.4:False", _
                     "NUM:[5SSSER]:10.4:False", _
                     "NUM:[6SSS01]:10.4:False", _
                     "NUM:[7EC01]:10.4:False", _
                     "NUM:[8MEDICARE]:10.4:False", _
                     "NUM:[9MED01]:10.4:False",
'                   5 "TXT:[0PAGIBIGNO]:15:False", _
                     "NUM:[1PAGIBIG]:10.4:False", _
                     "NUM:[2PAGIBIG Loan]:10.4:False", _
                     "TXT:[3TIN]:15:False", _
                     "NUM:[4WTax]:10.4:False", _
                     "TXT:[5TaxCode]:5:False", _
                     "TXT:[6TaxID]:5:False", _
                     "TXT:[7Remark]:50:True", _
                     "NUM:[8Uploaded]:1:False", _
                     "TXT:[9Birthday]:50:False", _
                    6 "TXT:[0Sex]:1:False", _
'                    "TXT:[1Last Name]:25:True", _
                     "TXT:[2First Name]:25:True", _
                     "TXT:[3Middle Name]:25:True", _

            cSqlStmt = "DELETE FROM PA87260 WHERE (PERIODID=" & cQuote & Text5.Text & cQuote & ")" & _
                       " AND (EMPID IN " & cParam & ")"
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            cSqlStmt = "DELETE FROM PA87263 WHERE (PERIODID=" & cQuote & Text5.Text & cQuote & ")" & _
                       " AND (EMPID IN " & cParam & ")"
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            OpenQueryDNS "SELECT * FROM PA7730 WHERE PERIODID=" & cQuote & Text5.Text & cQuote, objdbRs, False
            nPeriod = IIf(objdbRs.RecordCount > 0, objdbRs("status"), 0)
            
            ShowProgress 0, (.Rows - 1)
            
            For nCtr = 1 To (.Rows - 1)
            
                ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                
                If Val(.TextMatrix(nCtr, 13)) = 1 Then
                    If Val(.TextMatrix(nCtr, 17)) > 0 Then
                        nColaAmt = Val(.TextMatrix(nCtr, 15)) / Val(.TextMatrix(nCtr, 17))
                    Else
                        nColaAmt = 0
                    End If
                    
                    cSqlStmt = "INSERT INTO PA87260(PERIODID,PERIOD_STAT,EMPID,LASTNAME,FIRSTNAME,MNAME,FULLNAME," & _
                               "DEPID,EMP_STAT,PAYSTATUS,`ACTIVE`,RATE_AMT,COLA_AMT,POS_ALLOW," & _
                               "REG_DAY,REG_PAY,REG_OT_HR,REG_OT_PAY,SA_REG_OT,SA_REG_PAY," & _
                               "NDIFF_DAY,NDIFF_PAY,NDIFF_OT_HR,NDIFF_OT_PAY,SA_NDIFF_OT,SA_NDIFF_PAY," & _
                               "SUN_HR,SUN_PAY,SUN_OT,SUN_OT_PAY,`HOLIDAY`,HOL_PAY," & _
                               "ADJ_PAY,SA_ADJ_PAY,LEAVE_PAY,M13PAY,GROSS_PAY,TAXABLE,GROSS16231," & _
                               "SSSNUM,SSPREM,SSER,SSS01,EC001,MEDICARE,MED01," & _
                               "PAGIBIGNO,TINNUM,WTAX,TAXID)VALUES(" & _
                               cQuote & Text5.Text & cQuote & "," & nPeriod & "," & _
                               cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & cQuote & DecodeStr(.TextMatrix(nCtr, 2)) & cQuote & "," & cQuote & DecodeStr(.TextMatrix(nCtr, 3)) & cQuote & "," & cQuote & DecodeStr(.TextMatrix(nCtr, 4)) & cQuote & "," & cQuote & DecodeStr(.TextMatrix(nCtr, 5)) & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & _
                               Val(.TextMatrix(nCtr, 8)) & "," & Val(.TextMatrix(nCtr, 9)) & "," & IIf(Trim(.TextMatrix(nCtr, 11)) = "", 0, 1) & "," & Val(.TextMatrix(nCtr, 14)) & "," & nColaAmt & "," & Val(.TextMatrix(nCtr, 16)) & "," & _
                               Val(.TextMatrix(nCtr, 17)) & "," & Val(.TextMatrix(nCtr, 18)) & "," & Val(.TextMatrix(nCtr, 19)) & "," & Val(.TextMatrix(nCtr, 20)) & "," & Val(.TextMatrix(nCtr, 21)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
                               Val(.TextMatrix(nCtr, 23)) & "," & Val(.TextMatrix(nCtr, 24)) & "," & Val(.TextMatrix(nCtr, 25)) & "," & Val(.TextMatrix(nCtr, 26)) & "," & Val(.TextMatrix(nCtr, 27)) & "," & Val(.TextMatrix(nCtr, 28)) & "," & _
                               Val(.TextMatrix(nCtr, 29)) & "," & Val(.TextMatrix(nCtr, 30)) & "," & Val(.TextMatrix(nCtr, 31)) & "," & Val(.TextMatrix(nCtr, 32)) & "," & _
                               Val(.TextMatrix(nCtr, 33)) & "," & Val(.TextMatrix(nCtr, 34)) & "," & _
                               Val(.TextMatrix(nCtr, 35)) & "," & Val(.TextMatrix(nCtr, 36)) & "," & _
                               Val(.TextMatrix(nCtr, 37)) & "," & Val(.TextMatrix(nCtr, 38)) & "," & Val(.TextMatrix(nCtr, 39)) & "," & Val(.TextMatrix(nCtr, 40)) & "," & Val(.TextMatrix(nCtr, 41)) & "," & _
                               cQuote & .TextMatrix(nCtr, 42) & cQuote & "," & Val(.TextMatrix(nCtr, 44)) & "," & Val(.TextMatrix(nCtr, 45)) & "," & Val(.TextMatrix(nCtr, 46)) & "," & Val(.TextMatrix(nCtr, 47)) & "," & _
                               Val(.TextMatrix(nCtr, 48)) & "," & Val(.TextMatrix(nCtr, 49)) & "," & _
                               cQuote & .TextMatrix(nCtr, 50) & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 53) & cQuote & "," & Val(.TextMatrix(nCtr, 54)) & "," & cQuote & .TextMatrix(nCtr, 56) & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                    
                    
                    ' --> deduction entry...
                    cSqlStmt = "select DEDID,DEF_AMT,FIX_DED,DEDTAG, DEDTYPE from pa3330 where " & IIf(nPeriod = 1, "PERIOD1=1", "PERIOD2=1")
                    OpenQueryDNS cSqlStmt, oTempADO, False
                    If oTempADO.RecordCount > 0 Then
                        While Not oTempADO.EOF
                            cParam = "0,0"
                            Select Case oTempADO("dedid")
                                Case "001"  ' --> SSS Premium
                                    cParam = Val(.TextMatrix(nCtr, 44)) & "," & Val(.TextMatrix(nCtr, 45))
                                Case "002"  ' --> SSS Loan
                                    cParam = Val(.TextMatrix(nCtr, 43)) & ",0"
                                Case "003"  ' --> Pag-Ibig Premium
                                    cParam = Val(.TextMatrix(nCtr, 51)) & ",0"
                                Case "004"  ' --> Pag-Ibig Loan
                                    cParam = Val(.TextMatrix(nCtr, 52)) & ",0"
                                Case "005"  ' --> Medicare
                                    cParam = Val(.TextMatrix(nCtr, 48)) & ",0"
                                Case "006"  ' --> Withholding Tax
                                    cParam = Val(.TextMatrix(nCtr, 54)) & ",0"
                            End Select
                            
                            cSqlStmt = "INSERT INTO PA87263(PERIODID, PERIOD_STAT, " & _
                                       " EMPID, DEDID, DED_AMT, DED_AMT2)VALUES(" & _
                                       cQuote & Text5.Text & cQuote & "," & nPeriod & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       cQuote & oTempADO("dedid") & cQuote & "," & _
                                       cParam & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt

                            oTempADO.MoveNext
                        Wend
                    End If
                    
                    If Trim(.TextMatrix(nCtr, 11)) <> "" Then
                        cSqlStmt = "update di3670 set active=1, date_res=" & cQuote & ChkDate(.TextMatrix(nCtr, 11)) & cQuote & _
                                   " where empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
'                    Else
'                        cSqlStmt = "update di3670 set active=0" & _
'                                   " where empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote
'                        OpenQueryDNS cSqlStmt, objdbRs, True
'                        Script2File cSqlStmt
                    End If

                    .TextMatrix(nCtr, 58) = 1
                    
                End If
                
            Next nCtr
            
            ShowProgress 4
            
            ClearGrid MSHFlexGrid1, 58
            
            If (.Rows - 1) > 0 Then
                cSqlStmt = "System detected existence of undefined employee(s)..." & vbCrLf & _
                           "Would you like to upload this to master file automatically?"
                If MsgBox(cSqlStmt, vbYesNo, "Upload to Master...") = vbYes Then
                    
                    ShowProgress 0
                    
                    For nCtr = 1 To (.Rows - 1)
                        
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If .TextMatrix(nCtr, 1) <> "" Then
                            nActive = IIf(Trim(.TextMatrix(nCtr, 11)) <> "", 0, 1)
                            
'                            If Trim(.TextMatrix(nCtr, 59)) <> "" Then
'                                If Val(left(.TextMatrix(nCtr, 59), 2)) > 12 Then
'                                    cParam = Format(Now, "yyyy-mm-dd")
'                                Else
'                                    cParam = Format(.TextMatrix(nCtr, 59), "yyyy-mm-dd")
'                                End If
'                            Else
'                                cParam = Format(Now, "yyyy-mm-dd")
'                            End If
                            
                            If Val(.TextMatrix(nCtr, 17)) > 0 Then
                                nColaAmt = Val(.TextMatrix(nCtr, 15)) / Val(.TextMatrix(nCtr, 17))
                            Else
                                nColaAmt = 0
                            End If
                            
                            cSqlStmt = " insert into di3670(datereg,empid,lastname,firstname,mname,depid,emp_stat,paystatus,date_hire,date_res,rate_amt,cola_amt,pos_allow, " & _
                                       " ssnum,pagibigno,tin,taxcode,taxid,birthday,sex,active,cmpid)values(" & _
                                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & cQuote & .TextMatrix(nCtr, 4) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 9) & cQuote & "," & cQuote & ChkDate(.TextMatrix(nCtr, 10)) & cQuote & "," & _
                                       cQuote & ChkDate(.TextMatrix(nCtr, 11)) & cQuote & "," & _
                                       .TextMatrix(nCtr, 14) & "," & nColaAmt & "," & .TextMatrix(nCtr, 16) & "," & _
                                       cQuote & .TextMatrix(nCtr, 42) & cQuote & "," & cQuote & .TextMatrix(nCtr, 50) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 53) & cQuote & "," & cQuote & .TextMatrix(nCtr, 55) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 56) & cQuote & "," & cQuote & ChkDate(.TextMatrix(nCtr, 59)) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 60) & cQuote & "," & _
                                       nActive & "," & gCompanyID & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                    
                    Next nCtr
                    
                    ShowProgress 4
                
                End If
            End If
            
        End With
        
        CheckGrid
        MsgBox "Done!"
    End If
End Sub

Private Sub Command3_Click()
    CheckGrid
End Sub

Private Sub Command4_Click()
    frmLookup.showPopup 5
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        If cResult <> "" Then
            OpenQueryDNS "select * from pa7730 where periodid = " & cQuote & cResult & cQuote, objdbRs, False
            Label14.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("duration"), "")
            Text5.Text = cResult
        Else
            Text5.Text = ""
            Label14.Caption = ""
        End If
    End If
    Text5.SetFocus
End Sub

Private Sub Command5_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer
        
    With MSHFlexGrid1
        For nCtr = 1 To (.Rows - 1)
            
        Next nCtr
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"

    nAdd = 0
    myArray = Array("TXT:[Emp ID]:10:True", _
                    "TXT:[Last Name]:25:True", "TXT:[First Name]:25:True", "TXT:[Middle Name]:25:True", "TXT:[Full Name]:60:False", _
                    "TXT:[DEPID]:3:False", "TXT:[Department]:20:True", _
                    "TXT:[EMP_STAT]:1:False", "TXT:[PAY_STAT]:1:False", _
                    "TXT:[DATEHIRE]:10:False", "TXT:[DATERES]:10:False", "TXT:[Status]:15:True", "NUM:[Tag]:1:False", _
                    "NUM:[Rate Amt]:10.4:False", "NUM:[COLA Amt]:10.4:False", "NUM:[Pos Allow]:10.4:False", _
                    "NUM:[Reg Day]:10.4:False", "NUM:[Reg Pay]:10.4:False", "NUM:[Reg OT Hr]:10.4:False", "NUM:[Reg OT Pay]:10.4:False", "NUM:[SA Reg OT Hr]:10.4:False", "NUM:[SA Reg Pay]:10.4:False", _
                    "NUM:[NDiff Day]:10.4:False", "NUM:[NDiff Pay]:10.4:False", "NUM:[NDiff OT Hr]:10.4:False", "NUM:[NDiff OT Pay]:10.4:False", "NUM:[SA NDiff OT Hr]:10.4:False", "NUM:[SA NDiff Pay]:10.4:False", _
                    "NUM:[Sun Day]:10.4:False", "NUM:[Sun Pay]:10.4:False", "NUM:[Sun OT Hr]:10.4:False", "NUM:[Sun OT Pay]:10.4:False", _
                    "NUM:[Hol Day]:10.4:False", "NUM:[Hol Pay]:10.4:False", _
                    "NUM:[Adj Pay]:10.4:False", "NUM:[SA Adj Pay]:10.4:False", _
                    "NUM:[Incentive]:10.4:False", "NUM:[M13Pay]:10.4:False", "NUM:[Gross Pay]:10.4:False", "NUM:[Tax Amt]:10.4:False", "NUM:[Gross1631]:10.4:False", _
                    "TXT:[SSSNUM]:15:False", "NUM:[SSSLoan]:10.4:False", "NUM:[SSSPrem]:10.4:False", "NUM:[SSSER]:10.4:False", "NUM:[SSS01]:10.4:False", "NUM:[EC01]:10.4:False", _
                    "NUM:[MEDICARE]:10.4:False", "NUM:[MED01]:10.4:False", _
                    "TXT:[PAGIBIGNO]:15:False", "NUM:[PAGIBIG]:10.4:False", "NUM:[PAGIBIG Loan]:10.4:False", _
                    "TXT:[TIN]:15:False", "NUM:[WTax]:10.4:False", "TXT:[TaxCode]:5:False", "TXT:[TAXID]:3:False", _
                    "TXT:[Remark]:50:True", "NUM:[Uploaded]:1:False", _
                    "TXT:[Birthday]:50:False", "TXT:[Sex]:1:False", _
                    "TXT:[Last Name 2]:25:False", "TXT:[First Name 2]:25:False", "TXT:[Middle Name 2]:25:False")

    BtnEnable nAdd
    SetGridColumn myArray, MSHFlexGrid1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log2Audit Name, "CLOSE"
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
    Set oDBFConn = Nothing
End Sub


Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Text5.Text = "" Then
            Command4_Click
        Else
            OpenQueryDNS "select * from pa7730 where periodid = " & cQuote & Text5.Text & cQuote, objdbRs, False
            Label14.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("duration"), "")
        End If
    End If
End Sub
