VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalIncrease 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Increase "
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "TXT:RATE_AMT"
      Top             =   1350
      Width           =   1380
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
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "TXT:RATE_ADJ"
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   2010
      Width           =   9825
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8850
         Picture         =   "frmlSalIncrease.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   7125
         Picture         =   "frmlSalIncrease.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   8010
         Picture         =   "frmlSalIncrease.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6225
         Picture         =   "frmlSalIncrease.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5385
         Picture         =   "frmlSalIncrease.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmlSalIncrease.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmlSalIncrease.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmlSalIncrease.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4545
         Picture         =   "frmlSalIncrease.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmlSalIncrease.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmlSalIncrease.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
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
      ToolTipText     =   "TXT:SALIN"
      Top             =   135
      Width           =   720
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
      ToolTipText     =   "TXT:REMARK"
      Top             =   1650
      Width           =   5655
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   1725
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "DAT:DATECON"
      Top             =   720
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56492032
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1725
      TabIndex        =   17
      Tag             =   "1"
      ToolTipText     =   "DAT:DATEREG"
      Top             =   435
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56492032
      CurrentDate     =   38623
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Basic Rate"
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
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Adjust"
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
      TabIndex        =   15
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Confirm"
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
      TabIndex        =   14
      Top             =   750
      Width           =   1530
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Encoded"
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
      TabIndex        =   13
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No"
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
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Top             =   1710
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   2340
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmSalIncrease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSalIncrease
' description   :   Module for Salary Increase
' programmer    :   _-=[ srm ]=-_
' date          :   28 Nov 2014

Option Explicit
    Dim nAdd As Integer
    Dim cSeries As String
    Dim oTempADO As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then dbNavigator Screen.ActiveControl, Me, oTempADO
    ShowRecords
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrDeptSave
    Dim cString As String
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Salary Increase entry?", vbYesNoCancel, "Salary Increase Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA7250", "SALIN=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Deparment ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA7250"), oTempADO, True
                    Script2File InsertFields(Me, "PA7250")
                    
                    Log2Audit Name, "ADD Salary Increase Ref No -->" & Trim(Text1.Text)
                End If
            Else
                OpenQueryDNS EditField(Me, "PA7250", "PA7250.SALIN=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA7250", "PA7250.SALIN=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT Salary Increase Ref No -->" & Trim(Text1.Text)
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select
            
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "SALIN", cSeries
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "SALIN='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    ShowRecords
   
endsave:
    Exit Sub
ErrDeptSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name

End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            
            If Text1.Text <> cSeries Then ResetSeries "SALIN", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SALIN='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords
            
        End If
        
    End If

End Sub

Private Sub Command4_Click()
    On Error GoTo ErrApply
    
    Dim lProceed As Boolean, _
        nCtr As Integer, _
        nWapRate As Double, _
        cSqlStmt As String, _
        cString As String
        
    Dim oRecordSet As New ADODB.Recordset
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If

    If lProceed Then
        If MsgBox("Apply this Salary Increase entry?", vbYesNo, App.Title) = vbYes Then
        
            cString = Text1.Text
            
            nWapRate = Round(Val(Text4.Text) * 0.75, 2)
            
            ShowProgress 0
            
'            ShowProgress 2, , , , "Please Wait while Employee Master is under Backup Process... " & oTempADO("EMPID") & " - " & oTempADO("fullname")
            ShowProgress 2, , , , "Please Wait while Employee Master is under Backup Process... "
            
            cSqlStmt = " insert into dih3670(SALIN, DATECON, DATEBAK, EMPID, TCID, BCID, DATEREG, LOGDATE, FIRSTNAME, MNAME, LASTNAME, BIRTHDAY, DEPID, SHIFTID, POSID, POSITION, POS_ALLOW, PHEALTHNUM, PAGIBIGNO, SSNUM, TIN, TAXCODE, TAXID, ISUNION, SEX, STATUS, " & _
                       " EMP_STAT, WAP, ACTIVE, PAYSTATUS, RATE_AMT, COLA_AMT, SSER1215, SSPREM1215, PS1215, ES1215, MTD_GROSS, MTD_BASIC, MTD_TAXABLE, YTD_GROSS, YTD_GROSS_SA, YTD_BASIC, YTD_WTAX, YTD_COLA, SL_AVAIL, " & _
                       " VL_AVAIL, UL_AVAIL, SL_USE, VL_USE, UL_USE, DATE_HIRE, DATE_FIN, DATE_RES, REF_EMPID, CMPID, COLA1215, OLD_RATE, OLD_COLA, ADD_NO, ADD_BRGY, ADD_CITY, TEL_NUM, BACCNTNO, LVLCode, LABORTYPE, " & _
                       " COSTCENTERID, WORKCENTERID, ERP_ACTIVE, BEPWORKCENTERID, REMARK, S_REMARK ) " & _
                       " select " & _
                        cQuote & Text1.Text & cQuote & "," & _
                        cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                        cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                       " EMPID, TCID, BCID, DATEREG, LOGDATE, FIRSTNAME, MNAME, LASTNAME, BIRTHDAY, DEPID, SHIFTID, POSID, POSITION, POS_ALLOW, PHEALTHNUM, PAGIBIGNO, SSNUM, TIN, TAXCODE, TAXID, ISUNION, SEX, STATUS, " & _
                       " EMP_STAT, WAP, ACTIVE, PAYSTATUS, RATE_AMT, COLA_AMT, SSER1215, SSPREM1215, PS1215, ES1215, MTD_GROSS, MTD_BASIC, MTD_TAXABLE, YTD_GROSS, YTD_GROSS_SA, YTD_BASIC, YTD_WTAX, YTD_COLA, SL_AVAIL, " & _
                       " VL_AVAIL, UL_AVAIL, SL_USE, VL_USE, UL_USE, DATE_HIRE, DATE_FIN, DATE_RES, REF_EMPID, CMPID, COLA1215, OLD_RATE, OLD_COLA, ADD_NO, ADD_BRGY, ADD_CITY, TEL_NUM, BACCNTNO, LVLCode, LABORTYPE, " & _
                       " COSTCENTERID, WORKCENTERID, ERP_ACTIVE, BEPWORKCENTERID, REMARK, S_REMARK from di3670 "
                       
            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, True
            

            
'            cSqlStmt = "Select empid di36770 where date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote
            
            cSqlStmt = " SELECT a.empid, b.paystatus, b.emp_stat,b.wap FROM di36770 a " & _
                       " left join di3670 b on a.empid=b.empid " & _
                       " where a.date >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " group by empid"
            
            OpenQueryDNS cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
                While Not oRecordSet.EOF
                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Please Wait while Employee Master is under Backup Process... " & oRecordSet("EMPID")
                    
'                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
                    
'                    If oRecordSet("empid") = "005393" Then MsgBox "Stop!!"
'                    If oRecordSet("empid") = "002773" Then MsgBox "Stop!!"
'                    If oRecordSet("empid") = "005667" Then MsgBox "Stop!!"
                    
                    Select Case oRecordSet("emp_stat")
                        Case 2 ' regular
                            If oRecordSet("paystatus") = 0 Then
                                cSqlStmt = " Update di3670 set RATE_AMT = " & Val(Text4.Text) & _
                                           " where empid = " & cQuote & oRecordSet("empid") & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, True
                            End If
                        Case 1 ' contractual
                            If oRecordSet("paystatus") = 0 Then
                                If oRecordSet("wap") = 0 Then 'Daily
                                    
                                    cSqlStmt = " Update di3670 set RATE_AMT = " & Val(Text4.Text) & _
                                               " where empid = " & cQuote & oRecordSet("empid") & cQuote
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                Else 'Wap-C
                                    cSqlStmt = " Update di3670 set RATE_AMT = " & nWapRate & _
                                               " where empid = " & cQuote & oRecordSet("empid") & cQuote
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                End If
                            Else 'Emergency
                                    cSqlStmt = " Update di3670 set RATE_AMT = " & Val(Text4.Text) & _
                                               " where empid = " & cQuote & oRecordSet("empid") & cQuote
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                            End If
                        
                        Case 0 ' wap
                            cSqlStmt = " Update di3670 set RATE_AMT = " & nWapRate & _
                                       " where empid = " & cQuote & oRecordSet("empid") & cQuote
                            OpenQueryDNS cSqlStmt, objdbRs, True
                    End Select
                    oRecordSet.MoveNext
                Wend
            End If
            
            cSqlStmt = "update pa73887 set RATE_AMT = " & Val(Text4.Text)
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            cSqlStmt = "update PA7250 set status=1, " & _
                       " date_post=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                       " where SALIN=" & cQuote & Text1.Text & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            ShowProgress 4
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SALIN='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords

        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this Salary Increase!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Exit Sub
    
ErrApply:
    ErrorMsg Err.Number, Err.Description, "Apply Salary Increase #" & Text1.Text, Name

End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
        frmLookup.showPopup 27
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "SALIN='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
            If Not oTempADO.EOF Then
                GetFields Me, oTempADO
                ShowRecords
            End If
        End If
        
    Frame2.Enabled = True

End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    cSeries = GenerateSeries("SALIN")
    While IfExists("PA7250", "PA7250.SALIN=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("SALIN")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    ' --> modified 20050321
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Text1.Enabled = False
        DTPicker1.SetFocus
    End If

End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDeptDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA7250 WHERE SALIN=" & cQuote & Text1.Text & cQuote, oTempADO, True
        
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        Script2File "DELETE FROM PA7250 WHERE SALIN=" & cQuote & Text1.Text & cQuote
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
Exit Sub
ErrDeptDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT * FROM PA7250 ORDER BY SALIN", oTempADO, False
    
    GetFields Me, oTempADO
    ShowRecords
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (nAdd > 0) Then
        MsgBox "Please click CANCEL to abort this entry...", vbOKOnly, App.Title
        Cancel = 1
    Else
        Log2Audit Name, "CLOSE"
    End If
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1

End Sub


