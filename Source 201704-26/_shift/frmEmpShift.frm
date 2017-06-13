VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Begin VB.Form frmEmpShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Employee Shifting Schedule"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
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
      Index           =   2
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   405
      Width           =   4335
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5535
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   2370
      Visible         =   0   'False
      Width           =   1215
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
      Index           =   0
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   75
      Width           =   4335
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
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5250
      Left            =   3090
      TabIndex        =   2
      Top             =   750
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9260
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
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
      Height          =   870
      Left            =   7290
      TabIndex        =   12
      Top             =   5985
      Width           =   5760
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   165
         Picture         =   "frmEmpShift.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   2820
         Picture         =   "frmEmpShift.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1980
         Picture         =   "frmEmpShift.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cre&ate"
         Height          =   660
         Left            =   1140
         Picture         =   "frmEmpShift.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   4740
         Picture         =   "frmEmpShift.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   3780
         Picture         =   "frmEmpShift.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexPeriod 
      Height          =   5235
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   9234
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
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
   Begin ciaXPPanel.XPPanel XPPanel3 
      Height          =   750
      Left            =   90
      TabIndex        =   13
      Top             =   6090
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1323
      LicValid        =   -1  'True
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   390
         Width           =   465
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   75
         Width           =   990
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Period"
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
         Height          =   255
         Left            =   2385
         TabIndex        =   20
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays"
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
         Height          =   255
         Left            =   1905
         TabIndex        =   19
         Top             =   450
         Width           =   1515
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Period"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Regular Days"
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   1515
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   120
      TabIndex        =   11
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   120
      TabIndex        =   10
      Top             =   135
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   6990
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmEmpShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmEmpShift
' description   :   Custom Shifting Schedule by Employee module
' programmer    :   _-=[ srm ]=-_
' date          :   3 mar 2006

Option Explicit
    Dim nAdd As Integer, _
        cParam As String, cParam2 As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant, myArray2 As Variant
    
Sub SetFilter(ByVal cValue As String, cShiftInfo As String)
    cParam = cValue
    cParam2 = cShiftInfo
    Text1.Text = cValue
    MSHFlexPeriod_EnterCell
End Sub

Function ChkDateinGrid(ByVal cDateStr As String, ByVal oFlexGrid As MSHFlexGrid) As Integer
    Dim nCtr As Integer, _
        aDateInfo As Variant
    
    With oFlexGrid
        For nCtr = 1 To (.Rows - 1)
            aDateInfo = Array(Format(.TextMatrix(nCtr, 3), "mm/dd/yyyy"), Format(.TextMatrix(nCtr, 4), "mm/dd/yyyy"))
            If (DateValue(aDateInfo(0)) <= DateValue(cDateStr)) And (DateValue(aDateInfo(1)) >= DateValue(cDateStr)) Then
                ChkDateinGrid = nCtr
                Exit For
            End If
        Next nCtr
    End With
End Function

'Sub HiLyt2()
'    Dim nCtr As Integer
'    With MSHFlexGrid1
'        DoEvents
'        .Redraw = False
'        For nCtr = 1 To (.Rows - 1)
'            .Row = nCtr
'            .FillStyle = flexFillRepeat
'            .Col = 1
'            .ColSel = .Cols() - 1
'            If Val(.TextMatrix(nCtr, 8)) = 1 Then
'                .CellForeColor = vbBlue
'            Else
'                .CellForeColor = IIf(UCase(Trim(.TextMatrix(nCtr, 2))) = "SUN", vbRed, vbBlack)
'            End If
'            .FillStyle = flexFillSingle
'        Next nCtr
'        .Redraw = True
'    End With
'End Sub

Private Sub Command10_Click()
    On Error GoTo ErrEmpShiftSave
    Dim nCtr As Integer, _
        cSqlStmt As String
    
    Select Case MsgBox("Save/Update custom employee shifting entry?", vbYesNoCancel, "Custom Shift Entry...")
        Case vbYes
            If IfExists("DI36770", "EMPID=" & cQuote & Text1.Text & cQuote & " AND PERIODID=" & cQuote & Text2(4).Text & cQuote) Then
                cSqlStmt = "DELETE FROM DI36770 WHERE EMPID=" & cQuote & Text1.Text & cQuote & " AND PERIODID=" & cQuote & Text2(4).Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            End If
            
            ShowProgress 0
            
            With MSHFlexGrid1
                For nCtr = 1 To (.Rows - 1)
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,`remark`)values(" & _
                               cQuote & Text1.Text & cQuote & "," & _
                               cQuote & Text2(4).Text & cQuote & "," & _
                               cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                               cQuote & EncodeStr(.TextMatrix(nCtr, 7)) & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                    
                    Log2Audit Me.Name, "Assigned shift to Emp ID#" & Text1.Text & " for " & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd")
                Next nCtr
            End With
            
            ShowProgress 4
        
        Case vbCancel
            GoTo endsave
            
    End Select

    nAdd = 0
    CtrlPanel Me, nAdd
    
    MSHFlexPeriod.Enabled = True
    MSHFlexPeriod_EnterCell

endsave:
    Exit Sub
    
ErrEmpShiftSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
    Resume endsave
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            nAdd = 0
            CtrlPanel Me, nAdd
            
            MSHFlexPeriod.Enabled = True
            MSHFlexPeriod_EnterCell
        End If
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmp36770( [EMPID] char(6)," & _
               " [FULLNAME] char(100),  [DEPTNAME] char(100)," & _
               " [DURATION] char(50),   [REMARK] char(100)," & _
               " [DATE1] date,          [DATE2] date," & _
               " [SHIFTDESC] char(100), [TAG] integer," & _
               " [DAY_DATE] date,       [DAY_NAME] char(20)," & _
               " [TIME1] char(15),      [TIME2] char(15)," & _
               " [REGDAY] integer,      [HOLIDAY] integer)"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmp36770", oTempADO, True
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset
    
    CreateTemp
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                cSqlStmt = "insert into tmp36770(empid,fullname,deptname,[duration],[remark],date1,date2," & _
                           "shiftdesc,[tag],day_date,day_name,time1,time2,regday,[holiday])values(" & _
                           cQuote & Text1.Text & cQuote & "," & _
                           cQuote & EncodeStr2(Text2(0).Text) & cQuote & "," & _
                           cQuote & EncodeStr2(Text2(2).Text) & cQuote & "," & _
                           cQuote & EncodeStr2(Label4.Caption) & cQuote & "," & _
                           cQuote & EncodeStr2(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                           cQuote & Format(MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 3), "mm/dd/yyyy") & cQuote & "," & _
                           cQuote & Format(MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 4), "mm/dd/yyyy") & cQuote & "," & _
                           cQuote & EncodeStr2(.TextMatrix(nCtr, 4)) & cQuote & "," & _
                           .TextMatrix(nCtr, 8) & "," & _
                           cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                           cQuote & WeekdayName(Weekday(.TextMatrix(nCtr, 1))) & cQuote & "," & _
                           cQuote & .TextMatrix(nCtr, 5) & cQuote & "," & _
                           cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & _
                           Text2(6).Text & "," & _
                           Text2(1).Text & ")"
                QueryTemp cSqlStmt, objdbRs, True
            End If
            
        Next nCtr
        
        ShowProgress 3
        
        QueryTemp "select * from tmp36770", objdbRs, False
        If objdbRs.RecordCount > 0 Then
            GenerateReport "Employee Shifting Schedule", "PRV36770.RPT"
        Else
            MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
        End If
        
        ShowProgress 4
        
    End With
        
    Set oRecordSet = Nothing
End Sub

Private Sub Command7_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        nRowPos As Integer, _
        aShiftInfo As Variant, _
        aDateInfo As Variant
    
    aDateInfo = Array("", "")
    
    If (MSHFlexGrid1.Rows - 1) > 1 Then
        MsgBox "Please click [Edit] to change schedule for the selected period!", vbInformation, "System Advisory!!!"
        Exit Sub
    End If
    
    aShiftInfo = Array("", "", "")
    
    nAdd = 1
    
    CtrlPanel Me, nAdd
    MSHFlexPeriod.Enabled = False
    
    aDateInfo(0) = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 3)
    aDateInfo(1) = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 4)
    
    OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380 where shiftid=" & cQuote & cParam2 & cQuote, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        aShiftInfo(0) = DecodeStr(oTempADO("description"))
        aShiftInfo(1) = Format(oTempADO("time1"), "h:mm AM/PM")
        aShiftInfo(2) = Format(oTempADO("time2"), "h:mm AM/PM")
    End If

    With MSHFlexGrid1
        .Redraw = False
        
        DoEvents
        
        For nCtr = Day(aDateInfo(0)) To Day(aDateInfo(1))
        
            nRowPos = nRowPos + 1
            
            .Rows = nRowPos + 1
            .RowHeight(nRowPos) = 285
            
            .TextMatrix(nRowPos, 1) = Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "mmm dd")
            .TextMatrix(nRowPos, 2) = WeekdayName(Weekday(DateAdd("d", nRowPos - 1, aDateInfo(0))), True)
            If Weekday(DateAdd("d", nRowPos - 1, aDateInfo(0))) <> vbSunday Then
                .TextMatrix(nRowPos, 3) = cParam2
                .TextMatrix(nRowPos, 4) = aShiftInfo(0)
                .TextMatrix(nRowPos, 5) = aShiftInfo(1)
                .TextMatrix(nRowPos, 6) = aShiftInfo(2)
            Else
                HiLyt2 nRowPos, MSHFlexGrid1, vbRed
            End If
        
            ' --> check if date is holiday
            cSqlStmt = "select a.description from pa4329 a" & _
                       " where (a.date=" & cQuote & Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "yyyy-mm-dd") & cQuote & ") or" & _
                       " (date_format(a.date,'%m %d')=" & cQuote & Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "mmm dd") & cQuote & ")"
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nRowPos, 7) = objdbRs("description")
                .TextMatrix(nRowPos, 8) = 1
                HiLyt2 nRowPos, MSHFlexGrid1, vbBlue
            Else
                .TextMatrix(nRowPos, 8) = 0
            End If
            
            .TextMatrix(nRowPos, 9) = Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "yyyy-mm-dd")
            
        Next nCtr
        
        RefreshGrid MSHFlexGrid1, True
        .Redraw = True
    End With
    
    MSHFlexGrid1.SetFocus
End Sub

Private Sub Command8_Click()
'    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
'        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
'
'        nAdd = 2
'        CtrlPanel Me, nAdd
'
'        MSHFlexPeriod.Enabled = False
'    End If
    nAdd = 2
    CtrlPanel Me, nAdd
    
    MSHFlexPeriod.Enabled = False
    MSHFlexGrid1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim cSqlStmt As String, _
        nPos As Integer
    
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Date]:8:True", _
                    "TXT:[Day]:6:True", _
                    "TXT:[ShiftID]:5:False", _
                    "TXT:[Shift Info]:20:True", _
                    "TXT:[Start Time]:10:True", _
                    "TXT:[End Time]:10:True", _
                    "TXT:[Remark]:50:True", _
                    "NUM:[Holiday Tag]:1:False", _
                    "TXT:[Date]:10:False")
    myArray2 = Array("TXT:[ID]:7:True", _
                     "TXT:[Duration]:20:True", _
                     "TXT:[Start Date]:10:False", _
                     "TXT:[End Date]:10:False", _
                     "NUM:[Reg Day]:2:False", _
                     "NUM:[Holiday]:2:False")
    
    Tag = frmEmployee.Tag
    nAdd = 0
    
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT PERIODID, DURATION, DATE_START, DATE_END, WORKINDAYS, HOLIDAYS FROM PA7730 WHERE (PCLOSE=0) /*and (DATE_START < CURDATE() and DATE_END < CURDATE())*/ ORDER BY PERIODID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexPeriod, myArray2, , , True
        nPos = ChkDateinGrid(Format(Now, "mm/dd/yyyy"), MSHFlexPeriod)
        If nPos <= (MSHFlexPeriod.Rows - 1) Then
            MSHFlexPeriod.Row = nPos
        End If
    Else
        MSHFlexPeriod.Clear
        SetGridColumn myArray, MSHFlexPeriod
    End If
    
    MSHFlexPeriod_EnterCell
'    MSHFlexPeriod.SetFocus
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

Private Sub MSHFlexGrid1_DblClick()
    MSHFlexGrid1_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid1_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cSqlStmt As String
    
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                Select Case .ColSel
                    Case 4
                        frmLookup.showPopup 9
                        frmLookup.Show 1
                        If Trim(cResult) <> "" Then
                            cSqlStmt = "select shiftid, `description`, time1, time2 from pa74380 where shiftid=" & cQuote & cResult & cQuote
                            OpenQueryDNS cSqlStmt, objdbRs, False
                            If objdbRs.RecordCount > 0 Then
                                .TextMatrix(.Row, 3) = objdbRs("shiftid")
                                .TextMatrix(.Row, 4) = DecodeStr(objdbRs("description"))
                                .TextMatrix(.Row, 5) = Format(objdbRs("time1"), "h:mm AM/PM")
                                .TextMatrix(.Row, 6) = Format(objdbRs("time2"), "h:mm AM/PM")
                            End If
                        End If
                        .SetFocus
                    
                    Case 7
                        Command11.Cancel = False
                        txtFlex.ZOrder 0
                        txtFlex.Visible = True
                        txtFlex.Width = .CellWidth + 25
                        txtFlex.Height = .CellHeight
                        txtFlex.left = .CellLeft + .left
                        txtFlex.top = .CellTop + .top - 10
                        txtFlex.Text = .Text
                        txtFlex.SetFocus
                        
                End Select
                
            Case vbKeyDelete
                If .ColSel = 4 Then
                    If MsgBox("Delete shift entry for this day?", vbYesNo, "Confirm shift deletion...") = vbYes Then
                        .TextMatrix(.Row, 3) = ""
                        .TextMatrix(.Row, 4) = ""
                        .TextMatrix(.Row, 5) = ""
                        .TextMatrix(.Row, 6) = ""
                    End If
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex")
End Sub


Private Sub MSHFlexPeriod_EnterCell()
    Dim cSqlStmt As String, _
        nCtr As Integer

    Text2(4).Text = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 1)
    Label4.Caption = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 2)
    Text2(6).Text = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 5)
    Text2(1).Text = MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 6)
    
'    TIME_FORMAT(TIME1,'%h:%i %p')
    cSqlStmt = "select date_format(a.date,'%b %d') as `day`," & _
               "       date_format(a.date,'%a') as `dayname`," & _
               "       a.shiftid," & _
               "       ifnull(b.description,'') as shiftname, " & _
               "       ifnull(time_format(b.time1,'%h:%i %p'),'') as time1," & _
               "       ifnull(time_format(b.time2,'%h:%i %p'),'') as time2," & _
               "       if(c.description is not null and trim(a.remark)='',c.description,a.remark) as remark," & _
               "       if(c.description is null,0,1) as tag, a.date" & _
               " from di36770 a left join pa74380 b on a.shiftid=b.shiftid " & _
               " left join pa4329 c on (a.date=c.date) or (date_format(a.date,'%m %d')=date_format(c.date,'%m %d'))" & _
               " where a.empid=" & cQuote & Text1.Text & cQuote & _
               " and a.periodid=" & cQuote & MSHFlexPeriod.TextMatrix(MSHFlexPeriod.Row, 1) & cQuote & _
               " order by a.date"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, False, , True
        DoEvents
        With MSHFlexGrid1
            For nCtr = 1 To MSHFlexGrid1.Rows - 1
                If Val(.TextMatrix(nCtr, 8)) = 1 Then
                    HiLyt2 nCtr, MSHFlexGrid1, vbBlue
                Else
                    HiLyt2 nCtr, MSHFlexGrid1, IIf(UCase(Trim(.TextMatrix(nCtr, 2))) = "SUN", vbRed, vbBlack)
                End If
            
            Next nCtr
        End With
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 7) = txtFlex.Text
        
                txtFlex_LostFocus
                .SetFocus
                 
            Case vbKeyEscape
                txtFlex_LostFocus
                .SetFocus
                
        End Select
    End With
End Sub

Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
    Command11.Cancel = True
End Sub
