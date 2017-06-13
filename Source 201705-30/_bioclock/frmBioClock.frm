VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaXPFrame.ocx"
Begin VB.Form frmBioClock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bio-Clock Entry"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12615
   Icon            =   "frmBioClock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   2535
      Left            =   60
      TabIndex        =   14
      Top             =   6525
      Visible         =   0   'False
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   4471
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      Radius          =   20
      LicValid        =   -1  'True
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
         Left            =   5175
         TabIndex        =   30
         Tag             =   "1"
         ToolTipText     =   "TXT:TRAN_NO"
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10275
         TabIndex        =   29
         Top             =   90
         Width           =   225
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   450
         Left            =   9555
         TabIndex        =   28
         Tag             =   "20"
         Top             =   1905
         Width           =   915
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Delete"
         Height          =   450
         Left            =   8550
         TabIndex        =   27
         Tag             =   "19"
         Top             =   1905
         Width           =   915
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Edit"
         Height          =   450
         Left            =   7650
         TabIndex        =   26
         Tag             =   "18"
         Top             =   1905
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   450
         Left            =   6750
         TabIndex        =   25
         Tag             =   "17"
         Top             =   1905
         Width           =   915
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1635
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "TIM:TRANTIME"
         Top             =   1680
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   187367427
         UpDown          =   -1  'True
         CurrentDate     =   38381
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1635
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "DAT:TRANSDATE"
         Top             =   1365
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         _Version        =   393216
         Format          =   187367424
         CurrentDate     =   38381
      End
      Begin VB.ComboBox cmbFlex 
         Height          =   315
         ItemData        =   "frmBioClock.frx":1982
         Left            =   1635
         List            =   "frmBioClock.frx":198C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "NUM:TRANTYPE"
         Top             =   1995
         Width           =   1200
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
         Left            =   1635
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "TXT:SHIFTID"
         Top             =   930
         Width           =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   2295
         TabIndex        =   18
         Top             =   915
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   2415
         TabIndex        =   15
         Top             =   600
         Width           =   450
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
         Left            =   1635
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   615
         Width           =   750
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   1635
         TabIndex        =   0
         Tag             =   "1"
         ToolTipText     =   "DAT:LOGDATE"
         Top             =   180
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         _Version        =   393216
         Format          =   187367424
         CurrentDate     =   38381
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Time"
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
         Left            =   150
         TabIndex        =   24
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Date"
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
         Left            =   150
         TabIndex        =   23
         Top             =   1425
         Width           =   1380
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Log Date"
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
         Left            =   150
         TabIndex        =   22
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   150
         TabIndex        =   21
         Top             =   2055
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shift ID"
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
         Left            =   150
         TabIndex        =   20
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2775
         TabIndex        =   19
         Top             =   990
         Width           =   4470
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2895
         TabIndex        =   17
         Top             =   675
         Width           =   4470
      End
      Begin VB.Label Label29 
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
         Height          =   375
         Left            =   150
         TabIndex        =   16
         Top             =   660
         Width           =   1380
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   2475
         Left            =   15
         Top             =   15
         Width           =   1605
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Manual Entry"
      Height          =   660
      Left            =   11025
      Picture         =   "frmBioClock.frx":1999
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7470
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   660
      Left            =   11025
      Picture         =   "frmBioClock.frx":331B
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "21"
      Top             =   8220
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   10755
      TabIndex        =   8
      Top             =   0
      Width           =   1770
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   1710
         Width           =   1590
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   1455
         Width           =   1590
      End
      Begin VB.Line lnMinutes 
         BorderWidth     =   2
         X1              =   870
         X2              =   1035
         Y1              =   765
         Y2              =   1185
      End
      Begin VB.Line lnSeconds 
         X1              =   510
         X2              =   945
         Y1              =   555
         Y2              =   810
      End
      Begin VB.Line lnHours 
         BorderWidth     =   3
         X1              =   900
         X2              =   1170
         Y1              =   780
         Y2              =   885
      End
      Begin VB.Shape shpClock 
         Height          =   1035
         Left            =   435
         Shape           =   3  'Circle
         Top             =   315
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   285
         Picture         =   "frmBioClock.frx":4C9D
         Top             =   210
         Width           =   1200
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   8520
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   873
      LicValid        =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please place your finger on the sensor"
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
         Height          =   270
         Left            =   135
         TabIndex        =   7
         Top             =   120
         Width           =   8685
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   345
      Top             =   405
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6405
      Left            =   75
      TabIndex        =   11
      Top             =   90
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   11298
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
End
Attribute VB_Name = "frmBioClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmBioClock (formerly frmAttendance)
' description   :   Module for fingerprint scanner (U.are.U fingerprint sensor)
' programmer    :   _-=[ srm ]=-_
' date          :   17 Oct 2005 / modified 11 mar 2006

Option Explicit

Const Pi As Single = 3.141592654
Const sFactor60 As Single = Pi / 30
Const sFactor12 As Single = Pi / 6
Const sRotateFactor As Single = Pi / 2

Dim sRadius As Single, _
    nAdd As Integer, _
    oTempADO As New ADODB.Recordset, _
    myArray As Variant, _
    cSeries As String

Sub txtKeyDown(ByVal nMode As Integer, cString As String, oLabel As Label)
    If Trim(cString) = "" Then
        If nAdd <> 0 Then
            Select Case nMode
                Case 1
                    Command3_Click
                Case 2
                    Command1_Click
            End Select
        Else
            oLabel.Caption = ""
        End If
    Else
        Select Case nMode
            Case 1
                OpenQueryDNS "select concat(firstname,' ',lastname) as fullname from di3670 where empid=" & cQuote & cString & cQuote, objdbRs, False
                oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("fullname"), "")
                
                If nAdd <> 0 Then
                    OpenQueryDNS "select shiftid from di36770 where empid=" & cQuote & cString & cQuote & " and date=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote, objdbRs, False
                    Text1.Text = IIf(objdbRs.RecordCount > 0, objdbRs("shiftid"), "")
                    txtKeyDown 2, Text1.Text, Label2
                End If
                
            Case 2
'        OpenQueryDNS "SELECT * FROM PA74380 WHERE SHIFTID=" & cQuote & cResult & cQuote, objdbRs, False
                OpenQueryDNS "select `description`,time1,time2 from pa74380 where shiftid=" & cQuote & cString & cQuote, objdbRs, False
                oLabel.Caption = IIf(objdbRs.RecordCount > 0, Trim(objdbRs("description")) & "  " & Format(objdbRs("time1"), "hh:mm AMPM") & "-" & Format(objdbRs("time2"), "hh:mm AMPM"), "")
                
        End Select
    End If
End Sub

Sub ShowRecords()
    Dim cSqlStmt As String

'    myArray = Array("TXT:1[BCID]:5:True", _
'                    "TXT:2[TCID]:6:True", _
'                    "TXT:3[Emp Id]:8:True", _
'                    "TXT:4[Name]:30:True", _
'                    "TXT:5[Department]:20:False", _
'                    "TXT:6[Shift ID]:5:False", _
'                    "TXT:7[Shift]:20:True", _
'                    "DAT:8[Log Date]:12:True", _
'                    "DAT:9[Trans Date]:12:False", _
'                    "TXT:0[Trans Time]:12:True", _
'                    "TXT:1[Trantype]:1:False", _
'                    "TXT:2[Type]:6:True")

    cSqlStmt = "select a.bcid, " & _
               "       a.tcid, " & _
               "       a.empid, " & _
               "       concat(b.firstname,' ',b.lastname) as fullname, " & _
               "       c.linename, " & _
               "       a.shiftid, " & _
               "       d.description, " & _
               "       a.logdate, " & _
               "       a.transdate, " & _
               "       a.trantime, " & _
               "       a.trantype, " & _
               "       if(a.trantype=0,'In','Out') as trntype, " & _
               "       a.tran_no " & _
               "from ((pa84650 a left join di3670 b on a.empid=b.empid) " & _
               "  left join di5463 c on b.depid=c.lineid) " & _
               "  left join pa74380 d on a.shiftid=d.shiftid " & _
               "where (a.logdate=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & ")" & _
               " and (a.bcid='')" & _
               " order by a.transdate, a.trantime,a.empid"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
    
    MSHFlexGrid1_EnterCell
End Sub

Private Sub Command1_Click()
    frmLookup.showPopup 9
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text1.Text = cResult
        txtKeyDown 2, cResult, Label2
    End If
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            
            nAdd = 0
        
            ClearAll Me, False, True
            CtrlPanel Me, nAdd
            
            MSHFlexGrid1.Enabled = True
            Command2.Enabled = True
            
            ShowRecords
        End If
    End If
End Sub

Private Sub Command2_Click()
    XPFrame1.Visible = False
    MSHFlexGrid1.Height = 8415
End Sub

Private Sub Command3_Click()
    Dim cParam As String
    
    cParam = " WHERE (a.ACTIVE=0) or " & _
             " ((a.active=1) and (a.date_res =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")) or " & _
             " ((a.active=2) and (a.date_fin =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "))"
    
    frmLookup.showPopup 3, cParam
    frmLookup.Show 1

    If Trim(cResult) <> "" Then
        Text10.Text = cResult
        txtKeyDown 1, cResult, Label20
    End If
End Sub

Private Sub Command4_Click()
    Command1.Enabled = True
    Command3.Enabled = True
    
    Label2.Caption = ""
    Label20.Caption = ""
    
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd

    MSHFlexGrid1.Enabled = False
    cmbFlex.ListIndex = 0
    
    DTPicker3.SetFocus
End Sub

Private Sub Command5_Click()
    If Not isDataLock(Me.Name, Text2.ToolTipText, Text2.Text) Then
        Lock2User Me.Name, Text2.ToolTipText, Text2.Text, True
        
        MSHFlexGrid1.Enabled = False
        Command1.Enabled = True
        Command3.Enabled = True
        
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        DTPicker3.SetFocus
    End If
End Sub

Private Sub Command6_Click()
    On Error GoTo ErrDel
    
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deduction Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM pa84650 WHERE tran_no=" & cQuote & Text2.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM pa84650 WHERE tran_no=" & cQuote & Text2.Text & cQuote
        
        Log2Audit Name, "DELETE Bio-Clock Tran #" & Text2.Text & " - " & Trim(EncodeStr2(DecodeStr(Label20.Caption)))
        
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        ShowRecords
    End If
    
    Exit Sub
    
ErrDel:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Command7_Click()
    XPFrame1.Visible = True
    MSHFlexGrid1.Height = 6405
End Sub

Private Sub Command8_Click()
    On Error GoTo ErrSave
    Dim cString As String
    
    cString = Text2.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Manual Bio-Clock Time Entry?", vbYesNoCancel, "Bio-Clock Entry...")
        Case vbYes
            If nAdd = 1 Then
                cSeries = GenerateSeries("bio")
                While IfExists("pa84650", "pa84650.tran_no=" & cQuote & PadStr(cSeries, "0", 10) & cQuote)
                    cSeries = GenerateSeries("bio")
                Wend
                cSeries = PadStr(cSeries, "0", 10)
                Text2.Text = cSeries
                
                OpenQueryDNS InsertFields(Me, "pa84650"), oTempADO, True
                Script2File InsertFields(Me, "pa84650")
                
                Log2Audit Name, "ADD Bio-Clock Tran # -->" & Trim(Text2.Text)
                Log2Audit Name, "ADD Employee Name -->" & Trim(EncodeStr2(DecodeStr(Label20.Caption)))
            Else
                OpenQueryDNS EditField(Me, "pa84650", "pa84650.tran_no=" & cQuote & Text2.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "pa84650", "pa84650.tran_no=" & cQuote & Text2.Text & cQuote)
                
                Log2Audit Name, "Modify Bio-Clock Tran # -->" & Trim(Text2.Text)
            End If
        Case vbNo
            cString = ""
        
        Case vbCancel
            GoTo ErrSave
            
    End Select

    Lock2User Me.Name, Text2.ToolTipText, Text2.Text, False
    
'    If Text2.Text <> cSeries Then ResetSeries "bio", cSeries
    
    nAdd = 0

    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    MSHFlexGrid1.Enabled = True
    Command2.Enabled = True
    
    ShowRecords
    
ErrSave:
End Sub

Private Sub Form_Activate()
    lnSeconds.X1 = shpClock.left + shpClock.Width \ 2
    lnSeconds.Y1 = shpClock.top + shpClock.Height \ 2
    lnMinutes.X1 = lnSeconds.X1
    lnMinutes.Y1 = lnSeconds.Y1
    lnHours.X1 = lnSeconds.X1
    lnHours.Y1 = lnSeconds.Y1
    If ScaleWidth > ScaleHeight Then
        sRadius = shpClock.Height \ 2
    Else
        sRadius = shpClock.Width \ 2
    End If
    
    Timer1_Timer
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbKeyTab
End Sub

Private Sub Form_Load()
    Log2Audit Me.Name, "OPEN"
    
    myArray = Array("TXT:[BCID]:5:True", _
                    "TXT:[TCID]:6:True", _
                    "TXT:[Emp Id]:8:True", _
                    "TXT:[Name]:30:True", _
                    "TXT:[Department]:20:False", _
                    "TXT:[Shift ID]:5:False", _
                    "TXT:[Shift]:20:True", _
                    "DAT:[Log Date]:12:True", _
                    "DAT:[Trans Date]:12:False", _
                    "TXT:[Trans Time]:12:True", _
                    "TXT:[Trantype]:1:False", _
                    "TXT:[Type]:6:True", _
                    "TXT:[Tran #]:10:False")
    
    MSHFlexGrid1.Height = 8415
    
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
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

Private Sub MSHFlexGrid1_EnterCell()
    OpenQueryDNS "select * from pa84650 where tran_no=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 13) & cQuote, oTempADO, False
    GetFields Me, oTempADO
    If XPFrame1.Visible Then
        Command1.Enabled = nAdd <> 0
        Command3.Enabled = nAdd <> 0
        
        txtKeyDown 1, Text10.Text, Label20
        txtKeyDown 2, Text1.Text, Label2
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text10.Text) = "" Then
            Command1_Click
        Else
            txtKeyDown 2, Text1.Text, Label2
        End If
    End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text10.Text) = "" Then
            Command3_Click
        Else
            txtKeyDown 1, Text10.Text, Label20
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim t As Single
    ' Seconds
    t = Format(Time, "s")
    lnSeconds.X2 = Cos(t * sFactor60 - sRotateFactor) * sRadius + lnSeconds.X1
    lnSeconds.Y2 = Sin(t * sFactor60 - sRotateFactor) * sRadius + lnSeconds.Y1
    
    t = Format(Time, "n") + t / 60
    lnMinutes.X2 = Cos(t * sFactor60 - sRotateFactor) * (sRadius - Screen.TwipsPerPixelX * 5) + lnSeconds.X1
    lnMinutes.Y2 = Sin(t * sFactor60 - sRotateFactor) * (sRadius - Screen.TwipsPerPixelX * 5) + lnSeconds.Y1

    t = Format(Time, "h") + t / 60
    lnHours.X2 = Cos(t * sFactor12 - sRotateFactor) * (sRadius - Screen.TwipsPerPixelX * 10) + lnSeconds.X1
    lnHours.Y2 = Sin(t * sFactor12 - sRotateFactor) * (sRadius - Screen.TwipsPerPixelX * 10) + lnSeconds.Y1
    
    lblDate.Caption = Format(Date, "dd mmm yyyy")
    lblTime.Caption = Format(Time, "hh:mm:ss")
End Sub
