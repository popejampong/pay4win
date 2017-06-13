VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmpSLVL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Entry"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12720
   Begin VB.ComboBox cmbFlex2 
      Height          =   315
      ItemData        =   "frmEmpSLVL.frx":0000
      Left            =   2610
      List            =   "frmEmpSLVL.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   2535
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   5805
      TabIndex        =   25
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "Details >>"
         Height          =   375
         Left            =   5745
         TabIndex        =   31
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   30
         Top             =   630
         Width           =   375
      End
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
         Left            =   1455
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "TXT:PREP_BY"
         Top             =   645
         Width           =   660
      End
      Begin VB.TextBox Text6 
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
         Left            =   1455
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "TXT:REC_BY"
         Top             =   945
         Width           =   660
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   29
         Top             =   930
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   28
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox Text7 
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
         Left            =   1455
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   1245
         Width           =   660
      End
      Begin VB.TextBox Text8 
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
         Left            =   1455
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "TXT:NOTED_BY"
         Top             =   1545
         Width           =   660
      End
      Begin VB.CommandButton Command16 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   27
         Top             =   1530
         Width           =   375
      End
      Begin VB.TextBox Text9 
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
         Left            =   1455
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "TXT:APPR_BY"
         Top             =   1845
         Width           =   660
      End
      Begin VB.CommandButton Command17 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   26
         Top             =   1830
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   105
         TabIndex        =   42
         Top             =   690
         Width           =   1350
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Left            =   105
         TabIndex        =   41
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Noted By"
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
         Left            =   105
         TabIndex        =   40
         Top             =   1590
         Width           =   1350
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
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
         Left            =   105
         TabIndex        =   39
         Top             =   1290
         Width           =   1350
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Signatories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Recommended By"
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
         Left            =   105
         TabIndex        =   37
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   2565
         TabIndex        =   36
         Top             =   690
         Width           =   4215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   2565
         TabIndex        =   35
         Top             =   990
         Width           =   4215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   2565
         TabIndex        =   34
         Top             =   1290
         Width           =   4215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   2565
         TabIndex        =   33
         Top             =   1590
         Width           =   4215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   2565
         TabIndex        =   32
         Top             =   1890
         Width           =   4215
      End
   End
   Begin MSComCtl2.DTPicker dtFlex 
      Height          =   375
      Left            =   2610
      TabIndex        =   24
      Top             =   3315
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   56623104
      CurrentDate     =   38381
   End
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmEmpSLVL.frx":0025
      Left            =   2610
      List            =   "frmEmpSLVL.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2895
      Visible         =   0   'False
      Width           =   2355
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
      Tag             =   "1"
      ToolTipText     =   "TXT:LEAVE_NO"
      Top             =   105
      Width           =   1200
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2505
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   2010
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4710
      Left            =   60
      TabIndex        =   7
      Top             =   750
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   8308
      _Version        =   393216
      RowHeightMin    =   285
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      GridColor       =   -2147483632
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
      Height          =   885
      Left            =   2640
      TabIndex        =   20
      Top             =   5415
      Width           =   10020
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   8040
         Picture         =   "frmEmpSLVL.frx":00AB
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7080
         Picture         =   "frmEmpSLVL.frx":1A2D
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   6240
         Picture         =   "frmEmpSLVL.frx":33AF
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   5400
         Picture         =   "frmEmpSLVL.frx":4D31
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2640
         Picture         =   "frmEmpSLVL.frx":66B3
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1800
         Picture         =   "frmEmpSLVL.frx":8035
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   960
         Picture         =   "frmEmpSLVL.frx":99B7
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   9015
         Picture         =   "frmEmpSLVL.frx":B339
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   4560
         Picture         =   "frmEmpSLVL.frx":CCBB
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3600
         Picture         =   "frmEmpSLVL.frx":E63D
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   120
         Picture         =   "frmEmpSLVL.frx":FFBF
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_LEAVE"
      Top             =   420
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56623104
      CurrentDate     =   38623
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Ref No"
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
      Left            =   90
      TabIndex        =   22
      Top             =   165
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   90
      TabIndex        =   21
      Top             =   420
      Width           =   1455
   End
End
Attribute VB_Name = "frmEmpSLVL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmEmpSLVL
' description   :   Leave (SL/VL/FL) Module
' programmer    :   _-=[ srm ]=-_
' date          :   26 may 2006

Option Explicit
    Dim nAdd As Integer, _
        nLastRow As Integer, _
        cSeries As String, _
        cParam As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant

'Sub InsertToGrid cResult, nRow, MSHFlexGrid1
Sub InsertToGrid(ByVal cString As String, ByVal nRowPos As Integer, ByVal oFlexGrid As MSHFlexGrid)
    Dim cSqlStmt As String
    With MSHFlexGrid1
        If Trim(cString) <> "" Then
            .TextMatrix(nRowPos, 1) = cString
            
            cSqlStmt = "select a.empid, " & _
                       "  replace(concat(a.firstname,' ',a.lastname),CHAR(22),'" & cQuote & "') as fullname, " & _
                       "  replace(ifnull(b.posname,''),CHAR(22),'" & cQuote & "') as posname, " & _
                       "  replace(ifnull(c.linename,''),CHAR(22),'" & cQuote & "') as linename, " & _
                       "  a.sl_avail - a.sl_use as sl_avail, " & _
                       "  a.vl_avail - a.vl_use as vl_avail " & _
                       "from di3670 a " & _
                       "  left join di7670 b on a.posid=b.posid " & _
                       "  left join di5463 c on a.depid=c.lineid " & _
                       "where a.empid=" & cQuote & cString & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nRowPos, 2) = objdbRs("fullname")
                .TextMatrix(nRowPos, 3) = objdbRs("posname")
                .TextMatrix(nRowPos, 4) = objdbRs("linename")
                .TextMatrix(nRowPos, 5) = objdbRs("sl_avail")
                .TextMatrix(nRowPos, 6) = objdbRs("vl_avail")
'                .TextMatrix(nRowPos, 17) = objdbRs("ul_avail")
            End If
            
            cSqlStmt = "select (ul_avail - ul_use) as ul_avail from pa73887"
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nRowPos, 17) = objdbRs("ul_avail")
            End If
            
        End If
    End With
End Sub

Sub ShowData(cString As String, oLabel As Label)
    OpenQueryDNS "SELECT USERID,CONCAT(FIRSTNAME," & cQuote & " " & cQuote & ",LASTNAME) AS FULLNAME FROM pa2360 WHERE USERID=" & cQuote & cString & cQuote, objdbRs, False
    oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("FULLNAME"), "")
End Sub

Sub cmdClick(ByVal oTxtBox As TextBox, ByVal oLabel As Label)
    frmLookup.showPopup 1   ', " where sysuser = 1"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTxtBox.Text = cResult
        ShowData cResult, oLabel
    End If
End Sub

Sub txtKeyDown(nMode As Integer, cString As String, oLabel As Label)
    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command13_Click
                Case 2
                    Command14_Click
                Case 3
                    Command15_Click
                Case 4
                    Command16_Click
                Case 5
                    Command17_Click
            End Select
        Else
            ShowData cString, oLabel
        End If
    End If
End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
    
    
    ShowData Text5.Text, Label15
    ShowData Text6.Text, Label16
    ShowData Text7.Text, Label17
    ShowData Text8.Text, Label18
    ShowData Text9.Text, Label19
    
    Command13.Enabled = nAdd <> 0
    Command14.Enabled = nAdd <> 0
    Command15.Enabled = nAdd <> 0
    Command16.Enabled = nAdd <> 0
    Command17.Enabled = nAdd <> 0
    
    cSqlStmt = "select a.empid, " & _
               " replace(ifnull(concat(b.firstname,' ',b.lastname),''),CHAR(22),'" & cQuote & "') as fullname, " & _
               " replace(ifnull(c.posname,''),CHAR(22),'" & cQuote & "') as position, " & _
               " replace(ifnull(d.linename,''),CHAR(22),'" & cQuote & "') as linename, " & _
               " a.sl_avail, " & _
               " a.vl_avail, " & _
               " a.tag, " & _
               " if(a.tag=0,'Sick Leave',if(a.tag=1,'Vacation Leave',if(a.tag=2,'Emergency Leave',if(a.tag=3,'Maternity Leave',if(a.tag=4,'Paternity Leave',if(a.tag=5,'Force Leave','Union Leave')))))) as leave_desc, " & _
               " a.paytag, " & _
               " if(a.paytag=0,'With Pay','Without Pay') as paydesc, " & _
               " a.leave_cnt, " & _
               " a.date_start, " & _
               " a.date_end, " & _
               " a.remark, " & _
               " a.seq_no, " & _
               " a.status, " & _
               " a.ul_avail " & _
               "from PA367583 a left join di3670 b on a.empid=b.empid " & _
               " left join di7670 c on b.posid=c.posid " & _
               " left join di5463 d on b.depid=d.lineid " & _
               "where a.leave_no=" & cQuote & Text1.Text & cQuote & _
               " order by a.seq_no"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , , 1
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        ShowRecords
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrMatColorSave
    
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer
    
    If Not ChkPersonnel(Text5) Then Exit Sub
    If Not ChkPersonnel(Text6) Then Exit Sub
    If Not ChkPersonnel(Text7) Then Exit Sub
    If Not ChkPersonnel(Text8) Then Exit Sub
    If Not ChkPersonnel(Text9) Then Exit Sub
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Leave entry?", vbYesNoCancel, "Leave Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA367580", "LEAVE_NO=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Leave Reference Number already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA367580"), oTempADO, True
                    Script2File InsertFields(Me, "PA367580")
                    
                    Log2Audit Name, "ADD LEAVE_NO -->" & Trim(Text1.Text)
                
                    ShowProgress 0
                    
                    With MSHFlexGrid1
                        For nCtr = 1 To .Rows - 1
                        
                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                            
                            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                                cSqlStmt = "insert into PA367583(leave_no,date_leave,empid," & _
                                           " sl_avail,vl_avail,ul_avail,tag,paytag,leave_cnt, " & _
                                           " date_start,date_end,remark,seq_no)values(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                           Val(.TextMatrix(nCtr, 5)) & "," & _
                                           Val(.TextMatrix(nCtr, 6)) & "," & _
                                           Val(.TextMatrix(nCtr, 17)) & "," & _
                                           Val(.TextMatrix(nCtr, 7)) & "," & _
                                           Val(.TextMatrix(nCtr, 9)) & "," & _
                                           Val(.TextMatrix(nCtr, 11)) & "," & _
                                           cQuote & Format(IIf(Trim(.TextMatrix(nCtr, 12)) = "", Now, .TextMatrix(nCtr, 12)), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & Format(IIf(Trim(.TextMatrix(nCtr, 13)) = "", Now, .TextMatrix(nCtr, 13)), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & EncodeStr(.TextMatrix(nCtr, 14)) & cQuote & "," & _
                                           nCtr & ")"
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                            
                        Next nCtr
                    End With
                    
                    ShowProgress 4
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "PA367580", "LEAVE_NO=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA367580", "LEAVE_NO=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT LEAVE_NO -->" & Trim(Text1.Text)
            
                cSqlStmt = "delete from pa367583 where leave_no=" & cQuote & Text1.Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 0
                
                With MSHFlexGrid1
                    For nCtr = 1 To .Rows - 1
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            cSqlStmt = "insert into PA367583(leave_no,date_leave,empid," & _
                                       " sl_avail,vl_avail,ul_avail,tag,paytag,leave_cnt, " & _
                                       " date_start,date_end,remark,seq_no,status)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       Val(.TextMatrix(nCtr, 5)) & "," & _
                                       Val(.TextMatrix(nCtr, 6)) & "," & _
                                       Val(.TextMatrix(nCtr, 17)) & "," & _
                                       Val(.TextMatrix(nCtr, 7)) & "," & _
                                       Val(.TextMatrix(nCtr, 9)) & "," & _
                                       Val(.TextMatrix(nCtr, 11)) & "," & _
                                       cQuote & Format(IIf(Trim(.TextMatrix(nCtr, 12)) = "", Now, .TextMatrix(nCtr, 12)), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Format(IIf(Trim(.TextMatrix(nCtr, 13)) = "", Now, .TextMatrix(nCtr, 13)), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & EncodeStr(.TextMatrix(nCtr, 14)) & cQuote & "," & _
                                       nCtr & "," & _
                                       Val(.TextMatrix(nCtr, 16)) & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                    Next nCtr
                End With
                
                ShowProgress 4
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "LEAVE", cSeries
    
    Frame1.Height = 615
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "LEAVE_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO

    ShowRecords
endsave:
    Exit Sub
    
ErrMatColorSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "LEAVE", cSeries
            
            Frame1.Height = 615
            nAdd = 0
            
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "LEAVE_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            ShowRecords
        End If
    End If
End Sub

Private Sub Command13_Click()
    cmdClick Text5, Label15
    Text5.SetFocus
End Sub

Private Sub Command14_Click()
    cmdClick Text6, Label16
    Text6.SetFocus
End Sub

Private Sub Command15_Click()
    cmdClick Text7, Label17
    Text7.SetFocus
End Sub

Private Sub Command16_Click()
    cmdClick Text8, Label18
    Text8.SetFocus
End Sub

Private Sub Command17_Click()
    cmdClick Text9, Label19
    Text9.SetFocus
End Sub

Private Sub Command2_Click()
    Frame1.Height = IIf(Frame1.Height = 615, 2220, 615)
    Command2.Caption = IIf(Frame1.Height = 615, "Detail >>", "<< Hide")
End Sub

Private Sub Command4_Click()
    On Error GoTo ErrApply
    
    Dim lProceed As Boolean, _
        nCtr As Integer, _
        nCount As Integer, _
        cSqlStmt As String, _
        cString As String
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If

    If lProceed Then
        If MsgBox("Apply this Leave entry?", vbYesNo, App.Title) = vbYes Then
        
            cString = Text1.Text
            
            With MSHFlexGrid1
                For nCtr = 1 To .Rows - 1
                
                    If Val(.TextMatrix(nCtr, 16)) <> 1 Then
                        If Val(.TextMatrix(nCtr, 9)) = 0 Then
                            Select Case Val(.TextMatrix(nCtr, 7))
                                Case 0      ' --> sick leave
                                    cSqlStmt = "update di3670 set sl_use = sl_use + " & Val(.TextMatrix(nCtr, 11)) & _
                                               " where empid = " & cQuote & .TextMatrix(nCtr, 1) & cQuote
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                    Script2File cSqlStmt
                                Case 1      ' --> vacation leave
                                    cSqlStmt = "update di3670 set vl_use = vl_use + " & Val(.TextMatrix(nCtr, 11)) & _
                                               " where empid = " & cQuote & .TextMatrix(nCtr, 1) & cQuote
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                    Script2File cSqlStmt
'                                Case 6      ' --> union leave
'                                    cSqlStmt = "update pa73887 set ul_use = ul_use + " & Val(.TextMatrix(nCtr, 11))
'                                    OpenQueryDNS cSqlStmt, objdbRs, True
'                                    Script2File cSqlStmt
                            End Select
                        End If
                        
                        ' --> for union leave...
                        If Val(.TextMatrix(nCtr, 7)) = 6 Then
                            cSqlStmt = "update pa73887 set ul_use = ul_use + " & Val(.TextMatrix(nCtr, 11))
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                        cSqlStmt = "update PA367583 set status=1, " & _
                                   " date_post=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & _
                                   " where leave_no=" & cQuote & Text1.Text & cQuote & _
                                   " and seq_no=" & .TextMatrix(nCtr, 15)
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                    End If
                    
                    nCount = nCount + 1
                    
                Next nCtr
                
                If nCount = .Rows - 1 Then
                    cSqlStmt = "update pa367580 set status=1," & _
                               " date_post=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & _
                               " where leave_no=" & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
                
                oTempADO.Requery adAsyncFetch
                If Trim(cString) <> "" Then oTempADO.Find "LEAVE_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
                GetFields Me, oTempADO
            
                ShowRecords
                
            End With
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this shifting schedule entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Exit Sub
    
ErrApply:
    ErrorMsg Err.Number, Err.Description, "Apply Leave #" & Text1.Text, Name
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 11
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "LEAVE_NO='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command7_Click()
    SetGridColumn myArray, MSHFlexGrid1
    
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command13.Enabled = True
    Command14.Enabled = True
    Command15.Enabled = True
    Command16.Enabled = True
    Command17.Enabled = True
    
    Label15.Caption = ""
    Label16.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    
    dtFlex.Value = Now
    
    cSeries = GenerateSeries("LEAVE")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA367580", "PA367580.LEAVE_NO=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("LEAVE")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    Dim nCtr As Integer, _
        cSqlStmt As String
    
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Command13.Enabled = True
        Command14.Enabled = True
        Command15.Enabled = True
        Command16.Enabled = True
        Command17.Enabled = True
        
        DoEvents
        With MSHFlexGrid1
            For nCtr = 1 To .Rows - 1
                cSqlStmt = "select empid, sl_avail-sl_use as sl_avail, vl_avail-vl_use as vl_avail " & _
                           "from di3670 where empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    .TextMatrix(nCtr, 5) = objdbRs("sl_avail")
                    .TextMatrix(nCtr, 6) = objdbRs("vl_avail")
                End If
            Next nCtr
        End With
        
        dtFlex.Value = Now
        
        Text1.Enabled = False
        DTPicker3.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Emp ID]:8:True", _
                    "TXT:[Full Name]:50:True", _
                    "TXT:[Position]:30:True", _
                    "TXT:[Department]:25:True", _
                    "NUM:[SL Avail]:10:True", _
                    "NUM:[VL Avail]:10:True", _
                    "NUM:[Leave Tag]:1:False", _
                    "TXT:[Type of Leave]:25:True", _
                    "NUM:[Pay Idx]:1:False", _
                    "TXT:[Pay Tag]:15:True", _
                    "NUM:[No of Days]:12:True", _
                    "DAT:[Date Start]:15:True", _
                    "DAT:[Date End]:15:True", _
                    "TXT:[Remark]:50:True", _
                    "NUM:[Seq No]:2:False", _
                    "NUM:[Status]:1:False", _
                    "NUM:[UL Avail]:10:False")

    Tag = nAccess_Tag
    nAdd = 0
    Frame1.Height = 615
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
        
    OpenQueryDNS "SELECT * FROM PA367580 ORDER BY LEAVE_NO", oTempADO, False
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

Private Sub MSHFlexGrid1_DblClick()
    MSHFlexGrid1_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid1_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyDown
                If (nLastRow = .Row) And (nLastRow = .Rows - 1) Then
                    If (Trim(.TextMatrix(.Rows - 1, 1)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        .LeftCol = 1
                        .Col = 1
                        .ColSel = 1
                    End If
                Else
                    nLastRow = .Row
                End If

            Case vbKeyUp
                If .Rows - 1 > 1 Then
                    If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then
                        nLastRow = nLastRow - 1
                        .Rows = .Rows - 1
                    End If
                End If
                
            Case vbKeyInsert    ' --> 20050908
                If .TextMatrix(.RowSel, 1) <> "" Then
                    .AddItem "", .RowSel
                    .RowHeight(.RowSel) = 285
                    '.Row = .RowSel + 1
                    .SetFocus
                End If
        
            Case vbKeyReturn
                Select Case .ColSel
                    Case 1, 11, 14
                        If .ColSel = 11 Then
                            If Trim(.TextMatrix(.Row, 7)) = "" Then
                                MsgBox "Please specify type of leave to avail first!!!", vbCritical, "System Advisory!!!"
                                .SetFocus
                                Exit Sub
                            End If
                        End If
                        
                        Command11.Cancel = False
                        txtFlex.ZOrder 0
                        txtFlex.Visible = True
                        txtFlex.Width = .CellWidth + 25
                        txtFlex.Height = .CellHeight
                        txtFlex.left = .CellLeft + .left
                        txtFlex.top = .CellTop + .top - 10
                        txtFlex.Text = .Text
                        txtFlex.SetFocus
                    
                    Case 8
                        Command11.Cancel = False
                        cmbFlex.ZOrder 0
                        cmbFlex.Visible = True
                        cmbFlex.left = .CellLeft + .left - (cmbFlex.Width - .CellWidth)
                        cmbFlex.top = .CellTop + .top - 10
                        cmbFlex.ListIndex = IIf(Trim(.Text) = "", 0, Val(.TextMatrix(.Row, 7)))
                        cmbFlex.SetFocus
                    
                    Case 10
                        Command11.Cancel = False
                        cmbFlex2.ZOrder 0
                        cmbFlex2.Visible = True
                        cmbFlex2.left = .CellLeft + .left - (cmbFlex.Width - .CellWidth)
                        cmbFlex2.top = .CellTop + .top - 10
                        cmbFlex2.ListIndex = IIf(Trim(.Text) = "", 0, Val(.TextMatrix(.Row, 9)))
                        cmbFlex2.SetFocus
                    
                    Case 12
                        If Trim(.TextMatrix(.Row, 7)) = "" Then
                            MsgBox "Please specify type of leave to avail first!!!", vbCritical, "System Advisory!!!"
                            .SetFocus
                            Exit Sub
                        ElseIf Val(.TextMatrix(.Row, 11)) = 0 Then
                            MsgBox "Please specify number of day(s) to avail first!!!", vbCritical, "System Advisory!!!"
                            .SetFocus
                            Exit Sub
                        End If
                        
                        Command11.Cancel = False
                        dtFlex.Visible = True
                        dtFlex.left = .CellLeft + .left - (dtFlex.Width - .CellWidth)
                        dtFlex.top = .CellTop + .top - 10
                        dtFlex.Value = IIf(Trim(.Text) = "", Now, .Text)
                        dtFlex.SetFocus
                        
                End Select
            
            Case vbKeyDelete
                If (.RowSel < .Rows) Then
                    If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                        If MsgBox("Delete Record ?", vbYesNo, App.Title) = vbYes Then
                            If .Rows - 1 = 1 Then
                                .AddItem "", .Rows
                                .RowHeight(.RowSel + 1) = 285
                            End If
                            .RemoveItem .RowSel
                        End If
                    Else
                        .RemoveItem .RowSel
                    End If
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LeaveCell()
    nLastRow = MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then
        KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex") Or _
                     (Screen.ActiveForm.ActiveControl.Name <> "cmbFlex") Or _
                     (Screen.ActiveForm.ActiveControl.Name <> "cmbFlex2") Or _
                     (Screen.ActiveForm.ActiveControl.Name <> "dtFlex")
    End If
End Sub

Private Sub cmbFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 7) = cmbFlex.ListIndex
                .TextMatrix(.Row, 8) = cmbFlex.Text
                cmbFlex_LostFocus
                .SetFocus
                
            Case vbKeyEscape
                cmbFlex_LostFocus
                .SetFocus
                
        End Select
    End With
End Sub

Private Sub cmbFlex_LostFocus()
    cmbFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub cmbFlex2_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 9) = cmbFlex2.ListIndex
                .TextMatrix(.Row, 10) = cmbFlex2.Text
                cmbFlex2_LostFocus
                .SetFocus
                
            Case vbKeyEscape
                cmbFlex2_LostFocus
                .SetFocus
                
        End Select
    End With
End Sub

Private Sub cmbFlex2_LostFocus()
    cmbFlex2.Visible = False
    Command11.Cancel = True
End Sub

Private Sub dtFlex_DblClick()
    dtFlex_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nCtr As Integer, _
        nCtr2 As Integer, cSqlStmt As String
        
    With MSHFlexGrid1
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, 12) = Format(dtFlex.Value, "mm/dd/yyyy")
            For nCtr = 0 To (Val(.TextMatrix(.Row, 11)) - 1)
ikot:
                If Weekday(dtFlex.Value + nCtr + nCtr2) = vbSunday Then
                    nCtr2 = nCtr2 + 1
                    GoTo ikot
                End If
                
                ' --> check if holiday...
                cSqlStmt = "select * from pa4329 " & _
                           " where (date=" & cQuote & Format(dtFlex.Value + nCtr + nCtr2, "yyyy-mm-dd") & cQuote & ")" & _
                           " or ((month(date)=" & Month(dtFlex.Value + nCtr + nCtr2) & ") and (day(date)=" & Day(dtFlex.Value + nCtr + nCtr2) & ") and (fix_day=1))"
                OpenQueryDNS cSqlStmt, objdbRs, False
                If (objdbRs.RecordCount > 0) And (Weekday(dtFlex.Value + nCtr + nCtr2) <> vbSunday) Then
                    nCtr2 = nCtr2 + 1
                    GoTo ikot
                End If
                
            Next nCtr
            
            .TextMatrix(.Row, 13) = Format(dtFlex.Value + nCtr - 1 + nCtr2, "mm/dd/yyyy")
            
            dtFlex_LostFocus
            .SetFocus
            
        ElseIf KeyCode = vbKeyEscape Then
            dtFlex_LostFocus
            .SetFocus
        End If
    End With
End Sub

Private Sub dtFlex_LostFocus()
    dtFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cParam As String, _
        cSqlStmt As String, _
        nCtr As Integer
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                Select Case .ColSel
                    Case 1
                        For nCtr = 1 To .Rows - 1
                            If Trim(.TextMatrix(nCtr, 1)) <> "" Then cParam = cParam & cQuote & .TextMatrix(nCtr, 1) & cQuote & ","
                        Next nCtr
                        
                        If Trim(cParam) <> "" Then
                            cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
                        End If
                        
                        If Trim(txtFlex.Text) = "" Then
                            cSqlStmt = " WHERE ((a.ACTIVE=0) or " & _
                                       " ((a.active=1) and (a.date_res =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")) or " & _
                                       " ((a.active=2) and (a.date_fin =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")))" & _
                                       IIf(Trim(cParam) = "", "", " and a.empid not in " & cParam)
                            
                            frmLookup.showPopup 3, cSqlStmt
                            frmLookup.Show 1
                            If Trim(cResult) <> "" Then InsertToGrid cResult, .Row, MSHFlexGrid1
                        Else
                            nCtr = .Row
                            If Not ChkDupInGrid(txtFlex.Text, 1, MSHFlexGrid1) Then
                                cSqlStmt = "select a.empid from di3670 a " & _
                                           " WHERE ((a.ACTIVE=0) or " & _
                                           " ((a.active=1) and (a.date_res =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")) or " & _
                                           " ((a.active=2) and (a.date_fin =" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")))" & _
                                           " and a.empid=" & cQuote & txtFlex.Text & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                If objdbRs.RecordCount > 0 Then
                                    InsertToGrid txtFlex.Text, nCtr, MSHFlexGrid1
                                Else
                                    MsgBox "Invalid Employee ID entered!!!", vbCritical, "System Advisory!!!"
                                End If
                            Else
                                MsgBox "Employee ID already exist!", vbInformation, "System Advisory!!!"
                            End If
                            .Row = nCtr
                        End If
                        
                    Case 11
                        If Val(txtFlex.Text) > 0 Then
                            Select Case Val(.TextMatrix(.Row, 7))
                                Case 0      ' --> Sick Leave
                                    If Val(.TextMatrix(.Row, 9)) = 0 Then
                                        If Val(.TextMatrix(.Row, 5)) < Val(txtFlex.Text) Then
                                            MsgBox "Availment is less than the available Sick Leave!!!", vbCritical, "System Advisory!!!"
                                            Exit Sub
                                        End If
                                    End If
                                Case 1      ' --> Vacation Leave
                                    If Val(.TextMatrix(.Row, 9)) = 0 Then
                                        If Val(.TextMatrix(.Row, 6)) < Val(txtFlex.Text) Then
                                            MsgBox "Availment is less than the available Vacation Leave!!!", vbCritical, "System Advisory!!!"
                                            Exit Sub
                                        End If
                                    End If
                                Case 2      ' --> Emergency Leave
                                    If Val(txtFlex.Text) > 5 Then
                                        MsgBox "Only 5 days are allowed for Emergency Leave!!!", vbCritical, "System Advisory!!!"
                                        Exit Sub
                                    End If
                                Case 3      ' --> Maternity Leave
                                    If Val(txtFlex.Text) > 70 Then
                                        MsgBox "Maximum of 70 days only are allowed for Maternity Leave!!!", vbCritical, "System Advisory!!!"
                                        Exit Sub
                                    End If
                                    
                                Case 4      ' --> Paternity Leave
                                    If Val(txtFlex.Text) > 7 Then
                                        MsgBox "Maximum of 7 days only are allowed for Paternity Leave!!!", vbCritical, "System Advisory!!!"
                                        Exit Sub
                                    End If
                                    
                                Case 6
                                    If Val(txtFlex.Text) > Val(.TextMatrix(.Row, 17)) Then
                                        MsgBox "There are only " & Val(.TextMatrix(.Row, 17)) & " day(s) remaining Union leave to avail!!!", vbCritical, "System Advisory!!!"
                                        Exit Sub
                                    End If
                                    
'                    Sick Leave
'                    Vacation Leave
'                    Emergency Leave
'                    Maternity Leave
'                    Paternity Leave
'                    Force Leave
'                    Union Leave
                                    
                            End Select
                            .TextMatrix(.Row, 11) = Val(txtFlex.Text)
                            dtFlex_KeyDown vbKeyReturn, 0
                        End If
                        
                    Case 14
                        .TextMatrix(.Row, 14) = txtFlex.Text
                        MSHFlexGrid1_KeyDown vbKeyDown, 0
                        
                End Select
                
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

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text5.Text, Label15
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 2, Text6.Text, Label16
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 3, Text7.Text, Label17
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 4, Text8.Text, Label18
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 5, Text9.Text, Label19
End Sub
