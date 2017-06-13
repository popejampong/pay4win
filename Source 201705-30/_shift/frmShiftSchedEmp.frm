VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmShiftSchedEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shifting Schedule By Employee"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   10875
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
      Left            =   5145
      TabIndex        =   30
      Top             =   45
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4050
      TabIndex        =   25
      Text            =   "Text3"
      Top             =   4605
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtFlex2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6075
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   2490
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   1860
      TabIndex        =   21
      Top             =   690
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
      Left            =   1260
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:PERIODID"
      Top             =   690
      Width           =   585
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
      Left            =   1260
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:SHIFT_NO"
      Top             =   75
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1260
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:SHIFT_DATE"
      Top             =   375
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16515072
      CurrentDate     =   38623
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1845
      Left            =   75
      TabIndex        =   3
      Top             =   1425
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   3254
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5130
      Left            =   75
      TabIndex        =   4
      Top             =   3660
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   9049
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
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   15
      TabIndex        =   22
      Top             =   8730
      Width           =   10890
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4560
         Picture         =   "frmShiftSchedEmp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   8955
         Picture         =   "frmShiftSchedEmp.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7995
         Picture         =   "frmShiftSchedEmp.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   7155
         Picture         =   "frmShiftSchedEmp.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   6315
         Picture         =   "frmShiftSchedEmp.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2640
         Picture         =   "frmShiftSchedEmp.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1800
         Picture         =   "frmShiftSchedEmp.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   960
         Picture         =   "frmShiftSchedEmp.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   9930
         Picture         =   "frmShiftSchedEmp.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   5475
         Picture         =   "frmShiftSchedEmp.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3630
         Picture         =   "frmShiftSchedEmp.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   4
         Left            =   120
         Picture         =   "frmShiftSchedEmp.frx":11896
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   5070
      TabIndex        =   27
      Top             =   270
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16515072
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   285
      Left            =   5130
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16515072
      CurrentDate     =   38623
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   930
      Left            =   6120
      TabIndex        =   29
      Top             =   60
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shifting Schedule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   75
      TabIndex        =   26
      Top             =   3360
      Width           =   10725
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   75
      TabIndex        =   24
      Top             =   1125
      Width           =   10725
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2295
      TabIndex        =   20
      Top             =   750
      Width           =   4005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
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
      Left            =   135
      TabIndex        =   19
      Top             =   735
      Width           =   1350
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
      Left            =   135
      TabIndex        =   18
      Top             =   435
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Shift No"
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
      Left            =   135
      TabIndex        =   17
      Top             =   135
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   75
      Top             =   1065
      Width           =   10725
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   75
      Top             =   3285
      Width           =   10725
   End
End
Attribute VB_Name = "frmShiftSchedEmp"
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
' module        :   frmShiftSchedEmp
' description   :   Module for Shifting Schedule by Employee
' programmer    :   _-=[ srm ]=-_
' date          :   08 aug 2007

Option Explicit
    Dim nAdd As Integer, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant, _
        myArray2 As Variant

Sub InsertToGrid(ByVal cString As String, ByVal nRowPos As Integer, ByVal oFlexGrid As MSHFlexGrid)
    Dim cSqlStmt As String
    With MSHFlexGrid1
        If Trim(cString) <> "" Then
            .TextMatrix(nRowPos, 1) = cString
            
            cSqlStmt = "select a.empid, " & _
                       "  replace(concat(a.firstname,' ',a.lastname),CHAR(22),'" & cQuote & "') as fullname, " & _
                       "  replace(ifnull(c.linename,''),CHAR(22),'" & cQuote & "') as linename, " & _
                       "  replace(ifnull(b.posname,''),CHAR(22),'" & cQuote & "') as posname " & _
                       "from di3670 a " & _
                       "  left join di7670 b on a.posid=b.posid " & _
                       "  left join di5463 c on a.depid=c.lineid " & _
                       "where a.empid=" & cQuote & cString & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nRowPos, 2) = objdbRs("fullname")
                .TextMatrix(nRowPos, 3) = objdbRs("posname")
                .TextMatrix(nRowPos, 4) = objdbRs("linename")
            End If
        End If
    End With
End Sub

Sub HiLyt2()
    Dim nCtr As Integer
    With MSHFlexGrid2
        DoEvents
        .Redraw = False
        For nCtr = 1 To (.Rows - 1)
            .Row = nCtr
            .FillStyle = flexFillRepeat
            .Col = 1
            .ColSel = .Cols() - 1
            If Val(.TextMatrix(nCtr, 8)) = 1 Then
                .CellForeColor = vbBlue
            Else
                .CellForeColor = IIf(UCase(Trim(left(.TextMatrix(nCtr, 2), 3))) = "SUN", vbRed, vbBlack)
            End If
            .FillStyle = flexFillSingle
        Next nCtr
        .Redraw = True
    End With
End Sub

Sub FillGrid()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        nRowPos As Integer, _
        aDateInfo As Variant, _
        aShiftInfo As Variant, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset
    
    aShiftInfo = Array("", "", "", "")
    
    cSqlStmt = "select date_start, date_end from pa7730 where periodid=" & cQuote & Text2.Text & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        aDateInfo = Array(objdbRs("date_start"), objdbRs("date_end"))
        
        ' --> retrieve default shift...
        OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380 where `default`=1", oRSet2, False
        If oRSet2.RecordCount > 0 Then
            aShiftInfo(0) = oRSet2("shiftid")
            aShiftInfo(1) = DecodeStr(oRSet2("description"))
            aShiftInfo(2) = Format(oRSet2("time1"), "h:mm AM/PM")
            aShiftInfo(3) = Format(oRSet2("time2"), "h:mm AM/PM")
        End If
        
        With MSHFlexGrid2
            .Redraw = False
            
            DoEvents
            
            For nCtr = Day(aDateInfo(0)) To Day(aDateInfo(1))
                
                nRowPos = nRowPos + 1
                
                .Rows = nRowPos + 1
                .RowHeight(nRowPos) = 285
                
                .TextMatrix(nRowPos, 1) = Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "mmm dd")
                .TextMatrix(nRowPos, 2) = WeekdayName(Weekday(DateAdd("d", nRowPos - 1, aDateInfo(0))), True)
                                
                If UCase(.TextMatrix(nRowPos, 2)) <> "SUN" Then
                    .TextMatrix(nRowPos, 3) = aShiftInfo(0)
                    .TextMatrix(nRowPos, 4) = aShiftInfo(1)
                    .TextMatrix(nRowPos, 5) = aShiftInfo(2)
                    .TextMatrix(nRowPos, 6) = aShiftInfo(3)
                End If
                                
                ' --> check if date is holiday
                cSqlStmt = "select a.description from pa4329 a" & _
                           " where (a.date=" & cQuote & Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "yyyy-mm-dd") & cQuote & ") or" & _
                           " (date_format(a.date,'%m %d')=" & cQuote & Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "mmm dd") & cQuote & ")"
                OpenQueryDNS cSqlStmt, oRSet, False
                If oRSet.RecordCount > 0 Then
                    .TextMatrix(nRowPos, 7) = oRSet("description")
                    .TextMatrix(nRowPos, 8) = 1
                Else
                    .TextMatrix(nRowPos, 7) = ""
                    .TextMatrix(nRowPos, 8) = 0
                End If
            
                .TextMatrix(nRowPos, 10) = Format(DateAdd("d", nRowPos - 1, aDateInfo(0)), "yyyy-mm-dd")
            Next nCtr
            
'            RefreshGrid MSHFlexGrid2, True
            HiLyt2
            
            .Redraw = True
        End With
    End If
    
    Set oRSet = Nothing
    Set oRSet2 = Nothing
End Sub


Sub ShowData(cString As String, oLabel As Label)
    OpenQueryDNS "SELECT * FROM PA7730 where periodid=" & cQuote & cString & cQuote, objdbRs, False
    oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("duration"), "")
    DTPicker2.Value = IIf(objdbRs.RecordCount > 0, objdbRs("date_start"), Now)
    DTPicker3.Value = IIf(objdbRs.RecordCount > 0, objdbRs("date_end"), Now)
    
    If objdbRs.RecordCount > 0 Then
        Text3.Text = IIf(objdbRs("pclose") = 1, 1, IIf(objdbRs("isprocess") = 1, 2, 0))
    Else
        Text3.Text = 0
    End If
    With Label5
        Select Case Val(Text3.Text)
            Case 1  ' --> Close Period
                .Caption = "Warning!!!" & vbCrLf & _
                                 "Period is closed as of" & vbCrLf & _
                                 Format(objdbRs("date_close"), "mmmm dd, yyyy")
            Case 2  ' --> Processed Payroll
                .Caption = "Warning!!!" & vbCrLf & _
                                 "Payroll had been processed as of" & vbCrLf & _
                                 Format(objdbRs("date_process"), "mmmm dd, yyyy")
            Case Else
                .Caption = ""
        End Select
    End With
End Sub

        
Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
    
    ShowData Text2.Text, Label4
    Command2.Enabled = nAdd <> 0
    
    cSqlStmt = "select a.empid, " & _
               " replace(replace(ifnull(concat(b.firstname,' ',b.lastname),''),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as fullname, " & _
               " replace(replace(ifnull(c.posname,''),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as position, " & _
               " replace(replace(ifnull(d.linename,''),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as linename, " & _
               " ifnull(b.emp_stat,0) as emp_stat, " & _
               " ifnull(b.paystatus,0) as paystatus, " & _
               " ifnull(b.wap,0) as wap " & _
               "from PA3743 a left join di3670 b on a.empid=b.empid " & _
               " left join di7670 c on b.posid=c.posid " & _
               " left join di5463 d on b.depid=d.lineid " & _
               "where a.SHIFT_NO=" & cQuote & Text1.Text & cQuote & _
               " order by a.seq_no"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , , 1
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If

    cSqlStmt = "select date_format(a.date,'%b %d') as `day`," & _
               "       date_format(a.date,'%a') as `dayname`," & _
               "       a.shiftid," & _
               "       replace(replace(ifnull(b.description,''),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as shiftname, " & _
               "       ifnull(time_format(b.time1,'%h:%i %p'),'') as time1," & _
               "       ifnull(time_format(b.time2,'%h:%i %p'),'') as time2," & _
               "       replace(replace(if(c.description is not null and trim(a.remark)='',c.description,a.remark),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as remark," & _
               "       if(c.description is null,0,1) as tag," & _
               "       a.status, " & _
               "       a.date," & _
               "       a.a_shiftid," & _
               "       replace(replace(ifnull(d.description,''),CHAR(22),'" & cQuote & "'),CHAR(24)," & cQuote & "'" & cQuote & ") as a_shiftname, " & _
               "       ifnull(time_format(d.time1,'%h:%i %p'),'') as a_time1," & _
               "       ifnull(time_format(d.time2,'%h:%i %p'),'') as a_time2" & _
               " from PA3747 a left join pa74380 b on a.shiftid=b.shiftid " & _
               " left join pa74380 d on a.a_shiftid=d.shiftid " & _
               " left join pa4329 c on (a.date=c.date) or (date_format(a.date,'%m %d')=date_format(c.date,'%m %d'))" & _
               "where a.SHIFT_NO=" & cQuote & Text1.Text & cQuote & _
               " order by a.date"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid2, myArray2, , , , 1
        HiLyt2
    Else
        SetGridColumn myArray2, MSHFlexGrid2
    End If

End Sub
        
Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        ShowRecords
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrEmpShift
    
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Shifting Schedule By Employee entry?", vbYesNoCancel, "Leave Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("pa3740", "SHIFT_NO=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "EmpShift Reference Number already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "pa3740"), oTempADO, True
                    Script2File InsertFields(Me, "pa3740")
                    
                    Log2Audit Name, "ADD SHIFT_NO -->" & Trim(Text1.Text)
                
                    ShowProgress 0
                    
                    With MSHFlexGrid1
                        For nCtr = 1 To .Rows - 1
                        
                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                            
                            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            
                                cSqlStmt = "insert into pa3743(SHIFT_NO,SHIFT_DATE,EMPID,SEQ_NO,CMPID)values(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                           nCtr & "," & _
                                           cQuote & gCompanyID & cQuote & ")"
'                                MsgBox cSqlStmt
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                            
                        Next nCtr
                    End With
                    
                    With MSHFlexGrid2
                        For nCtr = 1 To .Rows - 1

                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100

                            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                                cSqlStmt = "insert into pa3747(SHIFT_NO,SHIFT_DATE,`date`,shiftid,a_shiftid,`remark`,seq_no,cmpid)values(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                           cQuote & EncodeStr(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                                           nCtr & "," & _
                                           cQuote & gCompanyID & cQuote & ")"
'                                MsgBox cSqlStmt
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If

                        Next nCtr
                    End With
                    
                    ShowProgress 4
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "pa3740", "SHIFT_NO=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "pa3740", "SHIFT_NO=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT SHIFT_NO -->" & Trim(Text1.Text)
            
                cSqlStmt = "delete from pa3743 where SHIFT_NO=" & cQuote & Text1.Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                cSqlStmt = "delete from pa3747 where SHIFT_NO=" & cQuote & Text1.Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 0
                
                With MSHFlexGrid1
                    For nCtr = 1 To .Rows - 1
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                        
                            cSqlStmt = "insert into pa3743(SHIFT_NO,SHIFT_DATE,EMPID,SEQ_NO,CMPID)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       nCtr & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
'                            MsgBox cSqlStmt
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                    Next nCtr
                End With
                
                With MSHFlexGrid2
                    For nCtr = 1 To .Rows - 1

                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100

                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            cSqlStmt = "insert into pa3747(SHIFT_NO,SHIFT_DATE,`date`,shiftid,a_shiftid,`remark`,seq_no,cmpid)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                       cQuote & EncodeStr(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                                       nCtr & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
'                            MsgBox cSqlStmt
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
    
    If Text1.Text <> cSeries Then ResetSeries "EMPSHIFT", cSeries
    
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "SHIFT_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO

    ShowRecords
    
endsave:
    Exit Sub
    
ErrEmpShift:
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
            
            If Text1.Text <> cSeries Then ResetSeries "EMPSHIFT", cSeries
            
            nAdd = 0
            
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SHIFT_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            ShowRecords
        End If
    End If

End Sub

Private Sub Command2_Click()
    frmLookup.showPopup 5, " where (pclose=0) and (isprocess=0)"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text2.Text = cResult
        ShowData cResult, Label4
        FillGrid
    Else
        ShowData "", Label4
        SetGridColumn myArray2, MSHFlexGrid2
    End If
    Text2.SetFocus
End Sub

Private Sub Command4_Click()
    On Error GoTo ErrApply
    
    Dim lProceed As Boolean, _
        nCtr As Integer, nCtr2 As Integer, _
        nCount As Integer, _
        cSqlStmt As String, _
        cString As String, _
        oRecordSet As New ADODB.Recordset, _
        aDateInfo As Variant
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If

    If lProceed Then
        If MsgBox("Apply this Employee Shifting  entry?", vbYesNo, App.Title) = vbYes Then
            With MSHFlexGrid1
            
'                ShowProgress 0
ShowProgress 4
                
                For nCtr = 1 To .Rows - 1
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , "Posting " & .TextMatrix(nCtr, 1)
                    
                    aDateInfo = Array("", "")
                    
                    If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                        
                        For nCtr2 = 1 To MSHFlexGrid2.Rows - 1
                        
                            ShowProgress 2, (nCtr2 / (MSHFlexGrid2.Rows - 1)) * 100 ', , "Posting " & Format(MSHFlexGrid2.TextMatrix(nCtr2, 3), "yyyy-mm-dd")
                            
    '                        myArray = Array("TXT:1[Emp ID]:8:True", _
    '                                        "TXT:2[Full Name]:50:True", _
    '                                        "TXT:3[Department]:25:True", _
    '                                        "TXT:4[Position]:30:True", _
    '                                        "NUM:5[Seq No]:2:False")
    '
    '                        myArray2 = Array("TXT:1[Date]:8:True", _
    '                                         "TXT:2[Day]:6:True", _
    '                                         "TXT:3[ShiftID]:5:False", _
    '                                         "TXT:4[Shift Info]:20:True", _
    '                                         "TXT:5[Start Time]:10:True", _
    '                                         "TXT:6[End Time]:10:True", _
    '                                         "TXT:7[Remark]:40:True", _
    '                                         "NUM:8[Holiday Tag]:1:False", _
    '                                         "NUM:9[status]:1:False", _
    '                                         "TXT:10[Date yyyy-mm-dd]:10:False", _
    '                                         "TXT:1[Alt ShiftID]:5:False", _
    '                                         "TXT:2[Alt Shift Info]:20:True", _
    '                                         "TXT:3[Start Time]:10:True", _
    '                                         "TXT:4[End Time]:10:True", _
    '                                         "NUM:5[Seq No]:2:False")
                            
                            If Trim(MSHFlexGrid2.TextMatrix(nCtr2, 3)) <> "" Then
                                
'                                If InStr(1, cDateParam, Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd")) = 0 Then
'                                    cDateParam = cDateParam & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & ","
'                                End If
                                
'                                If Trim(aDateInfo(0)) = "" Then aDateInfo(0) = cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote
'                                aDateInfo(1) = cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote
                                
                                If Trim(aDateInfo(0)) = "" Then aDateInfo(0) = Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd")
                                aDateInfo(1) = Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd")
                                
                                cSqlStmt = "(periodid = " & cQuote & Text2.Text & cQuote & ")" & _
                                           " and (date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & ")" & _
                                           " and (empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote & ")"
                                           
                                If IfExists("di36770", cSqlStmt) Then
                                    cSqlStmt = " update di36770 set shiftid = " & cQuote & MSHFlexGrid2.TextMatrix(nCtr2, 3) & cQuote & "," & _
                                               " time1=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 5), "HH:MM:SS") & cQuote & "," & _
                                               " time2=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 6), "HH:MM:SS") & cQuote & _
                                               " Where periodid = " & cQuote & Text2.Text & cQuote & " And Date = " & _
                                               cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & " And empid = " & _
                                               cQuote & .TextMatrix(nCtr, 1) & cQuote
                                Else
                                    cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,time1,time2,`remark`)values(" & _
                                               cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                               cQuote & Text2.Text & cQuote & "," & _
                                               cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & "," & _
                                               cQuote & MSHFlexGrid2.TextMatrix(nCtr2, 3) & cQuote & "," & _
                                               cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 5), "HH:MM:SS") & cQuote & "," & _
                                               cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 6), "HH:MM:SS") & cQuote & "," & _
                                               cQuote & DecodeStr(MSHFlexGrid2.TextMatrix(nCtr2, 7)) & cQuote & ")"
                                End If
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                                
                                cSqlStmt = " (empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote & ")" & _
                                           " and (logdate=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & ")"
                                
                                If IfExists("pa84650", cSqlStmt) Then
                                    ' --> DTR detail here
                                    cSqlStmt = "update pa84650 set shiftid=" & cQuote & MSHFlexGrid2.TextMatrix(nCtr2, 3) & cQuote & _
                                               " where (empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote & ")" & _
                                               " and (logdate=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & ")"
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                    Script2File cSqlStmt
                                End If
                                
                                ' --> update pa3747 here
                                cSqlStmt = "update pa3747 set status=1,date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "  where SHIFT_NO=" & cQuote & Text1.Text & cQuote & _
                                           " and date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & " and shiftid = " & cQuote & MSHFlexGrid2.TextMatrix(nCtr2, 3) & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                            
                        Next nCtr2
                        
                        ' --> update pa3747 here
                        cSqlStmt = "update pa3743 set status=1,date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "  where SHIFT_NO=" & cQuote & Text1.Text & cQuote & _
                                   " and empid=" & cQuote & .TextMatrix(nCtr, 1) & cQuote
                                    
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                    
                        ' --> re-process dtr here...
'                        myArray = Array("TXT:[Emp ID]:8:True", _
'                                        "TXT:[Full Name]:50:True", _
'                                        "TXT:[Department]:25:True", _
'                                        "TXT:[Position]:30:True", _
'                                        "NUM:[Seq No]:2:False", _
'                                        "NUM:[Emp Stat]:1:False", _
'                                        "NUM:[Pay Status]:1:False", _
'                                        "NUM:[WAP]:1:False")

                        If (Trim(aDateInfo(0)) <> "") And (Trim(aDateInfo(1)) <> "") Then
                            ComputeDays .TextMatrix(nCtr, 1), _
                                        Array(aDateInfo(0), aDateInfo(1), 0), _
                                        Array(.TextMatrix(nCtr, 6), .TextMatrix(nCtr, 8), .TextMatrix(nCtr, 7))
                        End If

                    End If
                    
                    nCount = nCount + 1
                    
                Next nCtr
                
                If nCount = .Rows - 1 Then
                    cSqlStmt = "update pa3740 set status=1,date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                               " where SHIFT_NO=" & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                    
                    Log2Audit Name, "Apply Employee shift #" & Text1.Text
                
'                    ' --> 20071005
                    cSqlStmt = "update di2340 set dtr_update=1"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
                
                
                ShowProgress 4
                
            End With
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SHIFT_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            ShowRecords
            
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this shifting schedule by employee entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Set oRecordSet = Nothing
    
    Exit Sub
    
ErrApply:
    Set oRecordSet = Nothing
    ErrorMsg Err.Number, Err.Description, "Apply EMPSHIFT #" & Text1.Text, Name

End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 14
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "SHIFT_NO='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If

End Sub

Sub CreateTemp()

 On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE TMPEMPSHIFT(" & _
               " [SHIFT_NO] char(6),   [PERIODID] char(5), " & _
               " [SHIFT_DATE] date,    [status] integer, " & _
               " [date_post] date,       [CMPNAME] char(100), " & _
               " [EMPID] char(6),        [SEQ_NO] integer, " & _
               " [FULLNAME] char(100),   [POSNAME] char(100), " & _
               " [DEPARTMENT] char(100), [DATE] date, " & _
               " [SHIFTID] char(5),      [DESCRIPTION] char(100), " & _
               " [TIME1] char(10),       [TIME2] char(10), " & _
               " [REMARK] char(50),      [A_SHIFTID] char(5), " & _
               " [A_DESC] char(100),     [A_TIME1] char(10), " & _
               " [A_TIME2] char(10),     [SEQ_NO2] integer, " & _
               " [DURATION] char(100))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM TMPEMPSHIFT", objdbRs, True
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        nCtr2 As Integer, _
        cCmpname As String, _
        oRecordSet As New ADODB.Recordset

    CreateTemp

    ShowProgress 0

    OpenQueryDNS "select * from di2660 where cmpid = " & cQuote & gCompanyID & cQuote, objdbRs, False
    cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
   
    With MSHFlexGrid2
        For nCtr = 1 To MSHFlexGrid1.Rows - 1

            ShowProgress 2, (nCtr / (MSHFlexGrid1.Rows - 1)) * 100, , , "Copying " & Trim(MSHFlexGrid1.TextMatrix(nCtr, 2)) & "..."
        
            For nCtr2 = 1 To .Rows - 1
            
                ShowProgress 2, (nCtr2 / (.Rows - 1)) * 100, , , "Copying " & Trim(.TextMatrix(nCtr2, 1)) & "..."

                cSqlStmt = " INSERT INTO TMPEMPSHIFT(SHIFT_NO,[PERIODID],[DURATION],SHIFT_DATE,CMPNAME,EMPID,[SEQ_NO],FULLNAME,[DEPARTMENT]," & _
                           " POSNAME,[DATE],SHIFTID,DESCRIPTION,TIME1,TIME2,[REMARK],A_SHIFTID,A_DESC,A_TIME1,A_TIME2,SEQ_NO2)VALUES(" & _
                            cQuote & Text1.Text & cQuote & "," & _
                            cQuote & Text2 & cQuote & "," & _
                            cQuote & Label4 & cQuote & "," & _
                            cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                            cQuote & cCmpname & cQuote & "," & _
                            cQuote & MSHFlexGrid1.TextMatrix(nCtr, 1) & cQuote & "," & _
                            nCtr & "," & _
                            cQuote & DecodeStr(EncodeStr(MSHFlexGrid1.TextMatrix(nCtr, 2))) & cQuote & "," & _
                            cQuote & DecodeStr(EncodeStr(MSHFlexGrid1.TextMatrix(nCtr, 3))) & cQuote & "," & _
                            cQuote & DecodeStr(EncodeStr(MSHFlexGrid1.TextMatrix(nCtr, 4))) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 1) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 3) & cQuote & "," & _
                            cQuote & DecodeStr(EncodeStr(.TextMatrix(nCtr2, 4))) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 5) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 6) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 7) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 11) & cQuote & "," & _
                            cQuote & DecodeStr(EncodeStr(.TextMatrix(nCtr2, 12))) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 13) & cQuote & "," & _
                            cQuote & .TextMatrix(nCtr2, 14) & cQuote & "," & _
                            nCtr2 & ")"

'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, oRecordSet, True
        
            Next nCtr2
        Next nCtr
        
        ShowProgress 3, , , , "Preparing report..."
        
        GenerateReport "", "", , True
        
'        GenerateReport "Shifting Schedule by Employee Preview", "PRV3740.RPT", , True

        ShowProgress 4
    End With

    Set oRecordSet = Nothing

End Sub

Private Sub Command7_Click()
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command2.Enabled = True
    Label4.Caption = ""

    SetGridColumn myArray, MSHFlexGrid1
    SetGridColumn myArray2, MSHFlexGrid2
    
    cSeries = GenerateSeries("EMPSHIFT")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA3740", "PA3740.SHIFT_NO=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("EMPSHIFT")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        
        Lock2User Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Command2.Enabled = True
        
        Text1.Enabled = False
        DTPicker1.SetFocus
    End If

End Sub

Private Sub Command9_Click()
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "delete from pa3740 where SHIFT_NO=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from pa3740 where SHIFT_NO=" & cQuote & Text1.Text & cQuote
        
        OpenQueryDNS "delete from pa3743 where SHIFT_NO=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from pa3743 where SHIFT_NO=" & cQuote & Text1.Text & cQuote
        
        OpenQueryDNS "delete from pa3747 where SHIFT_NO=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from pa3747 where SHIFT_NO=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE EMPSHIFT #" & Text1.Text
        
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Emp ID]:8:True", _
                    "TXT:[Full Name]:50:True", _
                    "TXT:[Department]:25:True", _
                    "TXT:[Position]:30:True", _
                    "NUM:[Seq No]:2:False", _
                    "NUM:[Emp Stat]:1:False", _
                    "NUM:[Pay Status]:1:False", _
                    "NUM:[WAP]:1:False")
                    
    myArray2 = Array("TXT:[Date]:8:True", _
                     "TXT:[Day]:6:True", _
                     "TXT:[ShiftID]:5:False", _
                     "TXT:[Shift Info]:20:True", _
                     "TXT:[Start Time]:10:True", _
                     "TXT:[End Time]:10:True", _
                     "TXT:[Remark]:40:True", _
                     "NUM:[Holiday Tag]:1:False", _
                     "NUM:[status]:1:False", _
                     "TXT:[Date yyyy-mm-dd]:10:False", _
                     "TXT:[Alt ShiftID]:5:False", _
                     "TXT:[Alt Shift Info]:20:True", _
                     "TXT:[Start Time]:10:True", _
                     "TXT:[End Time]:10:True", _
                     "NUM:[Seq No]:2:False")
                     
    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
        
    OpenQueryDNS "SELECT * FROM PA3740 ORDER BY SHIFT_NO", oTempADO, False
    GetFields Me, oTempADO

    ShowRecords
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
                If (Trim(.TextMatrix(.Rows - 1, 1)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        .LeftCol = 1
                        .Col = 1
                        .ColSel = 1
                    End If

            Case vbKeyUp
                If .Rows - 1 > 1 Then
                    If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then
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
                If .ColSel = 1 Then
                    Command11.Cancel = False
                    txtFlex2.ZOrder 0
                    txtFlex2.Visible = True
                    txtFlex2.Width = .CellWidth + 25
                    txtFlex2.Height = .CellHeight
                    txtFlex2.left = .CellLeft + .left
                    txtFlex2.top = .CellTop + .top - 10
                    txtFlex2.Text = .Text
                    txtFlex2.SetFocus
                End If
                
            Case vbKeyDelete
                If (.RowSel < .Rows) Then
                    If Trim(.TextMatrix(.RowSel, 1)) <> "" Then
                        If MsgBox("Delete Record?", vbYesNo, App.Title) = vbYes Then
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
Private Sub MSHFlexGrid1_LostFocus()
On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex2")
End Sub

Private Sub MSHFlexGrid2_DblClick()
    MSHFlexGrid2_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid2_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cSqlStmt As String
    
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid2
        Select Case KeyCode
            Case vbKeyReturn
                Select Case .ColSel
                    Case 4, 12
                        frmLookup.showPopup 9
                        frmLookup.Show 1
                        If Trim(cResult) <> "" Then
                            cSqlStmt = "select shiftid, `description`, time1, time2 from pa74380 where shiftid=" & cQuote & cResult & cQuote
                            OpenQueryDNS cSqlStmt, objdbRs, False
                            If objdbRs.RecordCount > 0 Then
                                .TextMatrix(.Row, IIf(.ColSel = 4, 3, 11)) = objdbRs("shiftid")
                                .TextMatrix(.Row, IIf(.ColSel = 4, 4, 12)) = DecodeStr(objdbRs("description"))
                                .TextMatrix(.Row, IIf(.ColSel = 4, 5, 13)) = Format(objdbRs("time1"), "h:mm AM/PM")
                                .TextMatrix(.Row, IIf(.ColSel = 4, 6, 14)) = Format(objdbRs("time2"), "h:mm AM/PM")
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
                If (.ColSel = 4) Or (.ColSel = 12) Then
                    If MsgBox("Delete shift entry for this day?", vbYesNo, "Confirm shift deletion...") = vbYes Then
                        .TextMatrix(.Row, IIf(.ColSel = 4, 3, 11)) = ""
                        .TextMatrix(.Row, IIf(.ColSel = 4, 4, 12)) = ""
                        .TextMatrix(.Row, IIf(.ColSel = 4, 5, 13)) = ""
                        .TextMatrix(.Row, IIf(.ColSel = 4, 6, 14)) = ""
                    End If
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid2_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex")
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text2.Text) = "" Then
            Command2_Click
        Else
            OpenQueryDNS "SELECT * FROM PA7730 where periodid=" & cQuote & Text2.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                ShowData Text2.Text, Label4
                FillGrid
            Else
                ShowData "", Label4
                SetGridColumn myArray2, MSHFlexGrid2
            End If
        End If
    End If
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)

    With MSHFlexGrid2
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

Private Sub txtFlex2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cParam As String, _
        cSqlStmt As String, _
        nCtr As Integer
        
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                If .ColSel = 1 Then
                        If Trim(txtFlex2.Text) = "" Then
                        
'    myArray = Array("TXT:[Emp ID]:8:True", _
'                    "TXT:[Full Name]:50:True", _
'                    "TXT:[Department]:25:True", _
'                    "TXT:[Position]:30:True", _
'                    "NUM:[Seq No]:2:False")
                    
                            cParam = ""
                            For nCtr = 1 To (.Rows - 1)
                                If Trim(.TextMatrix(nCtr, 1)) <> "" Then cParam = cParam & cQuote & .TextMatrix(nCtr, 1) & cQuote & ","
                            Next nCtr
                            
                            If Trim(cParam) <> "" Then
                                cParam = " and (a.empid not in (" & left(cParam, Len(cParam) - 1) & "))"
                            End If
                        
                            cSqlStmt = " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")))) or " & _
                                       "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "))))" & _
                                       "    or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "))"
                            
                            frmLookup.showPopup 3, cSqlStmt & IIf(Trim(cParam) <> "", cParam, "")
                            frmLookup.Combo1.ListIndex = 2
                            frmLookup.Show 1
                            
                            If Not ChkDupInGrid(cResult, 1, MSHFlexGrid1) Then
                                If Trim(cResult) <> "" Then InsertToGrid cResult, .Row, MSHFlexGrid1
                            Else
                                MsgBox "Employee ID already exist!", vbInformation, "System Advisory!!!"
                            End If
                        Else
                            nCtr = .Row
                            If Not ChkDupInGrid(txtFlex2.Text, 1, MSHFlexGrid1) Then
                                cSqlStmt = "select a.empid from di3670 a " & _
                                           " where ((a.active=1) and ((a.date_res between " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ")))) or " & _
                                           "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "))))" & _
                                           "    or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "))" & _
                                           " and a.empid=" & cQuote & txtFlex2.Text & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                If objdbRs.RecordCount > 0 Then
                                    InsertToGrid txtFlex2.Text, nCtr, MSHFlexGrid1
                                Else
                                    MsgBox "Invalid Employee ID entered!!!", vbCritical, "System Advisory!!!"
                                End If
                            Else
                                MsgBox "Employee ID already exist!", vbInformation, "System Advisory!!!"
                            End If
                            .Row = nCtr
                        End If
                End If
                
                txtFlex2_LostFocus
                .SetFocus
                
            Case vbKeyEscape
                txtFlex2_LostFocus
                .SetFocus
        End Select
    End With
End Sub

Private Sub txtFlex2_LostFocus()
    txtFlex2.Visible = False
    Command11.Cancel = True
End Sub
