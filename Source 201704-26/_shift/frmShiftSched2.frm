VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{30DA1A2F-A970-4238-AC17-5773BA9DC841}#1.1#0"; "CIAXPDatePicker.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaXPFrame.ocx"
Begin VB.Form frmShiftSched2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shifting Schedule"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11835
   Begin VB.CheckBox Check2 
      Caption         =   "Exclude Close Period"
      Height          =   255
      Left            =   150
      TabIndex        =   26
      Top             =   8895
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7140
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   3705
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5085
      Left            =   3090
      TabIndex        =   4
      Top             =   3240
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8969
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
   Begin ciaXPFrame.XPFrame XPFrame5 
      Height          =   8925
      Left            =   90
      TabIndex        =   12
      Top             =   0
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   15743
      Alignment       =   2
      BorderColor     =   0
      Caption         =   " Department List "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Radius          =   20
      LicValid        =   -1  'True
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8505
         Left            =   105
         TabIndex        =   13
         Top             =   240
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   15002
         _Version        =   393216
         GridColor       =   10416117
         GridColorUnpopulated=   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
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
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   5010
      TabIndex        =   15
      Top             =   8265
      Width           =   6720
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   4725
         Picture         =   "frmShiftSched2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   150
         Picture         =   "frmShiftSched2.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "16"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   3765
         Picture         =   "frmShiftSched2.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   2805
         Picture         =   "frmShiftSched2.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1965
         Picture         =   "frmShiftSched2.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   5700
         Picture         =   "frmShiftSched2.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   1125
         Picture         =   "frmShiftSched2.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   1650
      Left            =   3090
      TabIndex        =   16
      Top             =   90
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   2910
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
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   1545
      Left            =   3090
      TabIndex        =   17
      Top             =   1725
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   2725
      BorderColor     =   0
      Caption         =   " Shift Schedule Header Information "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   300
         Left            =   2385
         TabIndex        =   23
         Top             =   825
         Width           =   450
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
         Left            =   1755
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "TXT:DEPID"
         Top             =   840
         Width           =   600
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
         Left            =   1755
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "TXT:PERIODID"
         Top             =   1140
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   300
         Left            =   2385
         TabIndex        =   20
         Top             =   1125
         Width           =   450
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
         Left            =   1755
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "TXT:SCHED_NO"
         Top             =   540
         Width           =   1080
      End
      Begin ciaXPDatePicker.XPDatePicker XPDatePicker1 
         Height          =   315
         Left            =   1755
         TabIndex        =   0
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_SCHED"
         Top             =   210
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         FormatString    =   "long date"
         MouseIcon       =   "frmShiftSched2.frx":B28E
         CalendarDayBorder=   -1  'True
         CalendarDayBorderColor=   -2147483646
         CalendarMonthBorderColor=   8421504
         LicValid        =   -1  'True
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2895
         TabIndex        =   24
         Top             =   900
         Width           =   4005
      End
      Begin VB.Label Label6 
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
         Left            =   150
         TabIndex        =   22
         Top             =   1185
         Width           =   1470
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2895
         TabIndex        =   21
         Top             =   1200
         Width           =   4005
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Created"
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
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Number"
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
         Left            =   150
         TabIndex        =   18
         Top             =   585
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmShiftSched2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmShiftSched2 --> new Shifting Schedule
' description   :   Module for Shifting Schedule by Line
' programmer    :   _-=[ srm ]=-_
' date          :   17 apr 2006

Option Explicit
    Dim nAdd As Integer, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant, _
        myArray2 As Variant, _
        myArray3 As Variant, _
        lAllDept As Boolean
    
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

Sub txtKeyDown(ByVal nMode As Integer, cString As String, oLabel As Label)
    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command3_Click
                Case 2
                    Command2_Click
            End Select
        Else
            ShowData nMode, cString, oLabel
        End If
    End If
End Sub
        
Sub ShowData(ByVal nMode As Integer, cString As String, oLabel As Label)
    Select Case nMode
        Case 1      ' --> department
            OpenQueryDNS "SELECT linename as `description` FROM DI5463 WHERE lineid=" & cQuote & cString & cQuote, objdbRs, False
        Case 2      ' --> period
            OpenQueryDNS "SELECT duration as `description` FROM PA7730 WHERE PERIODID=" & cQuote & cString & cQuote, objdbRs, False
    End Select
    If objdbRs.RecordCount > 0 Then
        oLabel.Caption = objdbRs("description")
    Else
        oLabel.Caption = ""
    End If
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
    
    cSqlStmt = "select date_start, date_end from pa7730 where periodid=" & cQuote & Text5.Text & cQuote
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
            
            RefreshGrid MSHFlexGrid2, True
            HiLyt2
            
            .Redraw = True
        End With
    End If
    
    Set oRSet = Nothing
    Set oRSet2 = Nothing
End Sub

Private Sub Check2_Click()
    MSHFlexGrid1_EnterCell
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrSched
    Dim nCtr As Integer, nCtr2 As Integer, _
        cSqlStmt As String, _
        cString As String
    
    cString = Text1.Text

    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Shifting schedule entry?", vbYesNoCancel, App.Title)
        Case vbYes
            If nAdd = 1 Then
                If lAllDept Then
                    With MSHFlexGrid1
                        For nCtr = 1 To (.Rows - 1)
                            
                            cSeries = GenerateSeries("SH_SCHED")
                            While IfExists("di546370", "di546370.SCHED_NO=" & cQuote & PadStr(cSeries, "0", 10) & cQuote)
                                cSeries = GenerateSeries("SH_SCHED")
                            Wend
                            cSeries = PadStr(cSeries, "0", 10)
                            
                            'save header
                            cSqlStmt = "insert into di546370(SCHED_NO,DATE_SCHED,PERIODID,DEPID)values(" & _
                                       cQuote & cSeries & cQuote & "," & _
                                       cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Text5.Text & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            'save detail
                            With MSHFlexGrid2
                                ShowProgress 0
                                
                                For nCtr2 = 1 To (.Rows - 1)
                                                    
                                    ShowProgress 2, (nCtr2 / (.Rows - 1)) * 100
                                    
                                    If Trim(.TextMatrix(nCtr2, 1)) <> "" Then
                                        cSqlStmt = "insert into di546373(sched_no,`date`,shiftid,a_shiftid,`remark`,seq_no)values(" & _
                                                   cQuote & cSeries & cQuote & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr2, 10), "yyyy-mm-dd") & cQuote & "," & _
                                                   cQuote & .TextMatrix(nCtr2, 3) & cQuote & "," & _
                                                   cQuote & .TextMatrix(nCtr2, 11) & cQuote & "," & _
                                                   cQuote & EncodeStr(.TextMatrix(nCtr2, 7)) & cQuote & "," & _
                                                   nCtr2 & ")"
                                        OpenQueryDNS cSqlStmt, objdbRs, True
                                        Script2File cSqlStmt
                                    End If
                                    
                                Next nCtr2
                            End With
                        
                        Next nCtr
                        
                        ShowProgress 4
                    End With
                Else
                    If IfExists("di546370", "SCHED_NO=" & cQuote & Text1.Text & cQuote) Then
                        MsgBox "Shifting Schedule Number already exists!", vbOKOnly, App.Title
                        Text1.SetFocus
                        GoTo EndSched
                    Else
                        OpenQueryDNS InsertFields(Me, "di546370"), oTempADO, True
                        Script2File InsertFields(Me, "di546370")
                        
                        Log2Audit Name, "Create Schedule Number#" & Text1.Text & " for Line ID#" & Text6.Text
                        
                        With MSHFlexGrid2
                            ShowProgress 0
                            
                            For nCtr = 1 To (.Rows - 1)
                            
                                ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                                
                                If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                                    cSqlStmt = "insert into di546373(sched_no,`date`,shiftid,a_shiftid,`remark`,seq_no)values(" & _
                                               cQuote & Text1.Text & cQuote & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
                                               cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                               cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                               cQuote & EncodeStr(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                                               nCtr & ")"
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                    Script2File cSqlStmt
                                End If
                                
                            Next nCtr
                            
                            ShowProgress 4
                            
                        End With
                    End If
                End If
            Else
                OpenQueryDNS EditField(Me, "di546370", "sched_no=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "di546370", "sched_no=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "Update Schedule #" & Text1.Text
                
                OpenQueryDNS "delete from di546373 where sched_no=" & cQuote & Text1.Text & cQuote, objdbRs, True
                Script2File "delete from di546373 where sched_no=" & cQuote & Text1.Text & cQuote
            
                With MSHFlexGrid2
                    ShowProgress 0
                    
                    For nCtr = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            cSqlStmt = "insert into di546373(sched_no,`date`,shiftid,a_shiftid,`remark`,seq_no)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                       cQuote & EncodeStr(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                                       nCtr & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                    Next nCtr
                    
                    ShowProgress 4
                    
                End With
            End If
        
        Case vbNo
            cString = ""
        
        Case vbCancel
            GoTo EndSched
            
    End Select
    
    Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
    
    If Text1.Text <> cSeries Then ResetSeries "SH_SCHED", cSeries
    
    nAdd = 0
    CtrlPanel Me, nAdd
    
    Command2.Enabled = False
    Command3.Enabled = False
    
    MSHFlexGrid1.Enabled = True
    MSHFlexGrid3.Enabled = True
    
    MSHFlexGrid1_EnterCell
    
    lAllDept = False
    
EndSched:
    Exit Sub
    
ErrSched:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
    Resume EndSched
End Sub

Private Sub Command11_Click()
    Dim cString As String
    If nAdd = 0 Then
        Unload Me
    Else
        cString = IIf(nAdd = 2, Text1.Text, "")
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            
            If Text1.Text <> cSeries Then ResetSeries "SH_SCHED", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            
            Command2.Enabled = False
            Command3.Enabled = False
            
            MSHFlexGrid1.Enabled = True
            MSHFlexGrid3.Enabled = True
            
            MSHFlexGrid1_EnterCell
            lAllDept = False
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Not lAllDept Then
        If Trim(Text6.Text) = "" Then
            MsgBox "Please define a valid department first!", vbCritical, "System Advisory"
            Text6.SetFocus
            Exit Sub
        End If
    End If
    
    frmLookup.showPopup 5, " where (periodid not in (select periodid from di546370 where depid=" & cQuote & Text6.Text & cQuote & ")) and (pclose<>1)"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text5.Text = cResult
        ShowData 2, cResult, Label14
        
        FillGrid
    Else
        Label14.Caption = ""
        SetGridColumn myArray3, MSHFlexGrid2
    End If
    Text5.SetFocus
End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 2
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text6.Text = cResult
        ShowData 1, cResult, Label20
    End If
    
    Text6.SetFocus
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmp546370( " & _
               " [sched_no] char(10),       [date_sched] date," & _
               " [PERIODID] char(5),        [DURATION] char(50)," & _
               " [DEPTNAME] char(100),      [REMARK] char(100)," & _
               " [SHIFTDESC] char(100),     [TAG] integer," & _
               " [DAY_DATE] date,           [DAY_NAME] char(20)," & _
               " [TIME1] char(15),          [TIME2] char(15)," & _
               " [REGDAY] integer,          [HOLIDAY] integer)"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmp546370", oTempADO, True
End Sub

Private Sub Command4_Click()
    On Error GoTo ErrApply
    
    Dim lProceed As Boolean, _
        nCtr As Integer, _
        nCount As Integer, _
        cSqlStmt As String, _
        cString As String, _
        oRecordSet As New ADODB.Recordset
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If
    
    If lProceed Then
        If MsgBox("Apply this Shifting Schedule entry?", vbYesNo, App.Title) = vbYes Then
            With MSHFlexGrid2
                ShowProgress 0, , 1
                
                For nCtr = 1 To .Rows - 1
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , , "Applying shifting schedule for " & Trim(Label20.Caption) & " on " & .TextMatrix(nCtr, 2)
                    
                    cSqlStmt = "select empid, concat(lastname,', ',firstname) as fullname, depid " & _
                               "from di3670 where (depid=" & cQuote & Text6.Text & cQuote & ")" & _
                               "and (((active=0) and (date_hire<=" & cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & ")) " & _
                               "    or (((active=1) or (active=3)) and (date_res>=" & cQuote & Format(.TextMatrix(1, 10), "yyyy-mm-dd") & cQuote & ")) " & _
                               "    or ((active=2) and (date_fin>=" & cQuote & Format(.TextMatrix(1, 10), "yyyy-mm-dd") & cQuote & "))) "
                    OpenQueryDNS cSqlStmt, oRecordSet, False
                    If oRecordSet.RecordCount > 0 Then
                        DoEvents
                        While Not oRecordSet.EOF
                            ' --> DTR summary here
                            cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,time1,time2,`remark`)values(" & _
                                       cQuote & oRecordSet("empid") & cQuote & "," & _
                                       cQuote & Text5.Text & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 5), "HH:MM:SS") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 6), "HH:MM:SS") & cQuote & "," & _
                                       cQuote & DecodeStr(.TextMatrix(nCtr, 7)) & cQuote & _
                                       ") on duplicate key " & _
                                       "update shiftid=" & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       " periodid=" & cQuote & Text5.Text & cQuote & "," & _
                                       " time1=" & cQuote & Format(.TextMatrix(nCtr, 5), "HH:MM:SS") & cQuote & "," & _
                                       " time2=" & cQuote & Format(.TextMatrix(nCtr, 6), "HH:MM:SS") & cQuote
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            '---------->  Alternate Time Schedule here (201703-02)
                                        '-------------> Temporary Disabled (201705-30)
'                              cSqlStmt = "insert into DI36770A(empid,periodid,`date`,shiftid,time1,time2,`remark`)values(" & _
'                                       cQuote & oRecordSet("empid") & cQuote & "," & _
'                                       cQuote & Text5.Text & cQuote & "," & _
'                                       cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & "," & _
'                                       cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
'                                       cQuote & Format(.TextMatrix(nCtr, 13), "HH:MM:SS") & cQuote & "," & _
'                                       cQuote & Format(.TextMatrix(nCtr, 14), "HH:MM:SS") & cQuote & "," & _
'                                       cQuote & DecodeStr(.TextMatrix(nCtr, 7)) & cQuote & _
'                                       ") on duplicate key " & _
'                                       "update shiftid=" & cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
'                                       " periodid=" & cQuote & Text5.Text & cQuote & "," & _
'                                       " time1=" & cQuote & Format(.TextMatrix(nCtr, 13), "HH:MM:SS") & cQuote & "," & _
'                                       " time2=" & cQuote & Format(.TextMatrix(nCtr, 14), "HH:MM:SS") & cQuote
'                            OpenQueryDNS cSqlStmt, objdbRs, True
'                            Script2File cSqlStmt
                            
                            
                            ' --> DTR detail here - 20070307
                            cSqlStmt = "update pa84650 set shiftid=" & cQuote & .TextMatrix(nCtr, 3) & cQuote & _
                                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ")" & _
                                       " and (logdate=" & cQuote & Format(.TextMatrix(nCtr, 10), "yyyy-mm-dd") & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            oRecordSet.MoveNext
                        Wend
                    End If
                    
                    cSqlStmt = "update di546373 set status=1,date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                               " where sched_no=" & cQuote & Text1.Text & cQuote & _
                               " and `date`=" & cQuote & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd") & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                    
                    nCount = nCount + 1
                    
                Next nCtr
                
                If nCount = .Rows - 1 Then
                    cSqlStmt = "update di546370 set status=1,date_post=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                               " where sched_no=" & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                    
                    Log2Audit Name, "Apply Schedule #" & Text1.Text & " to " & Text6.Text, IIf(gUserLevel <> 1, cMGR_CODE, ""), IIf(gUserLevel <> 1, cMGR_NAME, "")
                    
                    
'                    ' --> 20071005
                    cSqlStmt = "update di2340 set dtr_update=1"
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
                
                ShowProgress 4
            
            End With
        
            MSHFlexGrid3_EnterCell
        
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this shifting schedule entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Set oRecordSet = Nothing
    
    Exit Sub
    
ErrApply:
    Set oRecordSet = Nothing
    ErrorMsg Err.Number, Err.Description, "Apply Schedule #" & Text1.Text, Name
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer

    CreateTemp
    
    With MSHFlexGrid2
    
        ShowProgress 0
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            cSqlStmt = "insert into tmp546370(sched_no, date_sched, periodid, duration, deptname, " & _
                       " day_date, day_name, shiftdesc, time1, time2, remark, tag)values(" & _
                       cQuote & Text1.Text & cQuote & "," & _
                       cQuote & Format(XPDatePicker1.CurrentDate, "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & Text5.Text & cQuote & "," & _
                       cQuote & EncodeStr2(Label14.Caption) & cQuote & "," & _
                       cQuote & EncodeStr2(Label20.Caption) & cQuote & "," & _
                       cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & EncodeStr2(.TextMatrix(nCtr, 2)) & cQuote & "," & _
                       cQuote & EncodeStr2(.TextMatrix(nCtr, 4)) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 5) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & _
                       cQuote & EncodeStr2(.TextMatrix(nCtr, 7)) & cQuote & "," & _
                       .TextMatrix(nCtr, 8) & ")"
            QueryTemp cSqlStmt, objdbRs, True
            
        Next nCtr
        
        ShowProgress 3
        
        GenerateReport "Shifting Schedule", "PRV546370.RPT"
        
        ShowProgress 4
        
    End With
End Sub

Private Sub Command7_Click()
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command2.Enabled = True
    Command3.Enabled = True
    
    Label14.Caption = ""
    Label20.Caption = ""
    
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)) <> "" Then
        Text6.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)
        Text6_KeyDown vbKeyReturn, 0
    End If
    
    MSHFlexGrid1.Enabled = False
    MSHFlexGrid3.Enabled = False
    
    SetGridColumn myArray3, MSHFlexGrid2
    
    If MsgBox("Create unified Shifting Schedule for all Department?", vbYesNo, "System Advisory") = vbYes Then
        lAllDept = True
        Text1.Text = ""
        Text6.Text = ""
        Text1.Enabled = False
        Text6.Enabled = False
        
        Label20.Visible = False
        
        Command3.Enabled = False
        
        Text5.SetFocus
    Else
        cSeries = GenerateSeries("SH_SCHED")
        While IfExists("di546370", "di546370.SCHED_NO=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
            cSeries = GenerateSeries("SH_SCHED")
            Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
        Wend
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
        
        Text1.SetFocus
    End If
End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        If oTempADO("status") = 1 Then
            MsgBox "This shifting schedule had been applied as of" & vbCrLf & _
                   Format(oTempADO("date_post"), "dddd mmmm dd, yyyy") & "." & vbCrLf & vbCrLf & _
                   "You may re-apply this schedule provided that" & vbCrLf & _
                   "the affected department does not contain any " & vbCrLf & _
                   "customized shifting schedule by employee." & vbCrLf & vbCrLf & _
                   "Some entry had been disabled for security purposes.", vbInformation, "System Advisory"
        End If
        
        Lock2User Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        MSHFlexGrid1.Enabled = False
        MSHFlexGrid3.Enabled = False
        
        Command2.Enabled = oTempADO("status") <> 1
        Command3.Enabled = oTempADO("status") <> 1
        
        Text5.Enabled = oTempADO("status") <> 1
        Text6.Enabled = oTempADO("status") <> 1
        
        Text1.Enabled = False
        XPDatePicker1.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "delete from di546370 where sched_no=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from di546370 where sched_no=" & cQuote & Text1.Text & cQuote
        
        OpenQueryDNS "delete from di546373 where sched_no=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from di546373 where sched_no=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE Schedule #" & Text1.Text
        
        MSHFlexGrid1_EnterCell
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    myArray = Array("TXT:[Dep ID]:3:False", _
                    "TXT:[Department]:25:True")

    myArray2 = Array("TXT:[Sched No]:12:True", _
                     "TXT:[Period ID]:5:False", _
                     "TXT:[Duration]:30:True", _
                     "NUM:[status]:1:False", _
                     "NUM:[close]:1:False", _
                     "TXT:[Remark]:40:True")

    myArray3 = Array("TXT:[Date]:8:True", _
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
                     "TXT:[End Time]:10:True")

    Tag = nAccess_Tag
    nAdd = 0
    
'    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    OpenQueryDNS "SELECT LINEID, LINENAME FROM DI5463 ORDER BY LINEID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , True
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
    
    MSHFlexGrid1_EnterCell
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
    Dim cSqlStmt As String
    
    cSqlStmt = "select a.sched_no, a.periodid, ifnull(b.duration,'') as duration, a.status, ifnull(if(b.pclose>0,b.pclose,if(b.isprocess=1,2,0)),0) as pclose, " & _
               " if(b.pclose=1,concat('Period closed as of ',date_format(b.date_close,'%b %e, %Y')),if(b.isprocess=1,concat('Payroll processed as of ',date_format(b.date_process,'%b %e, %Y')),if(a.status=1,concat('Applied last',' ',date_format(a.date_post,'%b %e, %Y')),''))) as remark " & _
               "from di546370 a left join pa7730 b on a.periodid=b.periodid " & _
               "where (a.depid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1) & cQuote & ")" & _
               IIf(Check2.Value = vbChecked, " and (b.pclose<>1) ", "") & _
               "order by a.depid, a.periodid"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid3, myArray2, False
    Else
        SetGridColumn myArray2, MSHFlexGrid3
    End If
    MSHFlexGrid3_EnterCell
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

Private Sub MSHFlexGrid3_EnterCell()
    Dim cSqlStmt As String
    
    ClearAll Me, False, True
    
    OpenQueryDNS "select * from di546370 where sched_no=" & cQuote & MSHFlexGrid3.TextMatrix(MSHFlexGrid3.RowSel, 1) & cQuote, oTempADO, False
    GetFields Me, oTempADO
    If oTempADO.RecordCount > 0 Then
        CtrlPanel Me, nAdd, Val(MSHFlexGrid3.TextMatrix(MSHFlexGrid3.RowSel, 5)) = 0
        
        Command2.Enabled = nAdd <> 0
        Command3.Enabled = nAdd <> 0
        
        ShowData 1, Text6.Text, Label20
        ShowData 2, Text5.Text, Label14
        
        cSqlStmt = "select date_format(a.date,'%b %d') as `day`," & _
                   "       date_format(a.date,'%a') as `dayname`," & _
                   "       a.shiftid," & _
                   "       ifnull(b.description,'') as shiftname, " & _
                   "       ifnull(time_format(b.time1,'%h:%i %p'),'') as time1," & _
                   "       ifnull(time_format(b.time2,'%h:%i %p'),'') as time2," & _
                   "       if(c.description is not null and trim(a.remark)='',c.description,a.remark) as remark," & _
                   "       if(c.description is null,0,1) as tag," & _
                   "       a.status, " & _
                   "       a.date," & _
                   "       a.a_shiftid," & _
                   "       ifnull(d.description,'') as a_shiftname, " & _
                   "       ifnull(time_format(d.time1,'%h:%i %p'),'') as a_time1," & _
                   "       ifnull(time_format(d.time2,'%h:%i %p'),'') as a_time2" & _
                   " from di546373 a left join pa74380 b on a.shiftid=b.shiftid " & _
                   " left join pa74380 d on a.a_shiftid=d.shiftid " & _
                   " left join pa4329 c on (a.date=c.date) or (date_format(a.date,'%m %d')=date_format(c.date,'%m %d'))" & _
                   " where a.sched_no=" & cQuote & MSHFlexGrid3.TextMatrix(MSHFlexGrid3.RowSel, 1) & cQuote & _
                   " order by a.date"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            QueryAttach objdbRs, MSHFlexGrid2, myArray3, False
            HiLyt2
        Else
            SetGridColumn myArray3, MSHFlexGrid2
        End If
    Else
        Label14.Caption = ""
        Label20.Caption = ""
        SetGridColumn myArray3, MSHFlexGrid2
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 2, Text5.Text, Label14
    If Trim(Text5.Text) <> "" Then
        FillGrid
    Else
        Label14.Caption = ""
        SetGridColumn myArray3, MSHFlexGrid2
    End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text6.Text, Label20
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 7) = txtFlex.Text
            txtFlex_LostFocus
            MSHFlexGrid2.SetFocus
            
        Case vbKeyEscape
            txtFlex_LostFocus
            MSHFlexGrid2.SetFocus
            
    End Select
End Sub

Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
    Command11.Cancel = True
End Sub
