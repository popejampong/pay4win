VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Object = "{083C8784-F106-4CC2-9930-876218A6B74C}#1.1#0"; "ciaXPButton.ocx"
Begin VB.Form frmBioProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download and Upload Process"
   ClientHeight    =   8295
   ClientLeft      =   720
   ClientTop       =   540
   ClientWidth     =   14205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   9510
      TabIndex        =   29
      Top             =   5835
      Visible         =   0   'False
      Width           =   2220
      Begin VB.CommandButton Command7 
         Caption         =   "Ca&ncel"
         Height          =   450
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1050
         Width           =   1965
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Text File"
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   615
         Width           =   1965
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Timekeeper File"
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   1965
      End
   End
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmBioProcess.frx":0000
      Left            =   255
      List            =   "frmBioProcess.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2565
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtFlex 
      Height          =   375
      Left            =   270
      TabIndex        =   26
      Top             =   2070
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580608
      CurrentDate     =   38381
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   2325
      Left            =   6015
      TabIndex        =   11
      Top             =   45
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4101
      LicValid        =   -1  'True
      Begin ciaXPPanel.XPPanel XPPanel2 
         Height          =   30
         Left            =   90
         TabIndex        =   23
         Top             =   885
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   53
         HasLeftBorder   =   0   'False
         HasRightBorder  =   0   'False
         HasBottomBorder =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton.XPButton XPButton1 
         Height          =   525
         Left            =   6480
         TabIndex        =   16
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   926
         Caption         =   "S&how Detail >>"
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Index           =   5
         Left            =   1185
         TabIndex        =   24
         Top             =   1050
         Width           =   4905
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Index           =   4
         Left            =   1185
         TabIndex        =   22
         Top             =   2025
         Width           =   4905
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Accessed"
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
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2025
         Width           =   1575
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Index           =   3
         Left            =   1185
         TabIndex        =   20
         Top             =   1725
         Width           =   4905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified"
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
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1725
         Width           =   1575
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Index           =   2
         Left            =   1185
         TabIndex        =   18
         Top             =   1425
         Width           =   4905
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Created"
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1425
         Width           =   1575
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename"
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
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   15
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Index           =   1
         Left            =   1185
         TabIndex        =   14
         Top             =   465
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename"
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
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   465
         Width           =   1575
      End
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   285
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1605
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6570
      Top             =   7485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6450
      Left            =   75
      TabIndex        =   2
      Top             =   915
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11377
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
   Begin ciaXPPanel.XPPanel XPPanel6 
      Height          =   810
      Left            =   60
      TabIndex        =   3
      Top             =   45
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1429
      LicValid        =   -1  'True
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
         ItemData        =   "frmBioProcess.frx":0017
         Left            =   720
         List            =   "frmBioProcess.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2895
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
         Left            =   735
         TabIndex        =   1
         Top             =   420
         Width           =   3390
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort"
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
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   465
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   9510
      TabIndex        =   7
      Top             =   7320
      Width           =   4635
      Begin VB.CommandButton Command3 
         Caption         =   "&Refresh"
         Height          =   660
         Left            =   1935
         Picture         =   "frmBioProcess.frx":004C
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Download"
         Height          =   660
         Left            =   960
         Picture         =   "frmBioProcess.frx":0916
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "20"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Upload"
         Height          =   660
         Left            =   90
         Picture         =   "frmBioProcess.frx":2298
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "17"
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   3660
         Picture         =   "frmBioProcess.frx":3C1A
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "21"
         Top             =   165
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmBioProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmBioProcess
' description   :   module to upload raw data from bio-clock...
' programmer    :   _-=[ srm ]=-_
' date          :   01 March 2006

Option Explicit
    Dim nAdd As Integer, _
        myArray As Variant, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        lProceed As Boolean, _
        lTimeKeeper As Boolean


Sub RebuildDTR(ByVal nMode As Integer, Optional ByVal cValue As String = "")
    Dim oDateRSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        aDateInfo As Variant, _
        aTimeInfo As Variant
    
    Select Case nMode
        Case 0      ' --> per employee
            cSqlStmt = "SELECT min(date_start) as dStart, curdate() as dEnd FROM PA7730 WHERE (PCLOSE=0) and (DATE_START < CURDATE())"
            OpenQueryDNS cSqlStmt, oDateRSet, False
            If oDateRSet.RecordCount > 0 Then
                aDateInfo = Array(oDateRSet("dStart"), oDateRSet("dEnd"))
                
                cSqlStmt = "select a.EMPID, b.emp_stat, b.wap, b.paystatus, a.PERIODID, a.DATE, a.SHIFTID, " & _
                           "  a.reg_hr + a.reg_ot_hr + a.sa_reg_ot + a.nd_hr + a.nd_ot_hr + a.sa_nd_ot + a.sun_hr + a.sun_ot_hr as tot_hr " & _
                           "from di36770 a left join di3670 b on a.empid=b.empid " & _
                           "where (a.empid=" & cValue & ") " & _
                           " and (a.date between " & cQuote & Format(aDateInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(aDateInfo(1), "yyyy-mm-dd") & cQuote & ")"
                OpenQueryDNS cSqlStmt, oTempADO, False
                If oTempADO.RecordCount > 0 Then
                    
                    While Not oTempADO.EOF
                    
                        DoEvents
                        
                        If (oTempADO("tot_hr") = 0) And (Trim(oTempADO("shiftid")) <> "") Then
                            aTimeInfo = ComputeDays(oTempADO("empid"), _
                                                    Array(oTempADO("date"), oTempADO("date"), 0), _
                                                    Array(oTempADO("emp_stat"), oTempADO("wap"), oTempADO("paystatus")))
                            cSqlStmt = "update di36770 set reg_hr = " & aTimeInfo(0) * 8 & "," & _
                                       " reg_ot_hr = " & aTimeInfo(1) & "," & _
                                       " sa_reg_ot = " & aTimeInfo(2) & "," & _
                                       " nd_hr = " & aTimeInfo(3) * 8 & "," & _
                                       " nd_ot_hr = " & aTimeInfo(4) & "," & _
                                       " sa_nd_ot = " & aTimeInfo(12) & "," & _
                                       " sun_hr = " & aTimeInfo(5) & "," & _
                                       " sun_ot_hr = " & aTimeInfo(6) & "," & _
                                       " remark = " & cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf(aTimeInfo(11) = 1, "No Entry or Absent", "On Leave"), "")) & cQuote & _
                                       " where (empid = " & cQuote & oTempADO("empid") & cQuote & ")" & _
                                       " and (date = " & cQuote & Format(oTempADO("date"), "yyyy-mm-dd") & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If
                        
                        oTempADO.MoveNext
                        
                    Wend
                    
                End If
            End If
        
    End Select
    
    Set oDateRSet = Nothing
End Sub

Sub ShowFileInfo(cFilename As String)
    Dim oTextFile As New FileSystemObject, _
        oFile As File, _
        nCtr As Integer
    
    If Trim(cFilename) <> "" Then
        Set oFile = oTextFile.GetFile(cFilename)
        
        lblFile(0).Caption = oFile.Name
        lblFile(1).Caption = left(oFile.Path, Len(oFile.Path) - Len(oFile.Name) - 1)
        lblFile(2).Caption = Format(oFile.DateCreated, "dddd, mmmm d, yyyy hh:mm:ss AMPM")
        lblFile(3).Caption = Format(oFile.DateLastModified, "dddd, mmmm d, yyyy hh:mm:ss AMPM")
        lblFile(4).Caption = Format(oFile.DateLastAccessed, "dddd, mmmm d, yyyy hh:mm:ss AMPM")
        lblFile(5).Caption = oFile.Type
    Else
        For nCtr = 0 To 5
            lblFile(nCtr) = ""
        Next nCtr
    End If
    
    Set oTextFile = Nothing
    Set oFile = Nothing
End Sub

Private Sub CheckGrid(Optional lSingle As Boolean = False, Optional nRow As Integer = 0)
    Dim nRowPos As Integer, _
        nDiffTime As Integer, _
        cSqlStmt As String
    
    With MSHFlexGrid1
    
        If lSingle Then
        
            If (.TextMatrix(nRow, 3) = "") Then
                HiLyt2 nRow, MSHFlexGrid1, vbRed
                .TextMatrix(nRow, 13) = "Undefined Employee"
                .TextMatrix(nRow, 14) = 1
            ElseIf (.TextMatrix(nRow, 6) = "") Then
                HiLyt2 nRow, MSHFlexGrid1, vbBlue
                .TextMatrix(nRow, 13) = "Undefined Shifting schedule"
                .TextMatrix(nRow, 14) = 1
            ElseIf .TextMatrix(nRow, 21) <> "" Then
                nDiffTime = Abs(DateDiff("n", .TextMatrix(nRow, 23), .TextMatrix(nRow, 15))) / 2
                If DateDiff("n", .TextMatrix(nRow, 15), .TextMatrix(nRow, 23)) > 0 Then
                    If DateDiff("n", DateAdd("n", nDiffTime, .TextMatrix(nRow, 15)), .TextMatrix(nRow, 10)) > 0 Then
                        .TextMatrix(nRow, 6) = .TextMatrix(nRow, 21)
                        .TextMatrix(nRow, 7) = .TextMatrix(nRow, 22)
                        .TextMatrix(nRow, 15) = .TextMatrix(nRow, 23)
                        .TextMatrix(nRow, 16) = .TextMatrix(nRow, 24)
                    End If
                Else
                    If DateDiff("n", DateAdd("n", nDiffTime, .TextMatrix(nRow, 23)), .TextMatrix(nRow, 10)) < 0 Then
                        .TextMatrix(nRow, 6) = .TextMatrix(nRow, 21)
                        .TextMatrix(nRow, 7) = .TextMatrix(nRow, 22)
                        .TextMatrix(nRow, 15) = .TextMatrix(nRow, 23)
                        .TextMatrix(nRow, 16) = .TextMatrix(nRow, 24)
                    End If
                End If
                HiLyt2 nRow, MSHFlexGrid1, vbBlack
            Else
                HiLyt2 nRow, MSHFlexGrid1, vbBlack
            End If
            
            ' --> 20070416
            If Not lTimeKeeper Then
                If Trim(.TextMatrix(nRow, 6)) <> "" Then
                    If ((DateDiff("n", DateAdd("h", -4, .TextMatrix(nRow, 16)), .TextMatrix(nRow, 10)) >= 0) And (DateDiff("n", .TextMatrix(nRow, 10), .TextMatrix(nRow, 16)) >= 0)) Or _
                       (DateDiff("n", .TextMatrix(nRow, 16), .TextMatrix(nRow, 10)) > 0) Then
                       .TextMatrix(nRow, 11) = 1
                       .TextMatrix(nRow, 12) = "Out"
                    End If
                End If
            End If
            
            If Trim(.TextMatrix(nRow, 6)) <> "" Then
                If Val(.TextMatrix(nRow, 11)) = 1 Then
                    If DateDiff("n", .TextMatrix(nRow, 15), .TextMatrix(nRow, 10)) < 0 Then
                        cSqlStmt = "select `date`, shiftid, description, time1, time2  " & _
                                   "from di36770 " & _
                                   "where (empid=" & cQuote & .TextMatrix(nRow, 3) & cQuote & ")" & _
                                   " and (date < " & cQuote & Format(.TextMatrix(nRow, 8), "yyyy-mm-dd") & cQuote & ")" & _
                                   " and (trim(shiftid)<>'')" & _
                                   " order by date desc limit 1"
                        OpenQueryDNS cSqlStmt, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                            .TextMatrix(nRow, 6) = objdbRs("shiftid")
                            .TextMatrix(nRow, 7) = objdbRs("description")
                            .TextMatrix(nRow, 8) = Format(objdbRs("date"), "yyyy-mm-dd")
                            .TextMatrix(nRow, 15) = objdbRs("time1")
                            .TextMatrix(nRow, 16) = objdbRs("time2")
                        End If
                    End If
                End If
            End If
            
        Else
        
            ShowProgress 0
            
            For nRowPos = 1 To (.Rows - 1)
                
                DoEvents
                
                .TopRow = nRowPos
                ShowProgress 2, (nRowPos / (.Rows - 1)) * 100
                If (.TextMatrix(nRowPos, 3) = "") Then
                    HiLyt2 nRowPos, MSHFlexGrid1, vbRed
                    .TextMatrix(nRowPos, 13) = "Undefined Employee"
                    .TextMatrix(nRowPos, 14) = 1
                ElseIf (.TextMatrix(nRowPos, 6) = "") Then
                    HiLyt2 nRowPos, MSHFlexGrid1, vbBlue
                    .TextMatrix(nRowPos, 13) = "Undefined Shifting schedule"
                    .TextMatrix(nRowPos, 14) = 1
                ElseIf .TextMatrix(nRowPos, 21) <> "" Then
                    If Val(.TextMatrix(nRowPos, 11)) = 0 Then
                        nDiffTime = Abs(DateDiff("n", .TextMatrix(nRowPos, 23), .TextMatrix(nRowPos, 15))) / 2
                        If DateDiff("n", .TextMatrix(nRowPos, 15), .TextMatrix(nRowPos, 23)) > 0 Then
                            If DateDiff("n", DateAdd("n", nDiffTime, .TextMatrix(nRowPos, 15)), .TextMatrix(nRowPos, 10)) > 0 Then
                                .TextMatrix(nRowPos, 6) = .TextMatrix(nRowPos, 21)
                                .TextMatrix(nRowPos, 7) = .TextMatrix(nRowPos, 22)
                                .TextMatrix(nRowPos, 15) = .TextMatrix(nRowPos, 23)
                                .TextMatrix(nRowPos, 16) = .TextMatrix(nRowPos, 24)
                            End If
                        Else
                            If DateDiff("n", DateAdd("n", nDiffTime, .TextMatrix(nRowPos, 23)), .TextMatrix(nRowPos, 10)) < 0 Then
                                .TextMatrix(nRowPos, 6) = .TextMatrix(nRowPos, 21)
                                .TextMatrix(nRowPos, 7) = .TextMatrix(nRowPos, 22)
                                .TextMatrix(nRowPos, 15) = .TextMatrix(nRowPos, 23)
                                .TextMatrix(nRowPos, 16) = .TextMatrix(nRowPos, 24)
                            End If
                        End If
                    End If
                    HiLyt2 nRowPos, MSHFlexGrid1, vbBlack
                Else
                    HiLyt2 nRowPos, MSHFlexGrid1, vbBlack
                End If
            
                ' --> 20070416
                If Not lTimeKeeper Then
                    If Trim(.TextMatrix(nRowPos, 6)) <> "" Then
                        If ((DateDiff("n", DateAdd("h", -4, .TextMatrix(nRowPos, 16)), .TextMatrix(nRowPos, 10)) >= 0) And (DateDiff("n", .TextMatrix(nRowPos, 10), .TextMatrix(nRowPos, 16)) >= 0)) Or _
                           (DateDiff("n", .TextMatrix(nRowPos, 16), .TextMatrix(nRowPos, 10)) > 0) Then
                           .TextMatrix(nRowPos, 11) = 1
                           .TextMatrix(nRowPos, 12) = "Out"
                        End If
                    End If
                End If
            
                If Trim(.TextMatrix(nRowPos, 6)) <> "" Then
                    If Val(.TextMatrix(nRowPos, 11)) = 1 Then
                        If DateDiff("n", .TextMatrix(nRowPos, 15), .TextMatrix(nRowPos, 10)) < 0 Then
                            cSqlStmt = "select `date`, shiftid, description, time1, time2  " & _
                                       "from di36770 " & _
                                       "where (empid=" & cQuote & .TextMatrix(nRowPos, 3) & cQuote & ")" & _
                                       " and (date < " & cQuote & Format(.TextMatrix(nRowPos, 8), "yyyy-mm-dd") & cQuote & ")" & _
                                       " and (trim(shiftid)<>'')" & _
                                       " order by date desc limit 1"
                            OpenQueryDNS cSqlStmt, objdbRs, False
                            If objdbRs.RecordCount > 0 Then
                                .TextMatrix(nRowPos, 6) = objdbRs("shiftid")
                                .TextMatrix(nRowPos, 7) = objdbRs("description")
                                .TextMatrix(nRowPos, 8) = Format(objdbRs("date"), "yyyy-mm-dd")
                                .TextMatrix(nRowPos, 15) = objdbRs("time1")
                                .TextMatrix(nRowPos, 16) = objdbRs("time2")
                            End If
                        End If
                    End If
                End If
                
            Next nRowPos
            
            ShowProgress 4
            
        End If
        
    End With
    
End Sub

Sub ClearGrid(ByVal oFlexGrid As MSHFlexGrid, ByVal nColPos As Integer)
    Dim nCtr As Integer
    
    With oFlexGrid
        ShowProgress 0
        
        .Visible = False
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
        
        RefreshGrid MSHFlexGrid1, True
        
        .Redraw = True
        .Visible = True
        
        ShowProgress 4
        
    End With
End Sub

Sub CreateBackup(ByVal oBackupDB As ADODB.Connection, ByVal cDateBckUp As String)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String

        cSqlStmt = "create table " & "timecard" & cDateBckUp & _
                   "([timecardid] char(10),     [employeeid] char(10), " & _
                   " [trandate] char(8),        [trantime] char(6), " & _
                   " [trantype] char(1),        [hash] char(20), " & _
                   " [stationid] char(10),      [accessidno] char(10), " & _
                   " [machineno] char(10),      [verifymode] char(5), " & _
                   " [field1] char(20),         [tag] integer, " & _
                   " [date_create] date )"
                   
    oBackupDB.Execute cSqlStmt
    While oBackupDB.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM " & "timecard" & cDateBckUp
    Set objdbRs = oBackupDB.Execute(cSqlStmt)
End Sub

Sub Save2Backup(oTempDB As ADODB.Connection)
    Dim cSqlStmt As String, _
        cString As String, _
        oRset1 As New ADODB.Recordset
        
    cSqlStmt = "Select date_end From pa7730 " & _
               "Where pclose = 1 order by date_end desc limit 1"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cString = Format(objdbRs("date_end"), "yyyymmdd")
    Else
        Exit Sub
    End If

    ShowProgress 0
    
    cSqlStmt = "select * from timecard where trandate <= " & cQuote & cString & cQuote
    Set oRset1 = oTempDB.Execute(cSqlStmt)
    If oRset1.RecordCount > 0 Then
    
        ' --> create backup first...
        CreateBackup oTempDB, Format(Now, "yymmdd")
    
        DoEvents
    
        While Not oRset1.EOF
        
            DoEvents
            
            ShowProgress 2, (oRset1.AbsolutePosition / oRset1.RecordCount) * 100

            cSqlStmt = "insert into timecard" & Format(Now, "yymmdd") & "(timecardid, employeeid, trandate, trantime, trantype, " & _
                       " hash, stationid, accessidno, machineno, verifymode, tag)values(" & _
                       cQuote & oRset1("timecardid") & cQuote & "," & _
                       cQuote & oRset1("employeeid") & cQuote & "," & _
                       cQuote & oRset1("trandate") & cQuote & "," & _
                       cQuote & oRset1("trantime") & cQuote & "," & _
                       cQuote & oRset1("trantype") & cQuote & "," & _
                       cQuote & oRset1("hash") & cQuote & "," & _
                       cQuote & oRset1("stationid") & cQuote & "," & _
                       cQuote & oRset1("accessidno") & cQuote & "," & _
                       cQuote & oRset1("machineno") & cQuote & "," & _
                       cQuote & oRset1("verifymode") & cQuote & "," & _
                       cQuote & oRset1("tag") & cQuote & ")"
            Set objdbRs = oTempDB.Execute(cSqlStmt)

            oRset1.MoveNext
            
        Wend
        
        ShowProgress 3, , , "Please wait...", "Clearing old DTR file..."
        cSqlStmt = "delete from timecard where trandate <= " & cQuote & cString & cQuote
        Set objdbRs = oTempDB.Execute(cSqlStmt)

    End If
    
    ShowProgress 4
    
    Set oRset1 = Nothing
End Sub

Sub UploadTxtFile(oTmpTxt As ADODB.Connection, cFilename As String, cTimeKeeperPath As String)
    On Error GoTo ErrUpload
    Dim myTextFile As New clsReadTextFile       ' Create a new class internal to your program.
    Dim intError As Integer
    
'    Dim oTmpTxt As New ADODB.Connection

    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cString As String, _
        cParam As Variant, _
        nCtr As Integer, _
        nCtr2 As Integer
        
    With myTextFile
        .FileName = cFilename            ' set the file name to read
        .NoBlankLines = True               ' Don't return blank lines!
        .CountOnlyNonBlankLines = False    ' Count all lines regardless of whether or not they are returned
        .StripLeadingSpaces = False        ' leave any leading spaces.
        .StripTrailingSpaces = False       ' leave the trailing spaces too!
        .StripNulls = True                 ' eliminate Chr$(0)'s
        .OnlyAlphaNumericCharacters = True ' don't send me any characters I can't use!
        
        ShowProgress 0
        'nCtr2 = 0
        intError = .cfOpenFile  ' open the file, return any errors in doing so (class doesn't handle them!)
        If intError = 0 Then
           ' no error in opening the file has occured
            While Not .EndOfFile                 ' Watch for the end of the file
            
                .csGetALine                       ' Tells the class to go to a new line
                cString = .Text
                ReDim cParam(99)

                nCtr = 0
                While InStr(1, cString, ",") > 0
                    cParam(nCtr) = left(cString, InStr(1, cString, ",") - 1)
                    cParam(nCtr) = IIf(InStr(1, cParam(nCtr), cQuote) = 0, cQuote, "") & cParam(nCtr) & IIf(InStr(1, cParam(nCtr), cQuote) = 0, cQuote, "")
                    cString = Mid(cString, InStr(1, cString, ",") + 1, Len(cString) - InStr(1, cString, ","))
                    nCtr = nCtr + 1
                Wend
                If Trim(cString) <> "" Then
                    cParam(nCtr) = IIf(InStr(1, cString, cQuote) = 0, cQuote, "") & cString & IIf(InStr(1, cString, cQuote) = 0, cQuote, "")
                    cSqlStmt = " select * from timecard where accessidno = " & Replace(cParam(1), cQuote, "") & " and " & _
                               " trandate = " & cQuote & Format(Replace(cParam(4), cQuote, ""), "yyyymmdd") & cQuote & " and " & _
                               " trantime = " & cQuote & Format(Replace(cParam(5), cQuote, ""), "hhmmss") & cQuote
                    Set objdbRs = oTmpTxt.Execute(cSqlStmt)
                    If Not objdbRs.RecordCount > 0 Then
                        cSqlStmt = "insert into timecard([machineno], [stationid], [accessidno], [field1], [trantype], [trandate], [trantime], [employeeid], [hash], [verifymode], [Tag])values(" & _
                                    cParam(0) & "," & _
                                    cParam(0) & "," & _
                                    cParam(1) & "," & _
                                    cParam(2) & "," & _
                                    cQuote & IIf(left(Replace(cParam(3), cQuote, ""), 1) > 1, "A", IIf(left(Replace(cParam(3), cQuote, ""), 1) = 0, "A", "Z")) & cQuote & "," & _
                                    cQuote & Format(Replace(cParam(4), cQuote, ""), "yyyymmdd") & cQuote & "," & _
                                    cQuote & Format(Replace(cParam(5), cQuote, ""), "hhmmss") & cQuote & ",0," & _
                                    cQuote & "0000000000" & cQuote & "," & _
                                    cQuote & "0" & cQuote & ",0)"
                        Set objdbRs = oTmpTxt.Execute(cSqlStmt)
                    End If
                End If
            Wend
        Else
           ' handle the error here!
           MsgBox Error(intError)
        End If
        .cfCloseFile    ' close the file. We're done with it!
    End With
    
    ShowProgress 4
    
    Set myTextFile = Nothing  ' Always set your objects to nothing when you're done with them!
    
    Exit Sub
    
ErrUpload:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub


Sub CreateFileTemp(oFileTempDB As ADODB.Connection)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String

        cSqlStmt = "create table Upload_Backup (" & _
                   " [u_fileName] char(50),     [u_date] date, " & _
                   " [u_time] char(50),         [u_UserId] char(6) )"
                   
    oFileTempDB.Execute cSqlStmt
    While oFileTempDB.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
'    ' in case table is already existing, let's clear it...
'    cSqlStmt = "DELETE FROM Upload_Backup"
'    Set objdbRs = oFileTempDB.Execute(cSqlStmt)
End Sub


Sub Bioprocess()
    On Error GoTo ErrCDlg
    
    Dim cSqlStmt As String, _
        oTmpTime As New ADODB.Connection, _
        oRecordSet As New ADODB.Recordset, _
        cLogDate As String, cTCID As String, _
        nTrnType As Integer, nRowPos As Integer, _
        aOtherInfo As Variant, _
        aShiftInfo As Variant
    Dim cFullName As String
    
    aOtherInfo = Array("", "", "", "")
    aShiftInfo = Array("", "", "", "", 0)
    
    With CommonDialog1
        .CancelError = True
    
        If lTimeKeeper Then
            .InitDir = CheckPath(gTimeKeeperPath)
            .Filter = "Time Keeper File |timekeeper.mdb"
        Else
            .InitDir = CheckPath(cUploadPath) & "Bioclock\"
            .Filter = "Bioclock File |*.txt"
        End If
    
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    '    MsgBox CommonDialog1.Filename
        
        ShowFileInfo IIf(Trim(gTimeKeeperPath) = "", "", .FileName)
    End With
    
    OpenQueryDNS "delete from `log` ", objdbRs, True

    With oTmpTime
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CheckPath(gTimeKeeperPath) & "timekeeper.MDB"
        .Open
    End With
    
    If Not lTimeKeeper Then
        
        CreateFileTemp oTmpTime
        cSqlStmt = "select * from Upload_Backup where u_fileName = " & cQuote & Trim(right(CommonDialog1.FileName, Len(CommonDialog1.FileName) - InStrRev(CommonDialog1.FileName, "\"))) & cQuote
        Set oRecordSet = oTmpTime.Execute(cSqlStmt)

        If oRecordSet.RecordCount > 0 Then

            OpenQueryDNS "SELECT concat(firstname,' ', left(mname,1),'. ', lastname) as fullname FROM pa2360 where userid = " & cQuote & oRecordSet("u_UserId") & cQuote, objdbRs, False
            cFullName = IIf(objdbRs.RecordCount > 0, objdbRs("fullname"), "")

            MsgBox " The Text File is already downloaded last " & vbCrLf & _
                    Format(oRecordSet("u_date"), "ddd mmm dd yyyy") & " - " & Trim(oRecordSet("u_time")) & vbCrLf & _
                   " By username " & cFullName, vbOKOnly + vbInformation, App.Title
        Else
            UploadTxtFile oTmpTime, CommonDialog1.FileName, gTimeKeeperPath
            
            cSqlStmt = "insert into Upload_Backup([u_fileName],[u_date],[u_time],[u_UserId])values(" & _
                        cQuote & right(CommonDialog1.FileName, Len(CommonDialog1.FileName) - InStrRev(CommonDialog1.FileName, "\")) & cQuote & "," & _
                        cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                        cQuote & Time & cQuote & "," & _
                        cQuote & gUserID & cQuote & ")"
            Set oRecordSet = oTmpTime.Execute(cSqlStmt)
        End If
    End If
    
    
    ' --> Save old DTR first here...
    Save2Backup oTmpTime

    ShowProgress 0
    cSqlStmt = "select machineno, accessidno, trandate, trantime, trantype, timecardid " & _
               "from timecard " & _
               "where tag=0 " & _
               "order by timecardid"
    Set oRecordSet = oTmpTime.Execute(cSqlStmt)
    If oRecordSet.RecordCount > 0 Then
        DoEvents
        While Not oRecordSet.EOF
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            cSqlStmt = "insert into `log`(bcid, tcid, dat1, dat2, transdate, trantime)values(" & _
                       cQuote & oRecordSet("machineno") & cQuote & "," & _
                       cQuote & oRecordSet("accessidno") & cQuote & "," & _
                       cQuote & PadStr(oRecordSet("timecardid"), "0", 6, PadLeft) & cQuote & "," & _
                       cQuote & IIf(oRecordSet("trantype") = "A", "0", "1") & cQuote & "," & _
                       cQuote & left(oRecordSet("trandate"), 4) & "-" & Mid(oRecordSet("trandate"), 5, 2) & "-" & right(oRecordSet("trandate"), 2) & cQuote & "," & _
                       cQuote & left(oRecordSet("trantime"), 2) & ":" & Mid(oRecordSet("trantime"), 3, 2) & ":" & right(oRecordSet("trantime"), 2) & cQuote & ")"
                       
            OpenQueryDNS cSqlStmt, objdbRs, True

            oRecordSet.MoveNext
        Wend
    End If
    ShowProgress 4
    
    ShowProgress 0
    
    cSqlStmt = " select a.bcid,a.tcid,ifnull(b.empid,'') as empid,ifnull(concat(b.lastname,', ',b.firstname),'') as fullname, ifnull(c.linename,'') as linename, " & _
               " ifnull(e.shiftid,'') as shiftid, ifnull(f.description,'') as shiftdesc, " & _
               "a.transdate as logdate, " & _
               " a.transdate,a.trantime, " & _
               " if(instr(a.dat2,'IN')>0 or instr(a.dat2,'0')>0,0,1) as trntype, " & _
               " if(instr(a.dat2,'IN')>0 or instr(a.dat2,'0')>0,'In','Out') as trndesc,'',0, " & _
               " ifnull(f.time1,'') as time1,ifnull(f.time2,'') as time2,ifnull(f.ndiff,0) as ndiff,0, " & _
               "a.dat1, ''," & _
               " ifnull(e.a_shiftid,'') as a_shiftid, " & _
               " ifnull(g.description,'') as a_shiftname, " & _
               " ifnull(g.time1,'') as a_time1,ifnull(g.time2,'') as a_time2 " & _
               " from ((log a left join di3670 b on a.tcid=b.tcid) left join di5463 c on b.depid=c.lineid) " & _
               " left join di546370 d on b.depid=d.depid and d.periodid=(select periodid from pa7730 where (a.transdate between date_start and date_end) and (13month=0)) " & _
               " left join di546373 e on d.sched_no=e.sched_no and e.date=a.transdate " & _
               " left join pa74380 f on e.shiftid=f.shiftid " & _
               " left join pa74380 g on e.a_shiftid=g.shiftid " & _
               "order by " & IIf(Not lTimeKeeper, "a.tcid, a.transdate, a.dat2 ", "a.dat1")
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        QueryAttach oRecordSet, MSHFlexGrid1, myArray, , , False, 1
        CheckGrid
    End If
    
    ShowProgress 4
 
    nAdd = 1
    CtrlPanel Me, nAdd
    
    MSHFlexGrid1.FixedCols = 3
    MSHFlexGrid1.LeftCol = 3
    MSHFlexGrid1.SetFocus

ErrCDlg:
    Set oTmpTime = Nothing
    Set oRecordSet = Nothing
End Sub

Private Sub Command1_Click()
    lProceed = False
    Frame1.Visible = True
'    Bioprocess
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
            nAdd = 0
            XPPanel1.Height = 825
            MSHFlexGrid1.FixedCols = 1
            lProceed = False
            
            CtrlPanel Me, nAdd
            
            ShowFileInfo ""
            
            SetGridColumn myArray, MSHFlexGrid1
        End If
    End If
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrBioProcessSave
    
    Dim oTmpTime As New ADODB.Connection, _
        nRowPos As Integer, _
        cSqlStmt As String, _
        cCmpname As String, _
        nTestTag As Integer
    
    Select Case MsgBox(IIf(nAdd = 1, "Download", "Cancel") & " Bio-Clock data file entry?", vbYesNoCancel, "Bio Clock Data File Entry...")
        Case vbYes
            If nAdd = 1 Then
            
'                If lTimeKeeper Then
                    With oTmpTime
                        .CursorLocation = adUseClient
                        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CheckPath(gTimeKeeperPath) & "timekeeper.MDB"
                        .Open
                    End With
'                End If
            
                ShowProgress 0
                
                With MSHFlexGrid1
                
                    DoEvents
                    
                    For nRowPos = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nRowPos / (.Rows - 1)) * 100, , , "Downloading... " & .TextMatrix(nRowPos, 3)
                        
                        .Row = nRowPos
                        
                         
                        If (.CellBackColor <> &HC0C0FF) And _
                           (.CellForeColor = vbBlack) Then
                           
                            cSeries = GenerateSeries("bio")
                            While IfExists("pa84650", "pa84650.tran_no=" & cQuote & PadStr(cSeries, "0", 10) & cQuote)
                                cSeries = GenerateSeries("bio")
                            Wend
                            cSeries = PadStr(cSeries, "0", 10)
    
                            cSqlStmt = "INSERT INTO pa84650(TRAN_NO,BCID,TCID,EMPID,SHIFTID,LOGDATE,TRANSDATE,TRANTIME,TRANTYPE,CMPID)VALUES(" & _
                                       cQuote & cSeries & cQuote & "," & _
                                       cQuote & PadStr(.TextMatrix(nRowPos, 1), "0", 2) & cQuote & "," & _
                                       cQuote & .TextMatrix(nRowPos, 2) & cQuote & "," & _
                                       cQuote & .TextMatrix(nRowPos, 3) & cQuote & "," & _
                                       cQuote & .TextMatrix(nRowPos, 6) & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nRowPos, 8), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nRowPos, 9), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nRowPos, 10), "HH:MM:SS") & cQuote & "," & _
                                       .TextMatrix(nRowPos, 11) & "," & _
                                       cQuote & gCompanyID & cQuote & ")"
'                            MsgBox cSqlStmt
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            
                            If .TextMatrix(nRowPos, 11) = 0 Then
                                cSqlStmt = " update di36770 set shiftid=" & cQuote & .TextMatrix(nRowPos, 6) & cQuote & _
                                           " , time1=" & cQuote & .TextMatrix(nRowPos, 15) & cQuote & _
                                           " , time2=" & cQuote & .TextMatrix(nRowPos, 16) & cQuote & _
                                           " , tag=0 " & _
                                           " Where empid = " & cQuote & .TextMatrix(nRowPos, 3) & cQuote & _
                                           " And Date = " & cQuote & Format(.TextMatrix(nRowPos, 9), "yyyy-mm-dd") & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, True
                            End If
                            
                            ' --> for security reason - 20070210
                            Log2Audit Name, "Upload bioclock entry of EmpID#" & .TextMatrix(nRowPos, 3) & " for " & Format(.TextMatrix(nRowPos, 8), "yyyy-mm-dd")
                            
                            
                            .TextMatrix(nRowPos, 18) = 1
                            
                            ' --> update timekeeper file...
'                            If lTimeKeeper Then
                                cSqlStmt = "update timecard set tag=1 where timecardid=" & Val(.TextMatrix(nRowPos, 19))
                                oTmpTime.Execute (cSqlStmt)
                                While oTmpTime.State = adStateExecuting
                                    DoEvents
                                Wend
'                            End If
                            
                            RebuildDTR 0, cQuote & .TextMatrix(nRowPos, 3) & cQuote
                            
                        End If

                    Next nRowPos
                    
                    ClearGrid MSHFlexGrid1, 18
                    
                End With
                
                ShowProgress 4
                
            End If
        
        Case vbNo
            
        Case vbCancel
            GoTo endsave
    End Select
    
    If Not ((MSHFlexGrid1.Rows - 1) > 1) Then
        nAdd = 0
        CtrlPanel Me, nAdd
    End If
    
endsave:
    Set oTmpTime = Nothing
    Exit Sub
    
ErrBioProcessSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command3_Click()
    Dim oRset1 As New ADODB.Recordset, _
        nCtr As Integer, _
        aOtherInfo As Variant, _
        aShiftInfo As Variant, _
        cSqlStmt As String
        
    aOtherInfo = Array("", "", "", "")
    aShiftInfo = Array("", "", "", "", 0)
        
    cSqlStmt = "select a.tcid, a.empid, concat(a.lastname,', ',a.firstname) as fullname, ifnull(b.linename,'') as linename, a.shiftid " & _
               " from di3670 a left join di5463 b on a.depid=b.lineid" & _
               " where a.tcid<>''"
    OpenQueryDNS cSqlStmt, oRset1, False
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        .Visible = False
        .Redraw = False
        
        DoEvents
        
        For nCtr = 1 To (.Rows - 1)
            
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
        '            myArray = Array("TXT:1[BCID]:5:True", _
        '                            "TXT:2[TCID]:6:True", _
        '                            "TXT:3[Emp Id]:8:True", _
        '                            "TXT:4[Name]:30:True", _
        '                            "TXT:5[Department]:20:True", _
        '                            "TXT:6[Shift ID]:5:False", _
        '                            "TXT:7[Shift]:20:True", _
        '                            "DAT:8[Log Date]:12:True", _
        '                            "DAT:9[Trans Date]:12:True", _
        '                            "TXT:0[Trans Time]:12:True", _
        '                            "TXT:1[Trantype]:1:False", _
        '                            "TXT:2[Type]:7:True", _
        '                            "TXT:3[Remark]:30:True", _
        '                            "NUM:4[tag]:1:false", _
        '                            "TXT:5[Time In]:10:False", _
        '                            "TXT:6[Time Out]:10:False", _
        '                            "NUM:7[ndiff]:1:False", _
        '                            "NUM:8[download]:1:False")
            If (Trim(.TextMatrix(nCtr, 3)) = "") Or (Trim(.TextMatrix(nCtr, 6)) = "") Then
                oRset1.Requery adAsyncFetch
                oRset1.Find "tcid='" & .TextMatrix(nCtr, 2) & "'"
                If Not oRset1.EOF Then
                    aOtherInfo(0) = oRset1("empid")
                    aOtherInfo(1) = DecodeStr(oRset1("fullname"))
                    aOtherInfo(2) = DecodeStr(oRset1("linename"))
                    aOtherInfo(3) = oRset1("shiftid")
                Else
                    aOtherInfo = Array("", "", "", "")
                End If
        
                If Trim(aOtherInfo(0)) <> "" Then
                    cSqlStmt = "select a.shiftid, ifnull(b.Description,'') as `description`, ifnull(b.ndiff,0) as ndiff, " & _
                               " ifnull(b.time1,'') as time1, ifnull(b.time2,'') as time2 " & _
                               "from di36770 a left join pa74380 b on a.shiftid=b.shiftid " & _
                               "where (a.empid=" & cQuote & aOtherInfo(0) & cQuote & ")" & _
                               " and (a.date=" & cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote & ")"
                    OpenQueryDNS cSqlStmt, objdbRs, False
                    If objdbRs.RecordCount = 0 Then
                        cSqlStmt = "select b.shiftid, b.Description, b.ndiff, b.time1, b.time2 " & _
                                   "from pa74380 b " & _
                                   "where b.shiftid=" & cQuote & aOtherInfo(3) & cQuote
                        OpenQueryDNS cSqlStmt, objdbRs, False
                    End If
                    
                    If objdbRs.RecordCount > 0 Then
                        aShiftInfo(0) = objdbRs("shiftid")
                        aShiftInfo(1) = DecodeStr(objdbRs("description"))
                        aShiftInfo(2) = Format(objdbRs("time1"), "hh:mm:ss")
                        aShiftInfo(3) = Format(objdbRs("time2"), "hh:mm:ss")
                        aShiftInfo(4) = objdbRs("ndiff")
                    Else
                        aShiftInfo = Array("", "", "", "", 0)
                    End If
                Else
                    aShiftInfo = Array("", "", "", "", 0)
                End If
            
                .TextMatrix(nCtr, 3) = aOtherInfo(0)
                .TextMatrix(nCtr, 4) = aOtherInfo(1)
                .TextMatrix(nCtr, 5) = aOtherInfo(2)
                .TextMatrix(nCtr, 6) = aShiftInfo(0)
                .TextMatrix(nCtr, 7) = aShiftInfo(1)
                .TextMatrix(nCtr, 15) = aShiftInfo(2)
                .TextMatrix(nCtr, 16) = aShiftInfo(3)
                .TextMatrix(nCtr, 17) = aShiftInfo(4)
            
                CheckGrid True, nCtr
            End If
        Next nCtr
        
        .Redraw = True
        
        .Visible = True
        
        ShowProgress 4
        
    End With

    Set oRset1 = Nothing
End Sub

Private Sub Command7_Click(Index As Integer)
    Frame1.Visible = False
    If (Index < 2) Then
        lTimeKeeper = Index = 0
        Bioprocess
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[BCID]:5:True", _
                    "TXT:[TCID]:6:True", _
                    "TXT:[Emp Id]:8:True", _
                    "TXT:[Name]:25:True", _
                    "TXT:[Department]:20:True", _
                    "TXT:[Shift ID]:5:False", _
                    "TXT:[Shift]:20:True", _
                    "DAT:[Log Date]:12:True", _
                    "DAT:[Trans Date]:12:True", _
                    "TXT:[Trans Time]:12:True", _
                    "TXT:[Trantype]:1:False", _
                    "TXT:[Type]:6:True", _
                    "TXT:[Remark]:30:True", _
                    "NUM:[tag]:1:False", _
                    "TXT:[Start Time]:10:True", _
                    "TXT:[End Time]:10:True", _
                    "NUM:[ndiff]:1:False", _
                    "NUM:[download]:1:False", _
                    "TXT:[tag]:10:False", _
                    "TXT:[time card id]:10:False", _
                    "TXT:[a Shift ID]:5:False", _
                    "TXT:[a Shift]:20:False", _
                    "TXT:[a Time In]:10:False", _
                    "TXT:[a Time Out]:10:False")

    Tag = nAccess_Tag
    nAdd = 0
    Combo1.ListIndex = 0
    XPPanel1.Height = 825
    
    CtrlPanel Me, nAdd
    
    ShowFileInfo ""
    
    SetGridColumn myArray, MSHFlexGrid1
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
    Dim cDepid As String
    
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid1
        
        Select Case KeyCode
'            Case vbKeyDown
'                If .Row = .Rows - 1 Then
'                    If (Trim(.TextMatrix(.Rows - 1, 2)) <> "") Then
'                        .AddItem "", .Rows
'                        .RowHeight(.RowSel + 1) = 285
'                        .Row = .RowSel + 1
'                        .TopRow = .Row
'
'                        .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)
'
'                        ' --> added 20051115
'                        HiLyt .Row, Int(.Row / 2) = .Row / 2, MSHFlexGrid1, , IIf(Int(.Row / 2) = .Row / 2, &HE0E0E0, vbWhite)
'
'                        .FixedCols = 1
'                        .LeftCol = 1
'                        .Col = 1
'                        .ColSel = 1
'                    End If
'                End If
'
'            Case vbKeyUp
'                If .Rows - 1 > 1 Then
'                    If Trim(.TextMatrix(.Rows - 1, 2)) = "" Then
'                        .Rows = .Rows - 1
'
'                        .FixedCols = 3
'                        .LeftCol = 3
'                        .Col = 3
'                        .ColSel = 3
'                    End If
'                End If
'
'            Case vbKeyRight
'                If .Col > 2 Then
'                    .FixedCols = 3
'                End If
'
'            Case vbKeyLeft
'                If .ColSel < 1 Then
'                    .LeftCol = 1
'                    .Col = 1
'                    .ColSel = 1
'                End If
'
'            Case vbKeyInsert    ' --> 20050908
'                If .TextMatrix(.RowSel, 1) <> "" Then
'                    .AddItem "", .RowSel
'                    .RowHeight(.RowSel) = 285
'
'                    RefreshGrid MSHFlexGrid1, True
'
'                    '.Row = .RowSel + 1
'                    .SetFocus
'                End If
                
            Case vbKeyReturn
                Select Case .ColSel
                    Case 3
                        If Trim(.TextMatrix(.Row, 3)) <> "" Then Exit Sub
                        
                        frmLookup.showPopup 2
                        frmLookup.Show 1
                        If Trim(cResult) <> "" Then
                            OpenQueryDNS "SELECT * FROM DI5463 WHERE LINEID=" & cQuote & cResult & cQuote, objdbRs, False
                            If objdbRs.RecordCount > 0 Then
                                cDepid = cResult
                                frmLookup.showPopup 3, "where a.depid = " & cQuote & cDepid & cQuote
                                frmLookup.Show 1
                                If Trim(cResult) <> "" Then
                                    OpenQueryDNS " select a.empid, ifnull(concat(a.lastname, ', ', a.firstname , ' ', a.mname),'') as fullname, " & _
                                                 " ifnull(b.linename,'') as linename from di3670 a " & _
                                                 " left join di5463 b on a.depid=b.lineid " & _
                                                 " Where a.empid = " & cQuote & cResult & cQuote, objdbRs, False
                                    If objdbRs.RecordCount > 0 Then
                                        .TextMatrix(.Row, 3) = cResult
                                        .TextMatrix(.Row, 4) = objdbRs("fullname")
                                        .TextMatrix(.Row, 5) = objdbRs("linename")
                                        CheckGrid True, .Row
                                    End If
                                End If
                            End If
                        End If
                        .SetFocus
                        
                    Case 8
                        Command11.Cancel = False
                        dtFlex.Visible = True
                        dtFlex.left = .CellLeft + .left - (dtFlex.Width - .CellWidth)
                        dtFlex.top = .CellTop + .top - 10
                        dtFlex.Value = IIf(Trim(.Text) = "", Now, .Text)
                        dtFlex.SetFocus
                    
                    Case 12
                        Command11.Cancel = False
                        cmbFlex.ZOrder 0
                        cmbFlex.Visible = True
                        cmbFlex.left = .CellLeft + .left - (cmbFlex.Width - .CellWidth)
                        cmbFlex.top = .CellTop + .top - 10
                        cmbFlex.ListIndex = IIf(Trim(.Text) = "", 0, Val(.TextMatrix(.Row, 11)))
                        cmbFlex.SetFocus
                        
                    Case 13
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
                
            
            Case vbKeySpace     ' --> marked/unmarked entry
                If .RowSel < .Rows Then
                    If Not lProceed Then
                        lProceed = MsgBox("You are about to marked/unmarked an entry as deleted." & vbCrLf & _
                                          "Marked entry are denoted by a light red highlight and" & vbCrLf & _
                                          "can be unmarked by pressing the [SPACEBAR] or" & vbCrLf & _
                                          "by [CTRL-LEFT CLICK] again." & vbCrLf & vbCrLf & _
                                          "Press [Yes] if you do not wish to continue receiving this message...", vbYesNo, App.Title) = vbYes
                    End If
                    HiLyt .Row, .CellBackColor <> &HC0C0FF, MSHFlexGrid1, , &HC0C0FF
                    .SetFocus
                End If
                
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then
        KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex") And _
                     (Screen.ActiveForm.ActiveControl.Name <> "cmbFlex") And _
                     (Screen.ActiveForm.ActiveControl.Name <> "dtFlex")
    End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If nAdd <> 0 Then
        If (Button = vbLeftButton) And ((Shift And vbCtrlMask) > 0) Then
            If Not lProceed Then
                lProceed = MsgBox("You are about to marked/unmarked an entry as deleted." & vbCrLf & _
                                  "Marked entry are denoted by a light red highlight and" & vbCrLf & _
                                  "can be unmarked by pressing the [SPACEBAR] or" & vbCrLf & _
                                  "by [CTRL-LEFT CLICK] again." & vbCrLf & vbCrLf & _
                                  "Press [Yes] if you do not wish to continue receiving this message...", vbYesNo, App.Title) = vbYes
            End If
            HiLyt MSHFlexGrid1.RowSel, MSHFlexGrid1.CellBackColor <> &HC0C0FF, MSHFlexGrid1, , IIf(MSHFlexGrid1.CellBackColor <> &HC0C0FF, &HC0C0FF, &H80000005)
        End If
    End If
End Sub

Private Sub cmbFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 11) = cmbFlex.ListIndex
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 12) = cmbFlex.Text
            cmbFlex_LostFocus
            MSHFlexGrid1.SetFocus
            
        Case vbKeyEscape
            cmbFlex_LostFocus
            MSHFlexGrid1.SetFocus
            
    End Select
End Sub

Private Sub cmbFlex_LostFocus()
    cmbFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub dtFlex_DblClick()
    dtFlex_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtFlex_LostFocus
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8) = Format(dtFlex.Value, "mm/dd/yyyy")
        MSHFlexGrid1.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtFlex_LostFocus
        MSHFlexGrid1.SetFocus
    End If
End Sub

Private Sub dtFlex_LostFocus()
    dtFlex.Visible = False
    Command11.Cancel = True
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid1
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 13) = txtFlex.Text
                MSHFlexGrid1_KeyDown vbKeyDown, 0
                
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

Private Sub XPButton1_Click()
    XPPanel1.Height = IIf(XPPanel1.Height = 825, 2385, 825)
    XPButton1.Caption = IIf(XPPanel1.Height = 825, "S&how Detail >>", "<< &Hide Detail")
End Sub
