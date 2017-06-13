VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{30DA1A2F-A970-4238-AC17-5773BA9DC841}#1.1#0"; "CIAXPDatePicker.ocx"
Begin VB.Form frmCheckSched 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shifting Schdule Checking"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12840
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   1890
      TabIndex        =   11
      Top             =   7275
      Width           =   10890
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4560
         Picture         =   "frmCheckSched.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   8955
         Picture         =   "frmCheckSched.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7995
         Picture         =   "frmCheckSched.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   7155
         Picture         =   "frmCheckSched.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   6315
         Picture         =   "frmCheckSched.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2640
         Picture         =   "frmCheckSched.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1800
         Picture         =   "frmCheckSched.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   960
         Picture         =   "frmCheckSched.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   9930
         Picture         =   "frmCheckSched.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   5475
         Picture         =   "frmCheckSched.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3630
         Picture         =   "frmCheckSched.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   4
         Left            =   120
         Picture         =   "frmCheckSched.frx":11896
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Data"
      Height          =   450
      Left            =   480
      TabIndex        =   10
      Top             =   1005
      Width           =   1575
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
      Left            =   1245
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:SSCHECKID"
      Top             =   30
      Width           =   1050
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
      Left            =   1245
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:PERIODID"
      Top             =   645
      Width           =   585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   1845
      TabIndex        =   1
      Top             =   645
      Width           =   375
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6060
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   3030
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1245
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "DAT:SSCHECKDATE"
      Top             =   330
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16580608
      CurrentDate     =   38623
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5790
      Left            =   45
      TabIndex        =   5
      Top             =   1485
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10213
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
   Begin ciaXPDatePicker.XPDatePicker XPDatePicker1 
      Height          =   315
      Left            =   4755
      TabIndex        =   24
      Top             =   585
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      FormatString    =   "dddd - MMM dd, yyyy"
      MouseIcon       =   "frmCheckSched.frx":13218
      CalendarDayBorder=   -1  'True
      CalendarDayBorderColor=   -2147483646
      CalendarMonthBorderColor=   8421504
      LicValid        =   -1  'True
   End
   Begin ciaXPDatePicker.XPDatePicker XPDatePicker2 
      Height          =   315
      Left            =   7305
      TabIndex        =   25
      Top             =   585
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      FormatString    =   "dddd - MMM dd, yyyy"
      MouseIcon       =   "frmCheckSched.frx":13234
      CalendarDayBorder=   -1  'True
      CalendarDayBorderColor=   -2147483646
      CalendarMonthBorderColor=   8421504
      LicValid        =   -1  'True
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
      Left            =   120
      TabIndex        =   9
      Top             =   90
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
      Left            =   120
      TabIndex        =   8
      Top             =   390
      Width           =   1455
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
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   705
      Width           =   4005
   End
End
Attribute VB_Name = "frmCheckSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmCheckSched
' description   :   Shifting Schedule Checking (SL/VL/FL) Module
' programmer    :   _-=[ srm ]=-_
' date          :   09 may 2008

Option Explicit
    Dim nAdd As Integer, _
        nLastRow As Integer, _
        cSeries As String, _
        cParam As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant
        
Function ChkHoliday(dDate As Date) As String
    Dim cSqlStmt As String
    cSqlStmt = "select a.description from pa4329 a" & _
               " where (a.date=" & cQuote & Format(dDate, "yyyy-mm-dd") & cQuote & ") or" & _
               " (date_format(a.date,'%m %d')=" & cQuote & Format(dDate, "mmm dd") & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then ChkHoliday = objdbRs("description")
End Function
        
        
Sub IsEqual(ByVal nRow As Integer)
    With MSHFlexGrid1
        If Val(.TextMatrix(nRow, 17)) = 1 Then
           HiLyt nRow, True, MSHFlexGrid1, , vbInfoBackground
        Else
           HiLyt nRow, False, MSHFlexGrid1, , vbInfoBackground
        End If
    End With
End Sub
        
Function shiftcheck(ByVal cShiftid As String, ByVal nCtr As Integer, ByVal cLogDate As Date, ByVal cTransdate As Date, Optional ByVal cTimein As String = "") As String
    Dim cTimeComp As String
    
    If cShiftid <> "" Then
    
        OpenQueryDNS " select SHIFTID, DESCRIPTION, TIME1, TIME2, allowance from PA74380 " & _
                   " Where shiftid = " & cQuote & cShiftid & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            
            cTimeComp = Format(Hour(objdbRs("time1")) & ":" & Minute(objdbRs("time1")) + objdbRs("allowance"), "hh:mm:ss")
            
            Select Case Format(cTimein, "hh:mm:ss")
                ' pag less than sya
                Case Is <= Format(cTimeComp, "hh:mm:ss")
                    If Format(cTimein, "hh:mm:ss") <= Format(Hour(cTimeComp) - 1 & ":" & Minute(cTimeComp) & ":" & Second(cTimeComp), "hh:mm:ss") Then
                        shiftcheck = "Wrong Shifting"
                    Else
                        shiftcheck = "Correct Shifting"
                    End If
                
                ' pag greater than sya
                Case Is >= Format(cTimeComp, "hh:mm:ss")
                
                    If Format(cTimein, "hh:mm:ss") >= Format(cTimeComp, "hh:mm:ss") Then
                        shiftcheck = "Wrong Shifting"
                    Else
                        shiftcheck = "Correct Shifting"
                    End If
            End Select
        End If
    Else
        shiftcheck = ""
    End If
        
        
End Function

Sub InsertToGrid(nRowPos As Integer, ByVal cString As String)
    Dim lOverWrite As Boolean
    
    If Trim(cString) <> "" Then
    
        With MSHFlexGrid1
            If Trim(.TextMatrix(nRowPos, 13)) = "" Then
                lOverWrite = True
            Else
                lOverWrite = (MsgBox("Warning!!!" & vbCrLf & "Do you wish to overwrite existing item?", vbYesNo, App.Title) = vbYes)
            End If
            
            If lOverWrite Then
                .TextMatrix(nRowPos, 7) = cString
                
                OpenQueryDNS "SELECT shiftid,ifnull(description,'')as description, " & _
                             " ifnull(time1,'')as time1,ifnull(time2,'')as time2 FROM pa74380 " & _
                             " where shiftid = " & cQuote & cString & cQuote, objdbRs, False
                .TextMatrix(nRowPos, 10) = DecodeStr(objdbRs("description"))
                .TextMatrix(nRowPos, 17) = "1"

                IsEqual nRowPos
                
                .Row = nRowPos
            End If
        End With
    End If
End Sub
        
Sub TranFillGRid()
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer
    
    With MSHFlexGrid1
          
            ShowProgress 0
            .Redraw = False
            
            For nCtr = 1 To (.Rows - 1)
            
                ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
                cSqlStmt = " select a.TRAN_NO , a.empid, ifnull(a.shiftid,ifnull(c.shiftid,'')) as shiftid, a.LOGDATE, a.transdate, ifnull(b.Description,'') as Description, a.TRANTIME, a.trantype " & _
                           " from pa84650 a " & _
                           " left join pa74380 b on a.shiftid=b.shiftid " & _
                           " left join di36770 c on a.empid=c.empid and a.transdate =c.date " & _
                           " Where a.transdate = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                           " And a.trantype = 0 And a.empid = " & cQuote & .TextMatrix(nCtr, 2) & cQuote & _
                           " group by a.empid "

'                MsgBox cSqlStmt
'                Script2File cSqlStmt
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
'                    If objdbRs("empid") = "039117" Then MsgBox "stop"
                
                    .TextMatrix(nCtr, 1) = objdbRs("tran_no")
                    .TextMatrix(nCtr, 7) = objdbRs("shiftid")
                    .TextMatrix(nCtr, 8) = Format(objdbRs("logdate"), "yyyy-mm-dd")
                    .TextMatrix(nCtr, 9) = Format(objdbRs("transdate"), "yyyy-mm-dd")
                    .TextMatrix(nCtr, 10) = objdbRs("description")
                    .TextMatrix(nCtr, 11) = objdbRs("trantime")
                    .TextMatrix(nCtr, 12) = objdbRs("trantype")
                    
                    cString = shiftcheck(.TextMatrix(nCtr, 7), nCtr, .TextMatrix(nCtr, 8), .TextMatrix(nCtr, 9), .TextMatrix(nCtr, 11))
                       
                    If cString <> "" Then
                        If cString = "Correct Shifting" Then
                            HiLyt2 nCtr, MSHFlexGrid1, vbBlack
                            .TextMatrix(nCtr, 16) = 1
                        Else
                            HiLyt2 nCtr, MSHFlexGrid1, vbBlue
                            .TextMatrix(nCtr, 16) = 0
                        End If
                    Else
                        cString = "Undefined Shifting"
                        HiLyt2 nCtr, MSHFlexGrid1, vbRed
                        .TextMatrix(nCtr, 16) = 2
                    End If
                    
                    .TextMatrix(nCtr, 13) = cString
                Else
                    .TextMatrix(nCtr, 16) = 1
                End If

            Next nCtr
            
            .TopRow = 1
            .Redraw = True
           
            ShowProgress 4
            
    End With
End Sub
        
        
Sub FillGrid()
        Dim nCtr As Integer
   
        If nAdd > 0 Then
            
            With MSHFlexGrid1
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
                    If Val(.TextMatrix(nCtr, 11)) = 0 Then
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
            
        End If

End Sub

Sub ShowData(ByVal cString As String, oLabel As Label, nMode As Integer)
    Dim cSqlStmt As String
    If nMode = 1 Then
        cSqlStmt = "select duration,date_start,date_end from pa7730 where periodid=" & cQuote & Text2.Text & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        Label4.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("duration"), "")
        XPDatePicker1.CurrentDate = IIf(objdbRs.RecordCount > 0, objdbRs("date_start"), Now)
        XPDatePicker2.CurrentDate = IIf(objdbRs.RecordCount > 0, objdbRs("date_end"), Now)
    End If
End Sub
        
        
Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1
    
    ShowData Text2.Text, Label4, 1
    
    Command2.Enabled = nAdd <> 0
    Command3.Enabled = nAdd <> 0
    
    cSqlStmt = "SELECT a.TRAN_NO, a.EMPID, concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'.') as fullname, " & _
               " b.depid,c.linename, d.posname , a.shiftid, a.LOGDATE, a.TRANSDATE,ifnull(e.description,'') as desc1, " & _
               " a.TRANTIME, a.TRANTYPE, a.REMARK , a.SEQ_NO, a.Status, a.Tag, a.sTag " & _
               " FROM pa7723 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " left join di7670 d on b.posid=d.posid " & _
               " left join PA74380 e on a.shiftid=e.shiftid " & _
               " where a.sscheckid = " & cQuote & Text1.Text & cQuote & _
               " order by a.seq_no "
               
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
    On Error GoTo ErrSSCHECK
    
    Dim cString As String, _
        cSqlStmt As String, _
        nCtr As Integer
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Shifting Schedule Checking entry?", vbYesNoCancel, "Shifting Schedule Checking Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("pa7720", "SSCHECKID=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Shifting Schedule Checking Reference Number already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "pa7720"), oTempADO, True
                    Script2File InsertFields(Me, "pa7720")
                    
                    Log2Audit Name, "ADD SSCHECKID -->" & Trim(Text1.Text)
                
                    ShowProgress 0
                    
                    With MSHFlexGrid1
                        For nCtr = 1 To .Rows - 1
                        
                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                            
                            If Trim(.TextMatrix(nCtr, 17)) = 1 Then
                            
                                cSqlStmt = "insert into pa7723(SSCHECKID,SSCHECKDATE,TRAN_NO,EMPID,LOGDATE,TRANSDATE,SHIFTID,TRANTIME, " & _
                                           " TRANTYPE,SEQ_NO,REMARK,status,tag,stag)values(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                           cQuote & Format(.TextMatrix(nCtr, 8), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 12) & cQuote & "," & _
                                           nCtr & "," & _
                                           cQuote & .TextMatrix(nCtr, 13) & cQuote & "," & _
                                           Val(.TextMatrix(nCtr, 15)) & "," & _
                                           Val(.TextMatrix(nCtr, 16)) & "," & _
                                           Val(.TextMatrix(nCtr, 17)) & ")"
'                                MsgBox cSqlStmt
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                            
                        Next nCtr
                    End With
                    
                    ShowProgress 4
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "pa7720", "SSCHECKID=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "pa7720", "SSCHECKID=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT SSCHECKID -->" & Trim(Text1.Text)
            
                cSqlStmt = "delete from pa7723 where SSCHECKID=" & cQuote & Text1.Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 0
                
                With MSHFlexGrid1
                    For nCtr = 1 To .Rows - 1
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 17)) = 1 Then
                        
                            cSqlStmt = "insert into pa7723(SSCHECKID,SSCHECKDATE,TRAN_NO,EMPID,LOGDATE,TRANSDATE,SHIFTID,TRANTIME, " & _
                                       " TRANTYPE,SEQ_NO,REMARK,status,tag,stag)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 8), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 12) & cQuote & "," & _
                                       nCtr & "," & _
                                       cQuote & .TextMatrix(nCtr, 13) & cQuote & "," & _
                                       Val(.TextMatrix(nCtr, 15)) & "," & _
                                       Val(.TextMatrix(nCtr, 16)) & "," & _
                                       Val(.TextMatrix(nCtr, 17)) & ")"
'                                MsgBox cSqlStmt
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
    
    If Text1.Text <> cSeries Then ResetSeries "SSCHECK", cSeries
    
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "SSCHECKID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO

    ShowRecords
endsave:
    Exit Sub
    
ErrSSCHECK:
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
            
            If Text1.Text <> cSeries Then ResetSeries "SSCHECK", cSeries
            
            nAdd = 0
            
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "SSCHECKID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            ShowRecords
        End If
    End If
End Sub

Private Sub Command2_Click()
    frmLookup.showPopup 5, " where pclose=0 "
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text2.Text = cResult
        ShowData Text2.Text, Label4, 1
    End If
End Sub
Private Sub Command3_Click()
    Dim cSqlStmt As String
    
    If Text2.Text <> "" Then
    
        cSqlStmt = " select '',a.EMPID, concat(a.LASTNAME,', ',a.FIRSTNAME,' ', left(a.MNAME,1),'.') as fullname," & _
                   " a.DEPID,ifnull(b.linename,'') as linename,ifnull(c.posname,'') as posname " & _
                   " From di3670 a " & _
                   " left join di5463 b on a.depid=b.lineid " & _
                   " left join di7670 c on a.posid=c.posid " & _
                   " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ")))) or " & _
                   "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))))" & _
                   " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))" & _
                   " group by a.empid " & _
                   " order by b.linename,a.lastname "
'        MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            QueryAttach objdbRs, MSHFlexGrid1, myArray, False, , , 1
            TranFillGRid
            FillGrid
        Else
            SetGridColumn myArray, MSHFlexGrid1
            MsgBox "No data detected", vbInformation
        End If
    Else
        MsgBox "Specify period information! ", vbInformation
        Text2.SetFocus
    End If
End Sub

Private Sub Command4_Click()
    On Error GoTo ErrSSCHECK
    
    Dim oRecordSet As New ADODB.Recordset, _
        lProceed As Boolean, _
        nCtr As Integer, _
        nCount As Integer, _
        cSqlStmt As String, _
        cString As String, _
        aTimeInfo As Variant, _
        cTime1, cTime2 As String
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If

    If lProceed Then
        If MsgBox("Apply this Shift Schedule Checking entry?", vbYesNo, App.Title) = vbYes Then
        
            cString = Text1.Text
            
            ShowProgress 0
            
            With MSHFlexGrid1
                For nCtr = 1 To .Rows - 1
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    
                    If Val(.TextMatrix(nCtr, 17)) <> 0 Then
                    
                        cSqlStmt = "update pa7723 set status=1, " & _
                                   " date_post=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                                   " where SSCheckID=" & cQuote & Text1.Text & cQuote & _
                                   " and seq_no=" & Val(.TextMatrix(nCtr, 14))
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt

                    ' update shiftid pa84650
                        cSqlStmt = "update pa84650 set shiftid=" & cQuote & .TextMatrix(nCtr, 7) & cQuote & _
                                   " where tran_no = " & cQuote & .TextMatrix(nCtr, 1) & cQuote & _
                                   " and empid = " & cQuote & .TextMatrix(nCtr, 2) & cQuote & _
                                   " and logdate =" & cQuote & Format(.TextMatrix(nCtr, 8), "yyyy-mm-dd") & cQuote & _
                                   " and transdate =" & cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote
'                        MsgBox cSqlStmt
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                        
                        
                   ' --> added 20161109(1)
                   ' update time out shiftid pa84650
                        cSqlStmt = "update pa84650 set shiftid=" & cQuote & .TextMatrix(nCtr, 7) & cQuote & _
                                   " where trantype = " & cQuote & 1 & cQuote & _
                                   " and empid = " & cQuote & .TextMatrix(nCtr, 2) & cQuote & _
                                   " and logdate =" & cQuote & Format(.TextMatrix(nCtr, 8), "yyyy-mm-dd") & cQuote & _
                                   " and transdate =" & cQuote & Format(.TextMatrix(nCtr, 9), "yyyy-mm-dd") & cQuote
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt

                        OpenQueryDNS " SELECT shiftid,description,time1,time2 FROM PA74380 " & _
                                     " where shiftid = " & cQuote & .TextMatrix(nCtr, 7) & cQuote, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                            cTime1 = objdbRs("time1")
                            cTime2 = objdbRs("time2")
                        Else
                            cTime1 = ""
                            cTime2 = ""
                        End If

                        cSqlStmt = " select empid,emp_stat,active,paystatus from di3670 " & _
                                   " where empid = " & cQuote & .TextMatrix(nCtr, 2) & cQuote
                        OpenQueryDNS cSqlStmt, oRecordSet, False
                        If oRecordSet.RecordCount > 0 Then
                            ' --> retrieve computed dtr here... 20060907
                            aTimeInfo = ComputeDays(.TextMatrix(nCtr, 2), _
                                                     Array(Format(DTPicker1.Value, "yyyy-mm-dd"), Format(DTPicker1.Value, "yyyy-mm-dd"), 0), _
                                                    Array(oRecordSet("emp_stat"), oRecordSet("active"), oRecordSet("paystatus")))
                            If IfExists("di36770", "(empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & ") and (di36770.date=" & cQuote & Format(DTPicker1, "yyyy-mm-dd") & cQuote & ")") Then
'                                MsgBox "d2"
                                cSqlStmt = "update di36770 set " & _
                                           "  shiftid=" & cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & _
                                           "  time1=" & cQuote & Format(cTime1, "HH:MM:SS") & cQuote & "," & _
                                           "  time2=" & cQuote & Format(cTime2, "HH:MM:SS") & cQuote & "," & _
                                           "  reg_hr=" & aTimeInfo(0) * 8 & "," & _
                                           "  reg_ot_hr=" & aTimeInfo(1) & "," & _
                                           "  sa_reg_ot=" & aTimeInfo(2) & "," & _
                                           "  tot_ot=" & aTimeInfo(1) + aTimeInfo(2) & "," & _
                                           "  nd_hr=" & aTimeInfo(3) * 8 & "," & _
                                           "  nd_ot_hr=" & aTimeInfo(4) & "," & _
                                           "  sa_nd_ot=" & aTimeInfo(12) & "," & _
                                           "  nd_tot_ot=" & aTimeInfo(4) + aTimeInfo(12) & "," & _
                                           "  sun_hr=" & aTimeInfo(5) & "," & _
                                           "  sun_ot_hr=" & aTimeInfo(6) & "," & _
                                           "  sun_nd=" & aTimeInfo(13) & "," & _
                                           "  sun_nd_ot=" & aTimeInfo(14) & "," & _
                                           "  remark=" & cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(.TextMatrix(nCtr, 7)) <> ""), "No Entry or Absent", IIf(aTimeInfo(11) = 2, "On Leave", ChkHoliday(Format(DTPicker1.Value, "yyyy-mm-dd")))), ChkHoliday(Format(DTPicker1.Value, "yyyy-mm-dd")))) & cQuote & "," & _
                                           "  tag=" & IIf(aTimeInfo(10) > 0, 3, IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(.TextMatrix(nCtr, 7)) <> ""), 1, IIf(aTimeInfo(11) = 2, 2, 0)), 0)) & _
                                           " where (empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & ") " & _
                                           " and (di36770.date=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")"
    
                            Else
'                                MsgBox "d d2"
                                cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,time1,time2,`remark`,`tag`," & _
                                           "reg_hr,reg_ot_hr,sa_reg_ot,tot_ot,nd_hr,nd_ot_hr,sa_nd_ot,nd_tot_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot)values(" & _
                                           cQuote & MSHFlexGrid1.TextMatrix(nCtr, 2) & cQuote & "," & _
                                           cQuote & Text2.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & _
                                           cQuote & Format(cTime1, "HH:MM:SS") & cQuote & "," & _
                                           cQuote & Format(cTime2, "HH:MM:SS") & cQuote & "," & _
                                           cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(.TextMatrix(nCtr, 7)) <> ""), "No Entry or Absent", IIf(aTimeInfo(11) = 2, "On Leave", ChkHoliday(Format(DTPicker1.Value, "yyyy-mm-dd")))), ChkHoliday(Format(DTPicker1.Value, "yyyy-mm-dd")))) & cQuote & "," & _
                                           IIf(aTimeInfo(10) > 0, 3, IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(.TextMatrix(nCtr, 7)) <> ""), 1, IIf(aTimeInfo(11) = 2, 2, 0)), 0)) & "," & _
                                           aTimeInfo(0) * 8 & "," & _
                                           aTimeInfo(1) & "," & _
                                           aTimeInfo(2) & "," & _
                                           aTimeInfo(1) + aTimeInfo(2) & "," & _
                                           aTimeInfo(3) * 8 & "," & _
                                           aTimeInfo(4) & "," & _
                                           aTimeInfo(12) & "," & _
                                           aTimeInfo(4) + aTimeInfo(12) & "," & _
                                           aTimeInfo(5) & "," & _
                                           aTimeInfo(6) & "," & _
                                           aTimeInfo(13) & "," & _
                                           aTimeInfo(14) & ")"
                            End If
'                            MsgBox cSqlStmt
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                        End If

                    End If
                        
                    nCount = nCount + 1
                    
                Next nCtr
                
                If nCount = .Rows - 1 Then
                    cSqlStmt = "update pa7720 set status=1," & _
                               " date_post=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                               " where SSCheckID=" & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
                
                ShowProgress 4
                
                oTempADO.Requery adAsyncFetch
                If Trim(cString) <> "" Then oTempADO.Find "SSCheckID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
                GetFields Me, oTempADO
            
                ShowRecords
                
            End With
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this Shift Schedule Checking entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Exit Sub
        
    Set oRecordSet = Nothing
        
ErrSSCHECK:
    ErrorMsg Err.Number, Err.Description, "SSCHECK Incentive #" & Text1.Text, Name
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 17
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "SSCheckID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmp7720( " & _
               " [SSCheckID] char(10),       [SSCheckDATE] date," & _
               " [DURATION] char(100),       [TRAN_NO] char(10)," & _
               " [EMPID] char(6),            [FULLNAME] char(100)," & _
               " [DEPID] char(3),            [LINENAME] char(100)," & _
               " [POSNAME] char(100),        [LOGDATE] date," & _
               " [TRANSDATE] date,           [SHIFTID] char(5)," & _
               " [DESCRIPTION] char(100),    [TRANTIME] char(10)," & _
               " [TRANTYPE] char(10),        [SEQ_NO] integer, " & _
               " [REMARK] char(50),          [status] integer, " & _
               " [date_post] date,           [TAG] integer, " & _
               " [CMPName] char(50))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmp7720", oTempADO, True
End Sub


Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer

    CreateTemp
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            cSqlStmt = " insert into tmp7720(SSCheckID,SSCheckDATE,DURATION,TRAN_NO,EMPID,FULLNAME,DEPID,LINENAME,POSNAME,LOGDATE,TRANSDATE, " & _
                       " SHIFTID,DESCRIPTION,TRANTIME,TRANTYPE,SEQ_NO,REMARK,status)values(" & _
                       cQuote & Text1.Text & cQuote & "," & _
                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Label4.Caption & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & cQuote & .TextMatrix(nCtr, 4) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 5) & cQuote & "," & cQuote & .TextMatrix(nCtr, 6) & cQuote & "," & _
                       cQuote & Format(.TextMatrix(nCtr, 8), "mm-dd-yyyy") & cQuote & "," & _
                       cQuote & Format(.TextMatrix(nCtr, 9), "mm-dd-yyyy") & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 7) & cQuote & "," & cQuote & .TextMatrix(nCtr, 10) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 11) & cQuote & "," & cQuote & .TextMatrix(nCtr, 12) & cQuote & "," & _
                       nCtr & "," & _
                       cQuote & .TextMatrix(nCtr, 13) & cQuote & "," & _
                       cQuote & .TextMatrix(nCtr, 15) & cQuote & ")"
                       
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
        Next nCtr
        
        ShowProgress 3

        GenerateReport "Shifting Schedule Chekking Report", "PRV7720.RPT"

        ShowProgress 4
        
    End With

End Sub

Private Sub Command7_Click()
    SetGridColumn myArray, MSHFlexGrid1
    
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    XPDatePicker1.CurrentDate = Now
    XPDatePicker2.CurrentDate = Now
    
    Command2.Enabled = True
    Command3.Enabled = True
    
    Label4.Caption = ""
    
    cSeries = GenerateSeries("SSCHECK")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("pa7720", "pa7720.SSCheckId=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("SSCHECK")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text1.SetFocus
End Sub

Private Sub Command8_Click()
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, True
        
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Command2.Enabled = True
        Command3.Enabled = True
        
        Text1.Enabled = False
        DTPicker1.SetFocus
        FillGrid

    End If

End Sub

Private Sub Command9_Click()
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "delete from pa7720 where sscheckid=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from pa7720 where sscheckid=" & cQuote & Text1.Text & cQuote
        
        OpenQueryDNS "delete from pa7723 where sscheckid=" & cQuote & Text1.Text & cQuote, objdbRs, True
        Script2File "delete from pa7723 where sscheckid=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE SSCHECK #" & Text1.Text
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        
        ShowRecords
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nCtr As Integer
'    MsgBox KeyAscii
    Select Case KeyAscii
    
        Case 21     ' --> CTRL-U for update material detail...
            If nAdd > 0 Then
                
                With MSHFlexGrid1
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
                        If Val(.TextMatrix(nCtr, 16)) = 1 Then
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
                
            End If
        
        Case 13
            SendKeys vbTab
    End Select
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Tran No]:10:False", _
                    "TXT:[Emp ID]:8:True", _
                    "TXT:[Full Name]:35:True", _
                    "TXT:[Depid]:5:False", _
                    "TXT:[Department]:20:True", _
                    "TXT:[Position]:20:True", _
                    "TXT:[ShiftID]:7:True", _
                    "DAT:[LOGDATE]:15:False", _
                    "DAT:[TRANSDATE]:15:False", _
                    "TXT:[Shift Info]:15:True", _
                    "TXT:[Time In]:10:True", _
                    "TXT:[Tran Type]:10:False", _
                    "TXT:[Remark]:25:True", _
                    "NUM:[Seq No]:5:False", _
                    "NUM:[Status]:1:False", _
                    "NUM:[TAG]:5:True", _
                    "NUM:[STAG]:5:False", _
                    "TXT:[Emp Status]:35:True")

    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    XPDatePicker1.CurrentDate = Now
    XPDatePicker2.CurrentDate = Now
        
    OpenQueryDNS "SELECT * FROM pa7720 ORDER BY SSCheckId", oTempADO, False
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
                
            Case vbKeyReturn
                Select Case .ColSel
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
        End Select
    End With
End Sub

Private Sub MSHFlexGrid1_LeaveCell()
    nLastRow = MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid1_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then
        KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex")
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text2.Text) = "" Then
            Command2_Click
        Else
            ShowData Text2.Text, Label4, 1
            Text2.SetFocus
        End If
    End If

End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nRowPos As Integer, _
        cString As String
        
    Select Case KeyCode
        Case vbKeyReturn
            Select Case MSHFlexGrid1.ColSel
                Case 7
                    If Trim(txtFlex.Text) = "" Then
                        frmLookup.showPopup 9
                        frmLookup.Show 1
                        
                        If Trim(cResult) <> "" Then
                            nRowPos = MSHFlexGrid1.Row
                            InsertToGrid nRowPos, cResult
                        End If
                    Else
                        OpenQueryDNS "SELECT * FROM PA74380 WHERE SHIFTID=" & cQuote & txtFlex.Text & cQuote, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                            nRowPos = MSHFlexGrid1.Row
                            InsertToGrid nRowPos, txtFlex.Text
                        Else
                            MsgBox "Shift ID not found!", vbCritical, App.Title
                            Exit Sub
                        End If
                    End If
                    cString = shiftcheck(MSHFlexGrid1.TextMatrix(nRowPos, 7), nRowPos, MSHFlexGrid1.TextMatrix(nRowPos, 8), MSHFlexGrid1.TextMatrix(nRowPos, 9), MSHFlexGrid1.TextMatrix(nRowPos, 11))
                
                   If cString <> "" Then
                        If cString = "Correct Shifting" Then
                            HiLyt2 nRowPos, MSHFlexGrid1, vbBlack
                            MSHFlexGrid1.TextMatrix(nRowPos, 16) = 1
                        Else
                            HiLyt2 nRowPos, MSHFlexGrid1, vbBlue
                           MSHFlexGrid1.TextMatrix(nRowPos, 16) = 0
                        End If
                    Else
                        cString = "Undefined Shifting"
                        HiLyt2 nRowPos, MSHFlexGrid1, vbRed
                        MSHFlexGrid1.TextMatrix(nRowPos, 16) = 2
                    End If
                    
                    MSHFlexGrid1.TextMatrix(nRowPos, 13) = cString
                
                MSHFlexGrid1.Col = 7
                MSHFlexGrid1.ColSel = 7
            End Select
            
            txtFlex.Visible = False
            Command11.Cancel = True
            MSHFlexGrid1.SetFocus
            
        Case vbKeyEscape
            txtFlex.Visible = False
            Command11.Cancel = True
            MSHFlexGrid1.SetFocus
    End Select
End Sub

Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
    Command11.Cancel = True
End Sub
