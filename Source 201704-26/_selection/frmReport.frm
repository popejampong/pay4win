VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Manpower Report"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      BackColor       =   &H00800000&
      Caption         =   "Detailed Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   450
      Left            =   6630
      TabIndex        =   33
      Top             =   2790
      Width           =   1440
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00800000&
      Caption         =   "&No ATM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   450
      Left            =   6630
      TabIndex        =   32
      Top             =   2370
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Range"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   825
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmReport.frx":0000
      Left            =   120
      List            =   "frmReport.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5490
      Width           =   4095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Preview"
      Height          =   795
      Left            =   6765
      Picture         =   "frmReport.frx":0043
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   795
      Left            =   6765
      Picture         =   "frmReport.frx":090D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1155
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   105
      TabIndex        =   11
      Top             =   1245
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Department"
      TabPicture(0)   =   "frmReport.frx":11D7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Signature"
      TabPicture(1)   =   "frmReport.frx":11F3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command16"
      Tab(1).Control(1)=   "Text8"
      Tab(1).Control(2)=   "Command13"
      Tab(1).Control(3)=   "Text6"
      Tab(1).Control(4)=   "Command15"
      Tab(1).Control(5)=   "Text7"
      Tab(1).Control(6)=   "Text5"
      Tab(1).Control(7)=   "Command14"
      Tab(1).Control(8)=   "Command5"
      Tab(1).Control(9)=   "Text1"
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(11)=   "Label16"
      Tab(1).Control(12)=   "Label6"
      Tab(1).Control(13)=   "Label10"
      Tab(1).Control(14)=   "Label15"
      Tab(1).Control(15)=   "Label4"
      Tab(1).Control(16)=   "Label11"
      Tab(1).Control(17)=   "Label5"
      Tab(1).Control(18)=   "Label7"
      Tab(1).Control(19)=   "Label3"
      Tab(1).ControlCount=   20
      Begin VB.CommandButton Command16 
         Caption         =   "..."
         Height          =   315
         Left            =   -73215
         TabIndex        =   21
         Top             =   1680
         Width           =   375
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
         Left            =   -73905
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   1695
         Width           =   660
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   315
         Left            =   -73215
         TabIndex        =   20
         Top             =   405
         Width           =   375
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
         Left            =   -73905
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "TXT:PREP_BY"
         Top             =   435
         Width           =   660
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   315
         Left            =   -73215
         TabIndex        =   19
         Top             =   1365
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
         Left            =   -73905
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "TXT:INSP_BY"
         Top             =   1380
         Width           =   660
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
         Left            =   -73905
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   750
         Width           =   660
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   315
         Left            =   -73215
         TabIndex        =   18
         Top             =   735
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   315
         Left            =   -73215
         TabIndex        =   17
         Top             =   1050
         Width           =   375
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
         Left            =   -73905
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "TXT:REC_BY"
         Top             =   1065
         Width           =   660
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3210
         Left            =   135
         TabIndex        =   2
         Top             =   405
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5010
         EndProperty
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&Select All"
         Height          =   315
         Left            =   135
         TabIndex        =   16
         Top             =   3585
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         Left            =   -74880
         TabIndex        =   31
         Top             =   1725
         Width           =   1215
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
         Left            =   -72765
         TabIndex        =   30
         Top             =   1755
         Width           =   3930
      End
      Begin VB.Label Label6 
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
         Left            =   -72765
         TabIndex        =   29
         Top             =   450
         Width           =   3930
      End
      Begin VB.Label Label10 
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   450
         Width           =   1215
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
         Left            =   -72765
         TabIndex        =   27
         Top             =   1425
         Width           =   3930
      End
      Begin VB.Label Label4 
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
         Left            =   -72765
         TabIndex        =   26
         Top             =   765
         Width           =   3930
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Note By"
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Verified By"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   1095
         Width           =   1215
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
         Left            =   -72765
         TabIndex        =   22
         Top             =   1095
         Width           =   3930
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_ISS"
      Top             =   105
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   609
      _Version        =   393216
      Format          =   123142144
      CurrentDate     =   38623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   720
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:DATE_ISS"
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   609
      _Version        =   393216
      Format          =   123142144
      CurrentDate     =   38623
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   120
      TabIndex        =   15
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   120
      TabIndex        =   14
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Type of Report to Generate:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   5250
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   5925
      Left            =   6435
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Cost Accounting System
' module        :   frmReport
' programmer    :   _-=[ srm ]=-_
' date          :   9 feb 2005

Option Explicit
    Dim oTempADO As New ADODB.Recordset
    
Sub GetRpt(ByVal nMode As Integer)

    Label3.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label15.Caption = ""
    Label16.Caption = ""

    SSTab1.TabVisible(1) = True
    Check1.Visible = True

    Tag = nMode
    Select Case nMode
        Case 1
            Caption = "Daily Manpower Report"
            With Combo1
                .Clear
                .AddItem "Detailed Report"
                .AddItem "Summary Report"
                .ListIndex = 0
            End With
            
            Check1.Visible = False
            Check4.Visible = False
            
        Case 2
            Caption = "Daily Manpower Late/Absent/Incomplete Report"
            With Combo1
                .Clear
                .AddItem "Late"
                .AddItem "Absent"
                .AddItem "Incomplete Entry"
                .ListIndex = 0
            End With
            
            SSTab1.TabVisible(1) = False
            Check1_Click
            
        Case 3
            Caption = "Daily Manpower OT Report"
            Combo1.Visible = False
            SSTab1.TabVisible(1) = False
            Label2.Visible = False
            Check1_Click
            
        Case 4
            Caption = "Daily Time Consumption Report"
            Combo1.Visible = True
            SSTab1.TabVisible(1) = False
            Label2.Visible = True
            Check1.Visible = False
            
            Check2.Value = vbChecked
            
            With Combo1
                .Clear
                .AddItem "OT Hour from 1-10.5"
                .AddItem "OT Hour from 11-16.5"
                .ListIndex = 0
            End With
            
            
        Case 5     '----------------------------------------------> actual manpower
            Caption = "Daily Actual Manpower Report"
            
            SSTab1.TabVisible(1) = False
            Check2.Visible = True
            SSTab1.top = 800

            Label2.Visible = True
            Label2.top = 5000
            Combo1.top = 5200
            With Combo1
                .Clear
                .AddItem "Summary Report"
                .AddItem "Detailed Report"
                .ListIndex = 0
            End With
            
            
            Check1.Visible = False
    
        Case 6     '----------------------------------------------> actual weekly Labor Cost
            Caption = "Actual Weekly Labor Cost"
            
            Check2.Visible = True
            Check4.Visible = False
            SSTab1.TabVisible(1) = False
            SSTab1.top = 800

            DTPicker1.Visible = True
            DTPicker2.Visible = True
            
            Label8.Visible = True
            Label1.Visible = True

            Label2.Visible = True
            Label2.top = 5000
            Combo1.top = 5200
            With Combo1
                .Clear
                .AddItem "Regular"
                .AddItem "SA"
                .AddItem "WAP"
                .AddItem "WAP SA"
                .AddItem "Emergency"
                .ListIndex = 0
            End With
            
            Check1.Visible = False
        Case 7
            Caption = "ERP-Employee List "
            Check1.Visible = False
            Check4.Visible = False
            Combo1.Visible = False
            Label2.Visible = False
            DTPicker2.Visible = False
    
        Case 8
            Caption = "TMS"
    
            Check2.Visible = True
            Check4.Visible = False
            SSTab1.TabVisible(1) = False
            SSTab1.top = 800
            
            Combo1.Visible = False
            Label2.Visible = False

            DTPicker1.Visible = True
            DTPicker2.Visible = False
            Check1.Visible = False
            
            Label8.Visible = True
            Label1.Visible = True
            
      Case 9
      
            Caption = "Daily Labor Cost"
            
            Check2.Visible = True
            Check4.Visible = False
            SSTab1.TabVisible(1) = False
            SSTab1.top = 800

            DTPicker1.Visible = True
            DTPicker2.Visible = True
            
            Label8.Visible = True
            Label1.Visible = True

            Label2.Visible = True
            Label2.top = 5000
            Combo1.top = 5200
            With Combo1
                .Clear
                .AddItem "Detailed Report"
                .AddItem "Summary Report"
              
                .ListIndex = 0
            End With
            
            Check1.Visible = False
            
            
            
     Case 10
           Caption = "Weekly Consumption Report"
           Combo1.Visible = False
           Label12.Visible = False
           Check4.Visible = False
           Check6.Visible = False
           SSTab1.TabVisible(1) = False
           ListView1.Enabled = False
           Check2.Visible = False
           Label2.Visible = False
           ListView1.Visible = False
           SSTab1.Enabled = False
           SSTab1.Visible = False
           Check1_Click
           
    
    End Select
    
End Sub

Function ChkLeave(ByVal oRecordSet As ADODB.Recordset) As Variant
    Dim nCtr As Integer, _
        aLeaveInfo As Variant, _
        cSqlStmt As String, _
        dStartDate As Date, dEndDate As Date
        
    aLeaveInfo = Array(0#, 0#)
    
    oRecordSet.MoveFirst
    While Not oRecordSet.EOF
        dStartDate = oRecordSet("date_start")
        dEndDate = oRecordSet("date_end")
        For nCtr = 0 To DateDiff("d", dStartDate, dEndDate)
            If Weekday(DateAdd("d", nCtr, dStartDate)) <> vbSunday Then
                aLeaveInfo(0) = aLeaveInfo(0) + 1
            End If
            cSqlStmt = "select * from pa4329 " & _
                       " where (date=" & cQuote & Format(DateAdd("d", nCtr, dStartDate), "yyyy-mm-dd") & cQuote & ")" & _
                       " or ((month(date)=" & Month(DateAdd("d", nCtr, dStartDate)) & ") and (day(date)=" & Day(DateAdd("d", nCtr, dStartDate)) & ") and (fix_day=1))"
            OpenQueryDNS cSqlStmt, objdbRs, False
            If (objdbRs.RecordCount > 0) And (Weekday(DateAdd("d", nCtr, dStartDate)) <> vbSunday) Then
                aLeaveInfo(1) = aLeaveInfo(1) + 1
            End If
        Next nCtr
        oRecordSet.MoveNext
    Wend
    
    ChkLeave = aLeaveInfo
End Function
    
' --> compute dtr summary here...
Function CheckDTR(ByVal cEmpID As String, _
                  aPeriodInfo As Variant, _
                  ByVal aEmpStat As Variant) As Variant
    Dim cSqlStmt As String, _
        oDTRRSet As New ADODB.Recordset, _
        aTimeInfo As Variant, _
        nCount As Integer, _
        cDateEnd As String, _
        cParam As String

    Dim oRecordSet As New ADODB.Recordset
    '2010-04-19
    '16-regot
    '17-saregot
    '18-ndregot
    '19-ndsaregot
'    If (gCompanyID = "0007") Or (gCompanyID = "0003") Then
'        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    Else
'        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#)
'    End If

    aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    
    
    If oTempADO("active") > 0 Then
        cDateEnd = Format(IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd")

        cSqlStmt = "select count(holidayid) as tot_day  from PA4329 " & _
                   "where (date between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & cDateEnd & cQuote & ") " & _
                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(cDateEnd) & ") and (fix_day=1))"
    Else
        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
                   "where (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1))"
    End If
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
        nCount = objdbRs("tot_day")


    ' --> for regular employee only
    cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
               "where ((date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
               "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1)))" & _
                   " and (tag=1)"
'        Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If aEmpStat(0) <> 2 Then
            nCount = nCount - objdbRs("tot_day")
        End If
    End If

'    '20101-01-13
'    If aEmpStat(0) = 2 Then
'        ' --> for regular employee only
'        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
'                   "where ((date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1)))" & _
'                       " and (tag=1)"
''        Script2File cSqlStmt
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            If aEmpStat(0) <> 2 Then
'                nCount = nCount - objdbRs("tot_day")
'            End If
'        Else
'            nCount = 0
'        End If
'    End If

    ' hired date between the selected period...
    If (DateDiff("d", aPeriodInfo(0), oTempADO("date_hire")) >= 0) And (DateDiff("d", oTempADO("date_hire"), aPeriodInfo(1)) >= 0) Then
        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
                   "where (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & ") " & _
                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(oTempADO("date_hire")) & ") and (fix_day=1))"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            nCount = nCount - objdbRs("tot_day")
        End If
    End If

    If nCount < 0 Then nCount = 0
'    '2010-04-19
'    cSqlStmt = "select EMPID, PERIODID, DATE, SHIFTID, " & _
'               "  sum(reg_hr/8) as reg_day, sum(reg_ot_hr) as reg_ot, sum(sa_reg_ot) as sa_reg_ot, " & _
'               "  sum(nd_hr/8) as nd_day, sum(nd_ot_hr) as nd_ot, sum(sa_nd_ot) as sa_nd_ot, " & _
'               "  sum(sun_hr) as sun_hr, sum(sun_ot_hr) as sun_ot, " & _
'               "  sum(sun_nd) as sun_nd, sum(sun_nd_ot) as sun_nd_ot, " & _
'               "  sum(inc_hr) as inc_hr " & _
'               "From di36770 " & _
'               "where (empid=" & cQuote & cEmpID & cQuote & ") " & _
'               "  and (date = " & cQuote & "2010-04-09" & cQuote & " ) " & _
'               "group by empid "
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oDTRRSet, False
'    If oDTRRSet.RecordCount > 0 Then
'        If aEmpStat(0) > 0 And aEmpStat(2) = 0 Then
'            aTimeInfo(16) = oDTRRSet("reg_ot")                               ' --> Reg OT
'            aTimeInfo(17) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
'            aTimeInfo(18) = oDTRRSet("nd_ot")                                ' --> NDiff OT
'            aTimeInfo(19) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
''        Else
''            MsgBox "stop"
'        End If
'    End If

    cSqlStmt = "select EMPID, PERIODID, DATE, SHIFTID, " & _
               "  sum(reg_hr/8) as reg_day, sum(reg_ot_hr) as reg_ot, sum(sa_reg_ot) as sa_reg_ot, " & _
               "  sum(nd_hr/8) as nd_day, sum(nd_ot_hr) as nd_ot, sum(sa_nd_ot) as sa_nd_ot, " & _
               "  sum(sun_hr) as sun_hr, sum(sun_ot_hr) as sun_ot, " & _
               "  sum(sun_nd) as sun_nd, sum(sun_nd_ot) as sun_nd_ot, " & _
               "  sum(inc_hr) as inc_hr " & _
               "From di36770 " & _
               "where (empid=" & cQuote & cEmpID & cQuote & ") " & _
               "  and (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ") " & _
               "group by empid "
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oDTRRSet, False
    If oDTRRSet.RecordCount > 0 Then

        aTimeInfo(0) = oDTRRSet("reg_day")      ' --> Reg Day
        aTimeInfo(3) = oDTRRSet("nd_day")       ' --> NDiff Day
        aTimeInfo(5) = oDTRRSet("sun_hr")       ' --> Sunday
        aTimeInfo(6) = oDTRRSet("sun_ot")       ' --> Sunday OT
        aTimeInfo(13) = oDTRRSet("sun_nd")      ' --> Sunday ND
        aTimeInfo(14) = oDTRRSet("sun_nd_ot")   ' --> Sunday NDiff OT
        aTimeInfo(15) = oDTRRSet("inc_hr")   ' --> Incentive Hour

        If aEmpStat(2) > 0 Then
'            aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
'            aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
'            aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
'            aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
'            aTimeInfo(7) = 0                                                ' --> No Holiday

            If gCompanyID = "0002" Then

                If lAudit = 1 Then
                    aTimeInfo(1) = oDTRRSet("reg_ot") + oDTRRSet("sa_reg_ot")       ' --> Reg OT
                    aTimeInfo(4) = oDTRRSet("nd_ot") + oDTRRSet("sa_nd_ot")         ' --> NDiff OT
                    aTimeInfo(2) = 0                                                ' --> SA Reg OT
                    aTimeInfo(12) = 0                                               ' --> SA NDiff OT
                Else
                    aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
                    aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
                    aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
                    aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
                End If
            Else
                aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
                aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
                aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
                aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
            End If
                aTimeInfo(7) = 0                                                ' --> No Holiday
        Else
            If gCompanyID = "0002" Then
                If lAudit = 1 Then
                    aTimeInfo(1) = oDTRRSet("reg_ot") + oDTRRSet("sa_reg_ot")       ' --> Reg OT
                    aTimeInfo(4) = oDTRRSet("nd_ot") + oDTRRSet("sa_nd_ot")         ' --> NDiff OT
                    aTimeInfo(2) = 0                                                ' --> SA Reg OT
                    aTimeInfo(12) = 0                                              ' --> SA NDiff OT
                Else
                    aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
                    aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
                    aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
                    aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
                End If
            Else
                aTimeInfo(1) = oDTRRSet("reg_ot")       ' --> Reg OT
                aTimeInfo(2) = oDTRRSet("sa_reg_ot")    ' --> SA Reg OT
                aTimeInfo(4) = oDTRRSet("nd_ot")        ' --> NDiff OT
                aTimeInfo(12) = oDTRRSet("sa_nd_ot")    ' --> SA NDiff OT

                'aTimeInfo(7) = IIf((aEmpStat(0) <> 0) And (Not ((aEmpStat(0) = 1) And (aEmpStat(1) = 1))), nCount, 0)

            End If
                aTimeInfo(7) = IIf((aEmpStat(0) <> 0) And (Not ((aEmpStat(0) = 1) And (aEmpStat(1) = 1))), nCount, 0)
        End If
    End If

    CheckDTR = aTimeInfo

    Set oDTRRSet = Nothing
    Set oRecordSet = Nothing
End Function


Sub cmdClick(ByVal oTxtBox As TextBox, ByVal oLabel As Label)
    frmLookup.showPopup 1   ', " where sysuser = 1"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTxtBox.Text = cResult
        ShowData cResult, oLabel
    End If
End Sub

Sub ShowData(cString As String, oLabel As Label)
    OpenQueryDNS "SELECT USERID,CONCAT(FIRSTNAME," & cQuote & " " & cQuote & ",LASTNAME) AS FULLNAME FROM PA2360 WHERE USERID=" & cQuote & cString & cQuote, objdbRs, False
    oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("FULLNAME"), "")
End Sub

Sub txtKeyDown(nMode As Integer, cString As String, oLabel As Label)
'    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command13_Click
                Case 2
                    Command14_Click
                Case 3
                    Command5_Click
                Case 4
                    Command15_Click
                Case 5
                    Command16_Click
            End Select
        Else
            ShowData cString, oLabel
        End If
'    End If
End Sub


' + -->
' |     Procedure Name  :   GenManpower(byval nMode as Integer)
' |     Description     :   Generate Daily Manpower Report
' |     Date Created    :   20 sep 2007
' + -->
'   where nMode
'       0   -   Detailed Report
'       1   -   Summary Report

Sub Create_Manpower(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTableName As String

    Select Case nMode
        Case 0
            cTableName = "tmpDManpower"
            cSqlStmt = " ([DATE] date, " & _
                       " [empid] char(6),      [emp_stat] integer," & _
                       " [fullname] char(100),  [wap] integer, " & _
                       " [tag] integer,         [paystatus] integer, " & _
                       " [depid] char(3),       [department] char(100), " & _
                       " [position] char(100),  [shift] char(100),      [Grand_OT] double,       [W_Tot] double, " & _
                       " [intime] char(15),     [outtime] char(15),     [Remark] char (100), " & _
                       " [PREPPOS] char(50),    [PREP_NAME] char(100),  [CHKPOS] char(100),    [CHK_NAME] char(100)," & _
                       " [VERPOS] char(100),    [VER_NAME] char(100),   [NOTEPOS] char(100),   [NOTE_NAME] char(100)," & _
                       " [APPRPOS] char(100),   [APPR_NAME] char(100),  [POSNAME] CHAR(100),   [tag_comp] integer, " & _
                       " [tag_Incomp] integer,  [tag_Leave] integer,    [tag_Absent] integer,  [tag_FCRes] integer, " & _
                       " [rn_hr_tot] double,     [rn_tot] double,      [rnsa_tot] double )"
            
        Case 1
            cTableName = "tmpSManpower"
            cSqlStmt = "([DATE_REG] date, " & _
                       " [LineID] char(3),      [LineName] char(100),  " & _
                       " [PROJ_REG] double,     [PROJ_CON] double,      [PROJ_WAP] double,      [PROJ_TTL] double, " & _
                       " [FLEAVE_REG] double,   [FLEAVE_CON] double,    [FLEAVE_WAP] double,    [FLEAVE_TTL] double, " & _
                       " [EMERG_CON] double,    [EMERG_WAP] double,     [EMERG_TTL] double, " & _
                       " [ACTUAL_REG] double,   [ACTUAL_CON] double,    [ACTUAL_WAP] double,    [ACTUAL_TRANS] double, [ACTUAL_TTL] double, " & _
                       " [BAL_CON] double,      [BAL_WAP] double, " & _
                       " [ABSENT_REG] double,   [ABSENT_CON] double,    [ABSENT_WAP] double,    [ABSENT_TTL] double, " & _
                       " [PREPPOS] char(50),    [PREP_NAME] char(100)," & _
                       " [CHKPOS] char(100),    [CHK_NAME] char(100)," & _
                       " [VERPOS] char(100),    [VER_NAME] char(100)," & _
                       " [NOTEPOS] char(100),   [NOTE_NAME] char(100)," & _
                       " [APPRPOS] char(100),   [APPR_NAME] char(100)," & _
                       " [PROJ_REG_M] double,   [PROJ_CON_M] double,    [PROJ_WAP_E] double, " & _
                       " [ACT_REG_M] double,    [ACT_CON_M] double,     [ACT_WAP_E] double, " & _
                       " [ABS_REG_M] double,    [ABS_CON_M] double,     [ABS_WAP_E] double, " & _
                       " [PROJ_WAP_C] double, " & _
                       " [ACT_WAP_C] double, " & _
                       " [ABS_WAP_C] double, " & _
                       " [remark] char(100))"
    End Select
    
    cSqlStmt = "CREATE TABLE " & cTableName & cSqlStmt
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM " & cTableName
    QueryTemp cSqlStmt, oTempADO, True

End Sub


Sub GenManpower(ByVal nMode As Integer, ByVal cParam As String)
    Dim cLineID As String, _
        cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        nAddEdit As Integer, _
        ntag_comp, ntag_Incomp, ntag_Leave, ntag_Absent, ntag_FC_Res As Integer, _
        aOtherInfo As Variant
    
    
    'tag_comp,tag_Incomp,tag_Leave,tag_Absent
    
    Create_Manpower nMode
    
    ntag_comp = 0
    ntag_Incomp = 0
    ntag_Leave = 0
    ntag_Absent = 0
    ntag_FC_Res = 0
    
   
    ShowProgress 0
    
    If nMode = 0 Then
        aOtherInfo = Array("", "", "", "")
    
        cSqlStmt = "select pclose from pa7730 " & _
                   "where " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " between date_start and date_end"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
        
            cSqlStmt = "select " & _
                       "  b.date, " & _
                       "  a.depid, ifnull(c.linename,'') as department, " & _
                       "  a.emp_stat, a.wap, if(a.paystatus=2,2,0) as paystatus, " & _
                       "  a.empid, " & _
                       "  concat(a.lastname,', ',a.firstname,if(trim(a.mname)='','',concat(' ',left(a.mname,1),'.'))) as fullname, " & _
                       "  ifnull(d.posname,'') as position, " & _
                       "  ifnull(concat(date_format(concat(b.date,' ',e.time1),'%l:%i %p'),' - ',date_format(concat(b.date,' ',e.time2),'%l:%i %p')),'') as shift, " & _
                       "  ifnull(b.Tag, 4) As Tag,ifnull((b.tot_ot+nd_tot_ot),0)as grand_ot, " & _
                       "  ifnull((b.reg_hr + b.reg_ot_hr + b.sa_reg_ot + b.nd_hr + b.nd_ot_hr + b.sa_nd_ot + b.sun_hr + b.sun_ot_hr + b.sun_nd + b.sun_nd_ot),0) as w_tot, " & _
                       "  ifnull((b.reg_hr + b.nd_hr),0) as rn_hr_tot, " & _
                       "  ifnull((b.reg_ot_hr + b.nd_ot_hr),0) as rn_tot, " & _
                       "  ifnull((b.sa_reg_ot + b.sa_nd_ot),0) as rnsa_tot " & _
                       "from ((((di3670 a " & _
                       "  left join di5463 c on a.depid=c.lineid) " & _
                       "  left join di7670 d on a.posid=d.posid) " & _
                       "  left join " & IIf(objdbRs("pclose") = 1, "dih36770", "di36770") & " b on (a.empid=b.empid) and (b.date=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) " & _
                       "  left join pa74380 e on b.shiftid=e.shiftid) " & _
                       "Where ((((a.active = 1) Or (a.active = 3)) And ((a.date_res >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") Or ((a.date_hire <= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") And (a.date_res > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
                       "  or ((a.active=2) and ((a.date_fin >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and(a.date_fin > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
                       "  or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))) " & _
                       IIf(Trim(cParam) <> "", " and (a.depid in " & cParam & ")", "")
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, oTempADO, False
            If oTempADO.RecordCount > 0 Then
                While Not oTempADO.EOF
                
                    ntag_comp = 0
                    ntag_Incomp = 0
                    ntag_Leave = 0
                    ntag_Absent = 0
                    ntag_FC_Res = 0
                    aOtherInfo = Array("", "", "", "")
    
                    ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
                    
                    cSqlStmt = "select empid, logdate, transdate, date_format(concat(logdate,' ',trantime),'%l:%i %p') as in_out, trantype " & _
                               "from " & IIf(objdbRs("pclose") = 1, "pah84650", "pa84650") & " " & _
                               "where (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (logdate=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") " & _
                               "order by logdate, transdate, trantime "
                    OpenQueryDNS cSqlStmt, oRecordSet, False
                    If oRecordSet.RecordCount > 0 Then
                        While Not oRecordSet.EOF
                            If oRecordSet("trantype") = 0 Then
                                If Trim(aOtherInfo(1)) = "" Then
                                    aOtherInfo(0) = oRecordSet("trantype")
                                    aOtherInfo(1) = oRecordSet("in_out")
                                End If
                            Else
                                aOtherInfo(0) = oRecordSet("trantype")
                                aOtherInfo(2) = oRecordSet("in_out")
                            End If
                            oRecordSet.MoveNext
                        Wend
                    End If
                    
                    Select Case oTempADO("Tag")
                        Case 0 'Complete
                            ntag_comp = 1
                        Case 1 'No entry
                            ntag_Absent = 1
                        Case 2 ' Leave
                            ntag_Leave = 1
                        Case 3 'Incomplete
                            ntag_Incomp = 1
                        Case 4 'FC/Res
                            ntag_FC_Res = 1
                            
                    End Select

                    cSqlStmt = "insert into tmpDManpower([date],[empid],[fullname],emp_stat,[wap],paystatus,[depid],[department]," & _
                               " [position],[shift]," & _
                               " [intime],[outtime],[grand_ot],[w_tot],[tag],tag_comp,tag_Incomp,tag_Leave,tag_Absent,tag_FCRes," & _
                               " rn_hr_tot,rn_tot,rnsa_tot)values(" & _
                               cQuote & Format(DTPicker1.Value, "mm/dd/yyyy") & cQuote & "," & _
                               cQuote & oTempADO("empid") & cQuote & "," & _
                               cQuote & oTempADO("fullname") & cQuote & "," & _
                               oTempADO("emp_stat") & "," & _
                               oTempADO("wap") & "," & _
                               oTempADO("paystatus") & "," & _
                               cQuote & oTempADO("depid") & cQuote & "," & _
                               cQuote & oTempADO("department") & cQuote & "," & _
                               cQuote & oTempADO("position") & cQuote & "," & _
                               cQuote & oTempADO("shift") & cQuote & "," & _
                               cQuote & aOtherInfo(1) & cQuote & "," & _
                               cQuote & aOtherInfo(2) & cQuote & "," & _
                               cQuote & oTempADO("grand_ot") & cQuote & "," & _
                               cQuote & oTempADO("w_tot") & cQuote & "," & _
                               oTempADO("tag") & "," & _
                               ntag_comp & "," & ntag_Incomp & "," & ntag_Leave & "," & ntag_Absent & "," & ntag_FC_Res & "," & _
                               oTempADO("rn_hr_tot") & "," & _
                               oTempADO("rn_tot") & "," & _
                               oTempADO("rnsa_tot") & ")"
                               
                               
                    QueryTemp cSqlStmt, objdbRs, True
                    
                    oTempADO.MoveNext
                Wend
                
                'sdhsadhsa
                
                
            End If
            
            ShowProgress 3
            
            GenerateReport "Daily Attendance Report for " & Format(DTPicker1.Value, "dddd, mmmm d, yyyy"), "RPTDMAN2.RPT"
        End If
        
    Else
        aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, "")
'        aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, "")
        
        cSqlStmt = "select pclose from pa7730 " & _
                   "where " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " between date_start and date_end"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
        
'            cSqlStmt = "select a.depid, ifnull(c.linename,'') as department, a.emp_stat, a.wap, a.paystatus, ifnull(b.tag,4) as tag, count(b.tag) as ttl, count(a.empid) as ettl " & _
'                       "from di3670 a " & _
'                       "  left join " & IIf(objdbRs("pclose") = 1, "dih36770", "di36770") & " b on a.empid=b.empid and b.date=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
'                       "  left join di5463 c on a.depid=c.lineid " & _
'                       "where (((a.active=1) or (a.active=3)) and ((a.date_res >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
'                       "  or ((a.active=2) and ((a.date_fin >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and(a.date_fin > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
'                       "  or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) " & _
'                       "group by a.depid, a.emp_stat, a.wap, b.tag " & _
'                       " order by a.depid,a.emp_stat, a.wap,b.tag "
                       
            cSqlStmt = "select a.depid, ifnull(c.linename,'') as department, a.emp_stat, a.wap, a.paystatus, ifnull(b.tag,4) as tag " & _
                       "from di3670 a " & _
                       "  left join " & IIf(objdbRs("pclose") = 1, "dih36770", "di36770") & " b on a.empid=b.empid and b.date=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                       "  left join di5463 c on a.depid=c.lineid " & _
                       "where (((a.active=1) or (a.active=3)) and ((a.date_res >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
                       "  or ((a.active=2) and ((a.date_fin >= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and(a.date_fin > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
                       "  or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) " & _
                       " order by a.depid,a.emp_stat, a.wap,b.tag "
'            Script2File cSqlStmt
            
            OpenQueryDNS cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
            
                While Not oRecordSet.EOF
                
                    ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
                    
                    If cLineID <> oRecordSet("depid") Then
                        'aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, "")
                        aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, "")
                        cLineID = oRecordSet("depid")
                        nAddEdit = 0
                    Else
                        nAddEdit = 1
                    End If
                    
                    Select Case oRecordSet("emp_stat")
                        Case 0      ' --> WAP
'                            If oRecordSet("depid") = "005" Then MsgBox "stop"
                            If oRecordSet("paystatus") = 2 Then
                                aOtherInfo(14) = aOtherInfo(14) + 1
                            Else
                                aOtherInfo(2) = aOtherInfo(2) + 1
                            End If
                            
                            If (oRecordSet("tag") = 0) Or (oRecordSet("tag") = 3) Then
                                If oRecordSet("paystatus") = 2 Then
                                    aOtherInfo(17) = aOtherInfo(17) + 1
                                Else
                                    aOtherInfo(6) = aOtherInfo(6) + 1
                                End If
                            Else
                                If oRecordSet("paystatus") = 2 Then
                                    aOtherInfo(20) = aOtherInfo(20) + 1
                                Else
                                    aOtherInfo(10) = aOtherInfo(10) + 1
                                End If
                            End If
                            
                        Case 1      ' --> Contractual
                            If oRecordSet("wap") = 0 Then
                                If oRecordSet("paystatus") = 1 Then  ' ----->  projected monthly
                                    aOtherInfo(13) = aOtherInfo(13) + 1 ' ---- con monthly
                                Else
                                    aOtherInfo(1) = aOtherInfo(1) + 1   ' ---- con daily
                                End If
                                
                                If (oRecordSet("tag") = 0) Or (oRecordSet("tag") = 3) Then
                                    If oRecordSet("paystatus") = 1 Then
                                        aOtherInfo(16) = aOtherInfo(16) + 1
                                    Else
                                        aOtherInfo(5) = aOtherInfo(5) + 1
                                    End If
                                Else
                                    If oRecordSet("paystatus") = 1 Then
                                        aOtherInfo(19) = aOtherInfo(19) + 1
                                    Else
                                        aOtherInfo(9) = aOtherInfo(9) + 1
                                    End If
                                End If
                            
                            Else
                                aOtherInfo(21) = aOtherInfo(21) + 1 ' ---> wap daily
                                
                                If (oRecordSet("tag") = 0) Or (oRecordSet("tag") = 3) Then
                                        aOtherInfo(22) = aOtherInfo(22) + 1
                                Else
                                        aOtherInfo(23) = aOtherInfo(23) + 1
                                End If

                            End If
                            
                        Case 2      ' --> Regular
                            If oRecordSet("paystatus") = 1 Then
                                aOtherInfo(12) = aOtherInfo(12) + 1
                            Else
                                aOtherInfo(0) = aOtherInfo(0) + 1
                            End If
                            If (oRecordSet("tag") = 0) Or (oRecordSet("tag") = 3) Then
                                If oRecordSet("paystatus") = 1 Then
                                    aOtherInfo(15) = aOtherInfo(15) + 1
                                Else
                                    aOtherInfo(4) = aOtherInfo(4) + 1
                                End If
                            Else
                                If oRecordSet("paystatus") = 1 Then
                                    aOtherInfo(18) = aOtherInfo(18) + 1
                                Else
                                    aOtherInfo(8) = aOtherInfo(8) + 1
                                End If
                            End If
                            
                    End Select
                    
                    aOtherInfo(3) = aOtherInfo(0) + aOtherInfo(1) + aOtherInfo(2) + aOtherInfo(12) + aOtherInfo(13) + aOtherInfo(14) + aOtherInfo(21)
                    aOtherInfo(7) = aOtherInfo(4) + aOtherInfo(5) + aOtherInfo(6) + aOtherInfo(15) + aOtherInfo(16) + aOtherInfo(17) + aOtherInfo(22)
                    aOtherInfo(11) = aOtherInfo(8) + aOtherInfo(9) + aOtherInfo(10) + aOtherInfo(18) + aOtherInfo(19) + aOtherInfo(20) + aOtherInfo(23)
                    
                    If nAddEdit = 0 Then
                        cSqlStmt = "insert into tmpSManpower([date_reg],[lineid],[linename]," & _
                                   "[proj_reg],[proj_con],[proj_wap],[proj_ttl]," & _
                                   "[actual_reg],[actual_con],[actual_wap],[actual_ttl]," & _
                                   "[absent_reg],[absent_con],[absent_wap],[absent_ttl]," & _
                                   "[PROJ_REG_M], [PROJ_CON_M], [PROJ_WAP_E], [ACT_REG_M], [ACT_CON_M], [ACT_WAP_E], [ABS_REG_M], [ABS_CON_M], [ABS_WAP_E], [PROJ_WAP_C], [ACT_WAP_C],[ABS_WAP_C]  )values(" & _
                                   cQuote & Format(DTPicker1.Value, "mm/dd/yyyy") & cQuote & "," & _
                                   cQuote & oRecordSet("depid") & cQuote & "," & _
                                   cQuote & oRecordSet("department") & cQuote & "," & _
                                   aOtherInfo(0) & "," & aOtherInfo(1) & "," & _
                                   aOtherInfo(2) & "," & aOtherInfo(3) & "," & _
                                   aOtherInfo(4) & "," & aOtherInfo(5) & "," & _
                                   aOtherInfo(6) & "," & aOtherInfo(7) & "," & _
                                   aOtherInfo(8) & "," & aOtherInfo(9) & "," & _
                                   aOtherInfo(10) & "," & aOtherInfo(11) & "," & _
                                   aOtherInfo(12) & "," & aOtherInfo(13) & "," & _
                                   aOtherInfo(14) & "," & aOtherInfo(15) & "," & _
                                   aOtherInfo(16) & "," & aOtherInfo(17) & "," & _
                                   aOtherInfo(18) & "," & aOtherInfo(19) & "," & _
                                   aOtherInfo(20) & "," & aOtherInfo(21) & "," & _
                                   aOtherInfo(22) & "," & aOtherInfo(23) & ")"
                    Else
                        cSqlStmt = "update tmpSManpower set " & _
                                   "[proj_reg]=" & aOtherInfo(0) & ", [proj_con]=" & aOtherInfo(1) & "," & _
                                   "[proj_wap]=" & aOtherInfo(2) & ", [proj_ttl]=" & aOtherInfo(3) & "," & _
                                   "[actual_reg]=" & aOtherInfo(4) & ", [actual_con]=" & aOtherInfo(5) & "," & _
                                   "[actual_wap]=" & aOtherInfo(6) & ", [actual_ttl]=" & aOtherInfo(7) & "," & _
                                   "[absent_reg]=" & aOtherInfo(8) & ", [absent_con]=" & aOtherInfo(9) & "," & _
                                   "[absent_wap]=" & aOtherInfo(10) & ", [absent_ttl]=" & aOtherInfo(11) & "," & _
                                   "[PROJ_REG_M]=" & aOtherInfo(12) & ", [PROJ_CON_M]=" & aOtherInfo(13) & "," & _
                                   "[PROJ_WAP_E]=" & aOtherInfo(14) & "," & _
                                   "[ACT_REG_M]=" & aOtherInfo(15) & "," & _
                                   "[ACT_CON_M]=" & aOtherInfo(16) & "," & _
                                   "[ACT_WAP_E]=" & aOtherInfo(17) & "," & _
                                   "[ABS_REG_M]=" & aOtherInfo(18) & "," & _
                                   "[ABS_CON_M]=" & aOtherInfo(19) & "," & _
                                   "[ABS_WAP_E]=" & aOtherInfo(20) & "," & _
                                   "[PROJ_WAP_C]=" & aOtherInfo(21) & "," & _
                                   "[ACT_WAP_C]=" & aOtherInfo(22) & "," & _
                                   "[ABS_WAP_C]=" & aOtherInfo(23) & _
                                   " where lineid=" & cQuote & oRecordSet("depid") & cQuote
                    End If
'                    Script2File cSqlStmt
                    QueryTemp cSqlStmt, objdbRs, True
                    
                    oRecordSet.MoveNext
                    
                Wend
                
            End If
            
            ShowProgress 3
    
            GenerateReport "Daily Attendance Report ", "RPTSMAN.RPT"
           
        End If
        
    End If
    
    ShowProgress 4
    
    Set oRecordSet = Nothing
End Sub


' + -->
' |     Procedure Name  :   GenLateAbsent(ByVal nMode As Integer, ByVal cDepid As String)
' |     Description     :   Generate Report for Late/Absent/Incomplete Entry
' |     Date Created    :   11 mar 2006
' + -->
Sub CreateTemp(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTblName As String

    Select Case nMode
        Case 0      ' --> Late/Absent Report
            cTblName = "tmpLateAbsent"
            cSqlStmt = "create table " & cTblName & _
                       "([empid] char(6),       [fullname] char(100), " & _
                       " [depid] char(3),       [department] char(100), " & _
                       " [date] date,           [time1] char(10), " & _
                       " [time2] char(10),      [shiftid] char(5), " & _
                       " [trantime] char(15),   [logdate] date, " & _
                       " [trantype] integer)"
        Case 1
            cTblName = "tmpEmpOt"
            cSqlStmt = " CREATE TABLE " & cTblName & _
                       " ([EMPID] char(6),           [FULLNAME] char(100), " & _
                       " [TCID] char(5),             [POSNAME] char(100), " & _
                       " [EMP_STAT] integer,         [ACTIVE] integer, " & _
                       " [PAYSTATUS] integer,        [DATE] date, " & _
                       " [SHIFTDESC] char(100),      " & _
                       " [TIME1] char(10),           [TIME2] char(10), " & _
                       " [reg_hr] double,            [reg_ot_hr] double, " & _
                       " [sa_reg_ot] double,         [tot_ot] double,       [nd_hr] double, " & _
                       " [nd_ot_hr] double,          [sa_nd_ot] double,     [nd_tot_ot] double, " & _
                       " [sun_hr] double,            [sun_ot_hr] double, " & _
                       " [sun_nd] double,            [sun_nd_ot] double, " & _
                       " [INTRANTIME] char(10),      [OUTTRANTIME] char(10), " & _
                       " [DEPNAME] char(100),        [SEQ_NO] integer )"
    
        Case 2
            cTblName = "tmpDTCons"
            
            cSqlStmt = " CREATE TABLE " & cTblName & _
                       " ([date] date,              [DEPNAME] char(100), " & _
                       " [EMPID] char(6),           [FULLNAME] char(100), " & _
                       " [TCID] char(5),            [POSNAME] char(100), " & _
                       " [EMP_STAT] integer,        [ACTIVE] integer, " & _
                       " [PAYSTATUS] integer,       [Manpower] integer, " & _
                       " [reg_hr] integer,          [reg_hr1] integer, " & _
                       " [absent] integer, " & _
                       " [ST_R] integer,            [ST_R_M] integer, " & _
                       " [ST_C] integer,            [ST_C_M] integer, " & _
                       " [ST_W] integer,            [ST_W_C] integer, " & _
                       " [ST_C_E] integer,          [ST_W_E] integer, " & _
                       " [hour0] integer,           [hour_5] integer,       [hour1] integer,            [hour1_5] integer, " & _
                       " [hour2] integer,           [hour2_5] integer,      [hour3] integer,            [hour3_5] integer, " & _
                       " [hour4] integer,           [hour4_5] integer,      [hour5] integer,            [hour5_5] integer, " & _
                       " [hour6] integer,           [hour6_5] integer,      [hour7] integer,            [hour7_5] integer, " & _
                       " [hour8] integer,           [hour8_5] integer,      [hour9] integer,            [hour9_5] integer, " & _
                       " [hour10] integer,          [hour10_5] integer,     [hour11] integer,          [hour11_5] integer, " & _
                       " [hour12] integer,          [hour12_5] integer,     [hour13] integer,          [hour13_5] integer, " & _
                       " [hour14] integer,          [hour14_5] integer,     [hour15] integer,          [hour15_5] integer, " & _
                       " [hour16] integer,          [hour16_5] integer,     [hour17] integer  )"


        Case 3
            cTblName = "tmpDActMan"    ' -------------------------------------------------------->table for actual manpower
            cSqlStmt = " CREATE TABLE " & cTblName & _
                       " ([date] date,              [DEPNAME] char(100), " & _
                       " [EMPID] char (20),         [FULLNAME] char(100), " & _
                       " [REG]integer,              [CONT]integer, " & _
                       " [WAP] integer,             [WAPC] integer, " & _
                       " [REMARK] integer )"
    
    
        Case 4
            cTblName = "tmpAWLCost"    ' -------------------------------------------------------->table for actual manpower
            cSqlStmt = " CREATE TABLE " & cTblName & "(" & _
                       " [EMPID] char(6),                     [ACTIVE] integer,                   [EMP_STAT] integer,                  [RATE_AMT] decimal(18,4), " & _
                       " [COLA_AMT] decimal(18,4),            [POS_ALLOW] decimal(18,4),          [REG_DAY] decimal(18,4),             [REG_PAY] decimal(18,4), " & _
                       " [REG_OT_HR] decimal(18,4),           [REG_OT_PAY] decimal(18,4),         [NDIFF_DAY] decimal(18,4),           [NDIFF_PAY] decimal(18,4), " & _
                       " [NDIFF_OT_HR] decimal(18,4),         [NDIFF_OT_PAY] decimal(18,4),       [HOLIDAY] decimal(18,4),             [HOL_PAY] decimal(18,4), " & _
                       " [SA_REG_OT] decimal(18,4),           [SA_REG_PAY] decimal(18,4),         [SA_NDIFF_OT] decimal(18,4),         [SA_NDIFF_PAY] decimal(18,4), " & _
                       " [SUN_HR] decimal(18,4),              [SUN_PAY] decimal(18,4),            [SUN_OT] decimal(18,4),              [SUN_OT_PAY] decimal(18,4), " & _
                       " [ADJ_PAY] decimal(18,4),             [SA_ADJ_PAY] decimal(18,4),         [OTHER_PAY] decimal(18,4),           [LEAVE_PAY] decimal(18,4), " & _
                       " [M13PAY] decimal(18,4),              [BACCNTNO] char(16), " & _
                       " [GROSS_PAY] decimal(18,4),           [NET_PAY] decimal(18,4),            [SA_NET_PAY] decimal(18,4),          [FULLNAME] char(100), " & _
                       " [PAYSTATUS] integer,                 [FIRSTNAME] char(25),               [MNAME] char(25),                    [LASTNAME] char(25), " & _
                       " [WAP] integer,                       [date_res] date,                    [DATE_HIRE] date,                    [SEQ_NO] integer, " & _
                       " [COLA] decimal(18,4),                [SUN_COLA] decimal(18,4), " & _
                       " [SUN_ND] decimal(18,4),              [SUN_ND_PAY] decimal(18,4),         [SUN_ND_OT] decimal(18,4),           [SUN_ND_OT_PAY] decimal(18,4), " & _
                       " [date_start] date,                   [date_end] date,                    [DEPNAME] char(100) ) "

        Case 5
        
            cTblName = "tmpDaiLCost"    ' -------------------------------------------------------->table for actual daily
            cSqlStmt = " CREATE TABLE " & cTblName & "(" & _
                       " [EMPID] char(6),                     [ACTIVE] integer,                   [EMP_STAT] integer,                  [RATE_AMT] decimal(18,4), " & _
                       " [COLA_AMT] decimal(18,4),            [POS_ALLOW] decimal(18,4),          [REG_DAY] decimal(18,4),             [REG_PAY] decimal(18,4), " & _
                       " [REG_OT_HR] decimal(18,4),           [REG_OT_PAY] decimal(18,4),         [NDIFF_DAY] decimal(18,4),           [NDIFF_PAY] decimal(18,4), " & _
                       " [NDIFF_OT_HR] decimal(18,4),         [NDIFF_OT_PAY] decimal(18,4),       [HOLIDAY] decimal(18,4),             [HOL_PAY] decimal(18,4), " & _
                       " [SA_REG_OT] decimal(18,4),           [SA_REG_PAY] decimal(18,4),         [SA_NDIFF_OT] decimal(18,4),         [SA_NDIFF_PAY] decimal(18,4), " & _
                       " [SUN_HR] decimal(18,4),              [SUN_PAY] decimal(18,4),            [SUN_OT] decimal(18,4),              [SUN_OT_PAY] decimal(18,4), " & _
                       " [ADJ_PAY] decimal(18,4),             [SA_ADJ_PAY] decimal(18,4),         [OTHER_PAY] decimal(18,4),           [LEAVE_PAY] decimal(18,4), " & _
                       " [M13PAY] decimal(18,4),              [BACCNTNO] char(16), " & _
                       " [GROSS_PAY] decimal(18,4),           [NET_PAY] decimal(18,4),            [SA_NET_PAY] decimal(18,4),          [FULLNAME] char(100), " & _
                       " [PAYSTATUS] integer,                 [FIRSTNAME] char(25),               [MNAME] char(25),                    [LASTNAME] char(25), " & _
                       " [WAP] integer,                       [date_res] date,                    [DATE_HIRE] date,                    [SEQ_NO] integer, " & _
                       " [COLA] decimal(18,4),                [SUN_COLA] decimal(18,4), " & _
                       " [SUN_ND] decimal(18,4),              [SUN_ND_PAY] decimal(18,4),         [SUN_ND_OT] decimal(18,4),           [SUN_ND_OT_PAY] decimal(18,4), " & _
                       " [date_start] date,                   [date_end] date,                    [DEPNAME] char(100)) "
    
    
         Case 6 '--> Weekly Consumption Report (201610-27)
        
            cTblName = "tmpWklyCon"
            
            cSqlStmt = " Create Table " & cTblName & "(" & _
                       " [EMPID] char(6),           [FULLNAME] char(100), " & _
                       " [DEPNAME] char(100),       [POSNAME] char(100), " & _
                       " [REG_HR] double,           [REG_OT_HR] double, " & _
                       " [SA_REG_OT] double,        [ND_HR] double, " & _
                       " [ND_OT_HR] double,         [SA_ND_OT] double, " & _
                       " [SUN_HR] double,           [SUN_OT_HR] double, " & _
                       " [SUN_ND] double,           [SUN_ND_OT] double, " & _
                       " [TOT_HR] double,           [EMPSTAT] char(30), " & _
                       " [PAYSTATUS] char(30),       [ACTIVE_STAT] char(30), " & _
                       " [DATE_FIN] char(10))"
                           
    
    
    End Select

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM " & cTblName
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Sub GenLateAbsent(ByVal nMode As Integer, ByVal cDepid As String)
    Dim cSqlStmt As String, _
        cParam As String, cParam2 As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset
    
    CreateTemp 0
    
    ShowProgress 0
        
    If Check1.Value = vbChecked Then
        cParam = "(a.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ")"
    Else
        cParam = "(a.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")"
    End If
    
    If Trim(cDepid) <> "" Then cDepid = " and (d.depid in " & cDepid & ")"
    
    Select Case nMode
        Case 0
            cParam2 = " and (b.trantype=0) and (time(concat(b.logdate,' ',b.trantime)) > time(date_add(concat(b.logdate,' ',a.time1),INTERVAL a.allowance MINUTE))) "
        Case 1
            cParam2 = " and a.tag = 1 "
        Case 2
            cParam2 = " and a.tag = 3 "
    End Select
                   
    cSqlStmt = "select a.empid, a.date, a.shiftid, a.time1, a.time2, " & _
               "  ifnull(d.active,0) as active, ifnull(if(d.active=1,d.date_res,d.date_fin),'') as date_fc, " & _
               "  ifnull(concat(d.lastname,', ',d.firstname,if(trim(d.mname)='','',concat(' ',left(d.mname,1),'.'))),'') as fullname, " & _
               "  ifnull(b.trantime,'') as trantime, ifnull(b.logdate,'') as logdate, " & _
               "  ifnull(d.depid,'') as depid, ifnull(e.linename,'') as department, " & _
               "  ifnull(b.trantype,'') as trantype " & _
               "from dih36770 a " & _
               "  left join di3670 d on a.empid=d.empid " & _
               "  left join di5463 e on d.depid=e.lineid " & _
               "  left join pah84650 b on a.empid=b.empid and a.date=b.logdate " & _
               "where " & cParam & cDepid & _
               cParam2 & _
               "Union All " & _
               "select a.empid, a.date, a.shiftid, a.time1, a.time2, " & _
               "  ifnull(d.active,0) as active, ifnull(if(d.active=1,d.date_res,d.date_fin),'') as date_fc, " & _
               "  ifnull(concat(d.lastname,', ',d.firstname,if(trim(d.mname)='','',concat(' ',left(d.mname,1),'.'))),'') as fullname, " & _
               "  ifnull(b.trantime,'') as trantime, ifnull(b.logdate,'') as logdate, " & _
               "  ifnull(d.depid,'') as depid, ifnull(e.linename,'') as department, " & _
               "  ifnull(b.trantype,'') as trantype " & _
               "from di36770 a " & _
               "  left join di3670 d on a.empid=d.empid " & _
               "  left join di5463 e on d.depid=e.lineid " & _
               "  left join pa84650 b on a.empid=b.empid and a.date=b.logdate " & _
               "where " & cParam & cDepid & _
               cParam2
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
'            ' --> here me natapos, at dpat may magawa d2...
            If oRecordSet("active") > 0 Then
                If (Format(oRecordSet("date_fc"), "yyyy-mm-dd") < Format(oRecordSet("date"), "yyyy-mm-dd")) Then GoTo loopd2
            End If
            
            If oRecordSet("trantype") <> "" Then
        
            
                cSqlStmt = "insert into tmpLateAbsent(empid,fullname,depid,department,[date],time1,time2,shiftid,trantime,logdate,trantype)values(" & _
                           cQuote & oRecordSet("empid") & cQuote & "," & _
                           cQuote & UCase(oRecordSet("fullname")) & cQuote & "," & _
                           cQuote & oRecordSet("depid") & cQuote & "," & _
                           cQuote & oRecordSet("department") & cQuote & "," & _
                           cQuote & Format(oRecordSet("date"), "mm/dd/yyyy") & cQuote & "," & _
                           cQuote & Format(oRecordSet("time1"), "hh:mm AM/PM") & cQuote & "," & _
                           cQuote & Format(oRecordSet("time2"), "hh:mm AM/PM") & cQuote & "," & _
                           cQuote & oRecordSet("shiftid") & cQuote & "," & _
                           cQuote & Format(IIf(Trim(oRecordSet("trantime")) = "", Now, oRecordSet("trantime")), "hh:mm:ss AM/PM") & cQuote & "," & _
                           cQuote & Format(IIf(Trim(oRecordSet("logdate")) = "", oRecordSet("date"), oRecordSet("logdate")), "mm/dd/yyyy") & cQuote & "," & _
                           oRecordSet("trantype") & ")"
            Else
            
                cSqlStmt = "insert into tmpLateAbsent(empid,fullname,depid,department,[date],time1,time2,shiftid,trantime,logdate,trantype)values(" & _
                           cQuote & oRecordSet("empid") & cQuote & "," & _
                           cQuote & UCase(oRecordSet("fullname")) & cQuote & "," & _
                           cQuote & oRecordSet("depid") & cQuote & "," & _
                           cQuote & oRecordSet("department") & cQuote & "," & _
                           cQuote & Format(oRecordSet("date"), "mm/dd/yyyy") & cQuote & "," & _
                           cQuote & Format(oRecordSet("time1"), "hh:mm AM/PM") & cQuote & "," & _
                           cQuote & Format(oRecordSet("time2"), "hh:mm AM/PM") & cQuote & "," & _
                           cQuote & oRecordSet("shiftid") & cQuote & "," & _
                           cQuote & Format(IIf(Trim(oRecordSet("trantime")) = "", Now, oRecordSet("trantime")), "hh:mm:ss AM/PM") & cQuote & "," & _
                           cQuote & Format(IIf(Trim(oRecordSet("logdate")) = "", oRecordSet("date"), oRecordSet("logdate")), "mm/dd/yyyy") & cQuote & ")"
             
            End If
            
            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
loopd2:
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        Select Case nMode
            Case 0
                GenerateReport "Late Report " & IIf(Check1.Value = vbChecked, " for the period " & Format(DTPicker1.Value, "mmm d, yyyy") & " to " & Format(DTPicker2.Value, "mmm d, yyyy"), " for " & Format(DTPicker1.Value, "mmm d, yyyy")), "rpt5283.rpt"
            Case 1
                GenerateReport "Absent Report " & IIf(Check1.Value = vbChecked, " for the period " & Format(DTPicker1.Value, "mmm d, yyyy") & " to " & Format(DTPicker2.Value, "mmm d, yyyy"), " for " & Format(DTPicker1.Value, "mmm d, yyyy")), "rpt5283A.rpt"
            Case 2
                GenerateReport "Incomplete Entry Report " & IIf(Check1.Value = vbChecked, " for the period " & Format(DTPicker1.Value, "mmm d, yyyy") & " to " & Format(DTPicker2.Value, "mmm d, yyyy"), " for " & Format(DTPicker1.Value, "mmm d, yyyy")), "rpt5283I.rpt"
        End Select
        
        ShowProgress 4
        
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub
Sub GenWklyCon(ByVal cDepid As String)
        Dim cSqlStmt As String, _
        cParam As String, _
        cCmpname As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset
        Dim cDateNow As Date
        cDateNow = DateValue(Now)
        
    CreateTemp 6

    ShowProgress 0
    
    '----> (201703-22) Updated (Renz)
If DateDiff("d", cDateNow, DTPicker1.Value) <= -30 Then
  If Check1.Value <> vbChecked Then

             cParam = " where ( " & _
                      " ((b.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_fin = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and a.active=2)) " & _
                      " or    ((b.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and ((a.active=1)or (a.active=3))) " & _
                      " or    ((b.date= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and ((a.date_hire<= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") )) "


               
        Else
   
             cParam = " where (((b.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ")  " & _
                      " and   ((a.date_fin between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") and (a.active=2))" & _
                      " or    ((b.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") " & _
                      " and (a.date_res between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "))  " & _
                      " and ((a.active=1)or (a.active=3)))) " & _
                      " or  ((b.date  between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_hire<= " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "))"
        End If
    

Else
      If Check1.Value <> vbChecked Then

             cParam = " where ( " & _
                      " ((b.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_fin = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and a.active=2)) " & _
                      " or    ((b.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_res = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and ((a.active=1)or (a.active=3))) " & _
                      " or    ((b.date= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and ((a.date_hire<= " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (a.active = 0))) "


               
        Else
   
             cParam = " where (((b.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ")  " & _
                      " and   ((a.date_fin between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") and (a.active=2))" & _
                      " or    ((b.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") " & _
                      " and (a.date_res between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "))  " & _
                      " and ((a.active=1)or (a.active=3)))) " & _
                      " or  ((b.date  between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") and (a.date_hire<= " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") and (a.active = 0))"
        End If
    
End If
 
    
    '----> (201703-22) Updated (Renz)
If DateDiff("d", cDateNow, DTPicker1.Value) <= -30 Then
   
   If Trim(cDepid) <> "" Then cDepid = " and (b.depid in " & cDepid & ")"
    
    
    cSqlStmt = " Select a.empid, b.date, concat(a.lastname, ', ',  a.firstname, ' ', left(a.mname,1), '.') as fullname, ifnull(c.LINENAME,'') as LINENAME, ifnull(d.POSNAME,'') as POSNAME, " & _
                      " round(ifnull(sum(b.reg_hr),0),3) as REG_HR, round(ifnull(sum(b.reg_ot_hr),0),3) as REG_OT_HR, round(ifnull(sum(b.sa_reg_ot),0),3) as SA_REG_OT, " & _
                      " round(ifnull(sum(b.nd_hr),0),3) as ND_HR, round(ifnull(sum(b.nd_ot_hr),0),3) as nd_ot_hr , round(ifnull(sum(b.sa_nd_ot),0),3) as sa_nd_ot, round(ifnull(sum(b.sun_hr),0),3) as sun_hr, round(ifnull(sum(b.sun_ot_hr),0),3) as sun_ot_hr, " & _
                      " round(ifnull(sum(b.sun_nd),0),3) as sun_nd, round(ifnull(sum(b.sun_nd_ot),0),3) as sun_nd_ot," & _
                      " round(ifnull(sum(b.reg_hr + b.reg_ot_hr + b.sa_reg_ot + b.nd_hr + b.nd_ot_hr + b.sa_nd_ot + b.sun_hr + b.sun_ot_hr + b.sun_nd + b.sun_nd_ot),0),3) as TOT_HR, " & _
                      " a.EMP_STAT, a.PAYSTATUS, a.ACTIVE , if(a.active = 0, '',if((a.active=1) or (a.active=3), a.date_res, a.date_fin)) as date_fin " & _
                      " from di3670 a left join dih36770 b on a.empid = b.empid left join  di5463 c on a.depid = c.lineid left join  di7670 d on a.posid=d.posid " & _
                      cParam & _
                      " group by a.empid " & _
                      " order by c.linename, d.posname, a.emp_stat, a.paystatus, a.lastname, a.firstname, a.mname "
                      
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
   
   
Else
   
   If Trim(cDepid) <> "" Then cDepid = " and (b.depid in " & cDepid & ")"
    
    
    cSqlStmt = " Select a.empid, b.date, concat(a.lastname, ', ',  a.firstname, ' ', left(a.mname,1), '.') as fullname, ifnull(c.LINENAME,'') as LINENAME, ifnull(d.POSNAME,'') as POSNAME, " & _
                      " round(ifnull(sum(b.reg_hr),0),3) as REG_HR, round(ifnull(sum(b.reg_ot_hr),0),3) as REG_OT_HR, round(ifnull(sum(b.sa_reg_ot),0),3) as SA_REG_OT, " & _
                      " round(ifnull(sum(b.nd_hr),0),3) as ND_HR, round(ifnull(sum(b.nd_ot_hr),0),3) as nd_ot_hr , round(ifnull(sum(b.sa_nd_ot),0),3) as sa_nd_ot, round(ifnull(sum(b.sun_hr),0),3) as sun_hr, round(ifnull(sum(b.sun_ot_hr),0),3) as sun_ot_hr, " & _
                      " round(ifnull(sum(b.sun_nd),0),3) as sun_nd, round(ifnull(sum(b.sun_nd_ot),0),3) as sun_nd_ot," & _
                      " round(ifnull(sum(b.reg_hr + b.reg_ot_hr + b.sa_reg_ot + b.nd_hr + b.nd_ot_hr + b.sa_nd_ot + b.sun_hr + b.sun_ot_hr + b.sun_nd + b.sun_nd_ot),0),3) as TOT_HR, " & _
                      " a.EMP_STAT, a.PAYSTATUS, a.ACTIVE , if(a.active = 0, '',if((a.active=1) or (a.active=3), a.date_res, a.date_fin)) as date_fin " & _
                      " from di3670 a left join di36770 b on a.empid = b.empid left join  di5463 c on a.depid = c.lineid left join  di7670 d on a.posid=d.posid " & _
                      cParam & _
                      " group by a.empid " & _
                      " order by c.linename, d.posname, a.emp_stat, a.paystatus, a.lastname, a.firstname, a.mname "
                      
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    
    
End If


    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
        cSqlStmt = "insert into tmpWklyCon (EMPID, FULLNAME,DEPNAME,POSNAME, REG_HR, REG_OT_HR, SA_REG_OT,ND_HR, ND_OT_HR, SA_ND_OT, SUN_HR, SUN_OT_HR, SUN_ND, SUN_ND_OT, TOT_HR, EMPSTAT, PAYSTATUS, ACTIVE_STAT, DATE_FIN) values ( " & _
               cQuote & oRecordSet("empid") & cQuote & "," & _
               cQuote & oRecordSet("fullname") & cQuote & "," & _
               cQuote & oRecordSet("linename") & cQuote & "," & _
               cQuote & oRecordSet("posname") & cQuote & "," & _
               cQuote & oRecordSet("reg_hr") & cQuote & "," & _
               cQuote & oRecordSet("reg_ot_hr") & cQuote & "," & _
               cQuote & oRecordSet("sa_reg_ot") & cQuote & "," & _
               cQuote & oRecordSet("nd_hr") & cQuote & "," & _
               cQuote & oRecordSet("nd_ot_hr") & cQuote & "," & _
               cQuote & oRecordSet("sa_nd_ot") & cQuote & "," & _
               cQuote & oRecordSet("sun_hr") & cQuote & "," & _
               cQuote & oRecordSet("sun_ot_hr") & cQuote & "," & _
               cQuote & oRecordSet("sun_nd") & cQuote & "," & _
               cQuote & oRecordSet("sun_nd_ot") & cQuote & "," & _
               cQuote & oRecordSet("tot_hr") & cQuote & "," & _
               cQuote & IIf(oRecordSet("emp_stat") = 0, "Wap", IIf(oRecordSet("emp_stat") = 1, "Contractual", "Regular")) & cQuote & "," & _
               cQuote & IIf(oRecordSet("paystatus") = 0, "Daily", IIf(oRecordSet("paystatus") = 1, "Monthly", "Emergency")) & cQuote & "," & _
               cQuote & IIf(oRecordSet("Active") = 0, "Active", IIf(oRecordSet("Active") = 1, "Resigned", IIf(oRecordSet("Active") = 2, "Finished", "Terminated"))) & cQuote & "," & _
               cQuote & oRecordSet("date_fin") & cQuote & ")"
               
        QueryTemp cSqlStmt, objdbRs, True
        
        oRecordSet.MoveNext
      Wend
      
        ShowProgress 3
        
        If Check1.Value <> vbChecked Then
        
        GenerateReport "Weekly Consumption Report for " & Format(DTPicker1.Value, "mmm d, yyyy"), "rpt9559266D.RPT"
        Else
        
        GenerateReport "Weekly Consumption Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "rpt9559266D.RPT"
        
        End If
        
        ShowProgress 4
        
    Else
        ShowProgress 4
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
            
End Sub



Sub GenEmpOt(ByVal cDepid As String)
    Dim cSqlStmt As String, _
        cParam As String, cParam2 As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        cShiftDesc As String, _
        aTrantype As Variant, _
        aTimeInfo As Variant

    CreateTemp 1
    ShowProgress 0
    If Check1.Value <> vbChecked Then
        cParam = " where a.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote
    Else
        cParam = " where a.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote
    End If
    
    If Trim(cDepid) <> "" Then cDepid = " and (b.depid in " & cDepid & ")"
    
    cSqlStmt = " SELECT a.EMPID, concat(b.LASTNAME,', ', b.FIRSTNAME,' ', left(b.MNAME,1),'.') as fullname,  a.DATE, a.SHIFTID, a.REMARK, a.TIME1, a.TIME2, a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.tot_ot, a.nd_hr, a.nd_ot_hr, a.sa_nd_ot, a.nd_tot_ot, a.sun_hr, a.sun_ot_hr, a.sun_nd, a.sun_nd_ot, " & _
               " b.DEPID , b.TCID, b.BCID, b.POSID, b.EMP_STAT, b.ACTIVE, b.PAYSTATUS, " & _
               " ifnull(d.posname,'') as posname, " & _
               " ifnull(c.linename,'') as depname " & _
               " FROM di36770 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " left join di7670 d on b.posid=d.posid " & _
               cParam & _
               " and b.active = 0 " & _
               cDepid & " order by a.empid,a.date "
    OpenQueryDNS cSqlStmt, oRecordSet, False
    
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            OpenQueryDNS " select * from pa74380 where shiftid = " & cQuote & oRecordSet("shiftid") & cQuote, objdbRs, False
            cShiftDesc = IIf(objdbRs.RecordCount > 0, objdbRs("description"), "")
            
            cSqlStmt = "select tran_no, " & _
                       "       transdate, " & _
                       "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                       "       trantype, " & _
                       "       if(trantype=0,'In','Out') as trn_type, " & _
                       "       trantime " & _
                       " from pa84650 " & _
                       " where empid=" & cQuote & oRecordSet("empid") & cQuote & _
                       "   and logdate=" & cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & _
                       " order by transdate, trantime"
            OpenQueryDNS cSqlStmt, oTempADO, False
            If oTempADO.RecordCount > 0 Then
                aTrantype = Array("", "", "", "")
                While Not oTempADO.EOF
                    aTrantype(3) = oTempADO("TRANSDATE")
                    If oTempADO("trantype") = 0 Then
                        If Trim(aTrantype(1)) = "" Then
                            aTrantype(0) = oTempADO("trantype")
                            aTrantype(1) = oTempADO("trantime")
                        End If
                    Else
                        aTrantype(0) = oTempADO("trantype")
                        aTrantype(2) = oTempADO("trantime")
                    End If

                    oTempADO.MoveNext

                    If Not oTempADO.EOF Then
                        If (oTempADO("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                            cSqlStmt = " insert into tmpEmpOt(EMPID,FULLNAME,TCID,POSNAME,EMP_STAT,ACTIVE,PAYSTATUS,[DATE],SHIFTDESC,TIME1,TIME2,reg_hr,reg_ot_hr, " & _
                                       " sa_reg_ot,tot_ot,nd_hr,nd_ot_hr,sa_nd_ot,nd_tot_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot,INTRANTIME,OUTTRANTIME,DEPNAME)values(" & _
                                       cQuote & oRecordSet("empid") & cQuote & "," & _
                                       cQuote & oRecordSet("fullname") & cQuote & "," & _
                                       cQuote & oRecordSet("tcid") & cQuote & "," & _
                                       cQuote & oRecordSet("posname") & cQuote & "," & _
                                       oRecordSet("emp_stat") & "," & _
                                       oRecordSet("active") & "," & _
                                       oRecordSet("paystatus") & "," & _
                                       cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & cShiftDesc & cQuote & "," & _
                                       cQuote & oRecordSet("time1") & cQuote & "," & _
                                       cQuote & oRecordSet("time2") & cQuote & "," & _
                                       oRecordSet("reg_hr") & "," & oRecordSet("reg_ot_hr") & "," & _
                                       oRecordSet("sa_reg_ot") & "," & oRecordSet("tot_ot") & "," & oRecordSet("nd_hr") & "," & _
                                       oRecordSet("nd_ot_hr") & "," & oRecordSet("sa_nd_ot") & "," & oRecordSet("nd_tot_ot") & "," & _
                                       oRecordSet("sun_hr") & "," & oRecordSet("sun_ot_hr") & "," & _
                                       oRecordSet("sun_nd") & "," & oRecordSet("sun_nd_ot") & "," & _
                                       cQuote & aTrantype(1) & cQuote & "," & cQuote & aTrantype(2) & cQuote & "," & _
                                       cQuote & oRecordSet("depname") & cQuote & ")"
                                       
                            QueryTemp cSqlStmt, objdbRs, True
                            
                            aTrantype = Array("", "", "", "")
                        End If
                    Else
                        cSqlStmt = " insert into tmpEmpOt(EMPID,FULLNAME,TCID,POSNAME,EMP_STAT,ACTIVE,PAYSTATUS,[DATE],SHIFTDESC,TIME1,TIME2,reg_hr,reg_ot_hr, " & _
                                   " sa_reg_ot,tot_ot,nd_hr,nd_ot_hr,sa_nd_ot,nd_tot_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot,INTRANTIME,OUTTRANTIME,DEPNAME)values(" & _
                                   cQuote & oRecordSet("empid") & cQuote & "," & _
                                   cQuote & oRecordSet("fullname") & cQuote & "," & _
                                   cQuote & oRecordSet("tcid") & cQuote & "," & _
                                   cQuote & oRecordSet("posname") & cQuote & "," & _
                                   oRecordSet("emp_stat") & "," & _
                                   oRecordSet("active") & "," & _
                                   oRecordSet("paystatus") & "," & _
                                   cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                                   cQuote & cShiftDesc & cQuote & "," & _
                                   cQuote & oRecordSet("time1") & cQuote & "," & _
                                   cQuote & oRecordSet("time2") & cQuote & "," & _
                                   oRecordSet("reg_hr") & "," & oRecordSet("reg_ot_hr") & "," & _
                                   oRecordSet("sa_reg_ot") & "," & oRecordSet("tot_ot") & "," & oRecordSet("nd_hr") & "," & _
                                   oRecordSet("nd_ot_hr") & "," & oRecordSet("sa_nd_ot") & "," & oRecordSet("nd_tot_ot") & "," & _
                                   oRecordSet("sun_hr") & "," & oRecordSet("sun_ot_hr") & "," & _
                                   oRecordSet("sun_nd") & "," & oRecordSet("sun_nd_ot") & "," & _
                                   cQuote & aTrantype(1) & cQuote & "," & cQuote & aTrantype(2) & cQuote & "," & _
                                   cQuote & oRecordSet("depname") & cQuote & ")"
                        QueryTemp cSqlStmt, objdbRs, True
                    End If
                Wend
            Else
                cSqlStmt = " insert into tmpEmpOt(EMPID,FULLNAME,TCID,POSNAME,EMP_STAT,ACTIVE,PAYSTATUS,[DATE],SHIFTDESC,TIME1,TIME2,reg_hr,reg_ot_hr, " & _
                           " sa_reg_ot,tot_ot,nd_hr,nd_ot_hr,sa_nd_ot,nd_tot_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot,DEPNAME)values(" & _
                           cQuote & oRecordSet("empid") & cQuote & "," & _
                           cQuote & oRecordSet("fullname") & cQuote & "," & _
                           cQuote & oRecordSet("tcid") & cQuote & "," & _
                           cQuote & oRecordSet("posname") & cQuote & "," & _
                           oRecordSet("emp_stat") & "," & _
                           oRecordSet("active") & "," & _
                           oRecordSet("paystatus") & "," & _
                           cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & cShiftDesc & cQuote & "," & _
                           cQuote & oRecordSet("time1") & cQuote & "," & _
                           cQuote & oRecordSet("time2") & cQuote & "," & _
                           oRecordSet("reg_hr") & "," & oRecordSet("reg_ot_hr") & "," & _
                           oRecordSet("sa_reg_ot") & "," & oRecordSet("tot_ot") & "," & oRecordSet("nd_hr") & "," & _
                           oRecordSet("nd_ot_hr") & "," & oRecordSet("sa_nd_ot") & "," & oRecordSet("nd_tot_ot") & "," & _
                           oRecordSet("sun_hr") & "," & oRecordSet("sun_ot_hr") & "," & _
                           oRecordSet("sun_nd") & "," & oRecordSet("sun_nd_ot") & "," & _
                           cQuote & oRecordSet("depname") & cQuote & ")"
                QueryTemp cSqlStmt, objdbRs, True
            End If

            oRecordSet.MoveNext
        Wend
        ShowProgress 3

        GenerateReport "Employee OT Report", "rpt36760.rpt"

        ShowProgress 4
    Else
        ShowProgress 4
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub

Sub GenDTCons(ByVal cDepid As String)
'    Dim cSqlStmt As String, _
'        cParam As String, cParam2 As String, _
'        nCtr As Integer, _
'        oRecordSet As New ADODB.Recordset, _
'        oRSet As New ADODB.Recordset, _
'        nManPower, nRegHour1, nRegHour0, _
'        nOtHour0, nOtHour1, nOtHour2, nOtHour3, nOtHour4, nOtHour5, nOtHour6, _
'        nOtHour7, nOtHour8, nOtHour9, nOtHour10, nOtHour11, nOtHour12, nOtHour13, _
'        nOtHour14, nOtHour15, nOtHour16, _
'        nTotOt, nAbsent As Double, _
'        aTimeInfo As Variant, aStatTot As Variant, _
'        nPClose As Integer
        
    Dim cSqlStmt As String, _
        cParam As String, cParam2 As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        nManPower, nRegHour1, nRegHour0, _
        nOtHour0, nOtHour0_5, nOtHour1, nOtHour1_5, nOtHour2, nOtHour2_5, nOtHour3, nOtHour3_5, nOtHour4, nOtHour4_5, nOtHour5, nOtHour5_5, nOtHour6, nOtHour6_5, nOtHour7, nOtHour7_5, nOtHour8, nOtHour8_5, nOtHour9, nOtHour9_5, nOtHour10, nOtHour10_5, nOtHour11, nOtHour11_5, nOtHour12, nOtHour12_5, nOtHour13, nOtHour13_5, nOtHour14, nOtHour14_5, nOtHour15, nOtHour15_5, nOtHour16, nOtHour16_5, nOtHour17, _
        nTotOt, nAbsent As Double, _
        aTimeInfo As Variant, aStatTot As Variant, _
        nPClose As Integer

    CreateTemp 2
    ShowProgress 0
    
    If Trim(cDepid) <> "" Then cDepid = " where (lineid in " & cDepid & ")"
    
    cSqlStmt = "select lineid,linename from di5463 " & cDepid
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
                        
            nManPower = 0
            nRegHour1 = 0
            nRegHour0 = 0
            nOtHour0 = 0
            nOtHour0_5 = 0
            nOtHour1 = 0
            nOtHour1_5 = 0
            nOtHour2 = 0
            nOtHour2_5 = 0
            nOtHour3 = 0
            nOtHour3_5 = 0
            nOtHour4 = 0
            nOtHour4_5 = 0
            nOtHour5 = 0
            nOtHour5_5 = 0
            nOtHour6 = 0
            nOtHour6_5 = 0
            nOtHour7 = 0
            nOtHour7_5 = 0
            nOtHour8 = 0
            nOtHour8_5 = 0
            nOtHour9 = 0
            nOtHour9_5 = 0
            nOtHour10 = 0
            nOtHour10_5 = 0
            nOtHour11 = 0
            nOtHour11_5 = 0
            nOtHour12 = 0
            nOtHour12_5 = 0
            nOtHour13 = 0
            nOtHour13_5 = 0
            nOtHour14 = 0
            nOtHour14_5 = 0
            nOtHour15 = 0
            nOtHour15_5 = 0
            nOtHour16 = 0
            nOtHour16_5 = 0
            nOtHour17 = 0
            nTotOt = 0
            nAbsent = 0
            
            

            
            cSqlStmt = "select pclose,periodid from pa7730 " & _
                        "where " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " between date_start and date_end"
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

            nPClose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), 0)
            
            cSqlStmt = " SELECT b.tcid, a.EMPID, a.DATE, b.FIRSTNAME, b.MNAME, b.LASTNAME, concat(b.LASTNAME,', ',b.FIRSTNAME,' ',left(b.MNAME,1),'. ') as fullname, " & _
                        " b.DEPID , c.LINENAME, b.emp_stat, b.wap, b.paystatus,b.active,ifnull(e.posname,'') as posname  " & _
                        " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & " a " & _
                        " left join di3670 b on a.empid=b.empid  " & _
                        " left join di5463 c on b.depid=c.lineid " & _
                        " left join di7670 e on b.posid=e.posid " & _
                        " where  a.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                        " and b.depid = " & cQuote & oRecordSet("lineid") & cQuote & _
                        " and b.active=0 "
                        
                       
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, oRSet, False
            If oRSet.RecordCount > 0 Then
                While Not oRSet.EOF
                    nManPower = 0
                    nRegHour1 = 0
                    nRegHour0 = 0
                    nOtHour0 = 0
                    nOtHour0_5 = 0
                    nOtHour1 = 0
                    nOtHour1_5 = 0
                    nOtHour2 = 0
                    nOtHour2_5 = 0
                    nOtHour3 = 0
                    nOtHour3_5 = 0
                    nOtHour4 = 0
                    nOtHour4_5 = 0
                    nOtHour5 = 0
                    nOtHour5_5 = 0
                    nOtHour6 = 0
                    nOtHour6_5 = 0
                    nOtHour7 = 0
                    nOtHour7_5 = 0
                    nOtHour8 = 0
                    nOtHour8_5 = 0
                    nOtHour9 = 0
                    nOtHour9_5 = 0
                    nOtHour10 = 0
                    nOtHour10_5 = 0
                    nOtHour11 = 0
                    nOtHour11_5 = 0
                    nOtHour12 = 0
                    nOtHour12_5 = 0
                    nOtHour13 = 0
                    nOtHour13_5 = 0
                    nOtHour14 = 0
                    nOtHour14_5 = 0
                    nOtHour15 = 0
                    nOtHour15_5 = 0
                    nOtHour16 = 0
                    nOtHour16_5 = 0
                    nOtHour17 = 0
                    nTotOt = 0
                    nAbsent = 0
                    
'                    aStatTot(0) = [ST_R]
'                    aStatTot(1) = [ST_R_M]
'                    aStatTot(2) = [ST_C]
'                    aStatTot(3) = [ST_C_M]
'                    aStatTot(4) = [ST_W]
'                    aStatTot(5) = [ST_W_C]
'                    aStatTot(6) = [ST_C_E]
'                    aStatTot(7) = [ST_W_E]
                    
                    aStatTot = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                
'                    If oRSet("depid") = "043" Then Stop
                    ' --> retrieve computed dtr here... 20080526
                    aTimeInfo = ComputeDays(oRSet("empid"), _
                                            Array(Format(DTPicker1.Value, "yyyy-mm-dd"), Format(DTPicker1.Value, "yyyy-mm-dd"), 0), _
                                            Array(oRSet("emp_stat"), oRSet("wap"), oRSet("paystatus")), IIf(nPClose = 0, False, True))
                    
                    If ((aTimeInfo(0) * 8) <> 0) Or ((aTimeInfo(3) * 8) <> 0) Or ((aTimeInfo(5) * 8) <> 0) Or ((aTimeInfo(13) * 8) <> 0) Then
                        If ((aTimeInfo(0) * 8) < 8) And ((aTimeInfo(3) * 8) < 8) And ((aTimeInfo(5) * 8) < 8) And ((aTimeInfo(13) * 8) < 8) Then
                            nRegHour0 = 1
                        Else
                            nRegHour1 = 1
                                                                                
                        End If
                        nTotOt = aTimeInfo(1) + aTimeInfo(2) + aTimeInfo(4) + _
                                 aTimeInfo(12) + aTimeInfo(6) + aTimeInfo(14)
                        Select Case nTotOt
                            Case Is < 0.5
                                nOtHour0 = 1
                            Case Is < 1
                                nOtHour0_5 = 1
                            Case Is < 1.5
                                nOtHour1 = 1
                            Case Is < 2
                                nOtHour1_5 = 1
                            Case Is < 2.5
                                nOtHour2 = 1
                            Case Is < 3
                                nOtHour2_5 = 1
                            Case Is < 3.5
                                nOtHour3 = 1
                            Case Is < 4
                                nOtHour3_5 = 1
                            Case Is < 4.5
                                nOtHour4 = 1
                            Case Is < 5
                                nOtHour4_5 = 1
                            Case Is < 5.5
                                nOtHour5 = 1
                            Case Is < 6
                                nOtHour5_5 = 1
                            Case Is < 6.5
                                nOtHour6 = 1
                            Case Is < 7
                                nOtHour6_5 = 1
                            Case Is < 7.5
                                nOtHour7 = 1
                            Case Is < 8
                                nOtHour7_5 = 1
                            Case Is < 8.5
                                nOtHour8 = 1
                            Case Is < 9
                                nOtHour8_5 = 1
                            Case Is < 9.5
                                nOtHour9 = 1
                            Case Is < 10
                                nOtHour9_5 = 1
                            Case Is < 10.5
                                nOtHour10 = 1
                            Case Is < 11
                                nOtHour10_5 = 1
                            Case Is < 11.5
                                nOtHour11 = 1
                            Case Is < 12
                                nOtHour11_5 = 1
                            Case Is < 12.5
                                nOtHour12 = 1
                            Case Is < 13
                                nOtHour12_5 = 1
                            Case Is < 13.5
                                nOtHour13 = 1
                            Case Is < 14
                                nOtHour13_5 = 1
                            Case Is < 14.5
                                nOtHour14 = 1
                            Case Is < 15
                                nOtHour14_5 = 1
                            Case Is < 15.5
                                nOtHour15 = 1
                            Case Is < 16
                                nOtHour15_5 = 1
                            Case Is < 16.5
                                nOtHour16 = 1
                            Case Is < 17
                                nOtHour16_5 = 1
                            Case Is < 17.5
                                nOtHour17 = 1
                            Case Is >= 17
                                nOtHour17 = 1
                        End Select
                        
                    Else
                        nAbsent = 1
                    End If
                    
                    If ((aTimeInfo(0) * 8) <> 0) Or ((aTimeInfo(3) * 8) <> 0) Or ((aTimeInfo(5) * 8) <> 0) Or ((aTimeInfo(13) * 8) <> 0) Then
                        If oRSet("active") = 0 Then
                            nManPower = 1
                        End If
                    End If
                                        
                    Select Case oRSet("emp_stat")
                        Case Is = 0 'wap
                            If oRSet("paystatus") = 2 Then
                                aStatTot(7) = 1
                            Else
                                aStatTot(4) = 1
                            End If
                            
                        Case Is = 1 'contractual
                            If oRSet("paystatus") = 2 Then
                                aStatTot(6) = 1
                            Else
                                If oRSet("paystatus") = 1 Then
                                    aStatTot(3) = 1
                                Else
                                    If oRSet("wap") = 1 Then
                                        aStatTot(5) = 1
                                    Else
                                        aStatTot(2) = 1
                                    End If
                                End If
                            End If
                        Case Is = 2 'regular
                            If oRSet("paystatus") = 1 Then
                                aStatTot(1) = 1
                            Else
                                aStatTot(0) = 1
                            End If
                        
            
                        
                    End Select
                    
'                    cSqlStmt = " insert into tmpDTCons([date],DEPNAME,Manpower,reg_hr,reg_hr1, " & _
'                               " hour0 , hour_5, hour1, hour1_5, hour2, hour2_5, hour3, hour3_5, hour4, hour4_5, hour5, hour5_5, hour6, hour6_5, hour7, hour7_5,  hour8,  hour8_5, hour9, hour9_5, hour10, hour10_5, hour11, " & _
'                               " absent," & _
'                               " EMPID,FULLNAME,TCID,POSNAME,EMP_STAT,ACTIVE,PAYSTATUS,ST_R, ST_R_M, ST_C, ST_C_M, ST_W, ST_W_C, ST_C_E, ST_W_E)values(" & _
'                               cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
'                               cQuote & oRecordSet("linename") & cQuote & "," & _
'                               nManPower & "," & _
'                               nRegHour0 & "," & _
'                               nRegHour1 & "," & _
'                               nOtHour0 & "," & nOtHour0_5 & "," & nOtHour1 & "," & nOtHour1_5 & "," & nOtHour2 & "," & nOtHour2_5 & "," & _
'                               nOtHour3 & "," & nOtHour3_5 & "," & nOtHour4 & "," & nOtHour4_5 & "," & nOtHour5 & "," & nOtHour5_5 & "," & _
'                               nOtHour6 & "," & nOtHour6_5 & "," & nOtHour7 & "," & nOtHour7_5 & "," & nOtHour8 & "," & nOtHour8_5 & "," & _
'                               nOtHour9 & "," & nOtHour9_5 & "," & nOtHour10 & "," & nOtHour10_5 & "," & nOtHour11 & "," & nAbsent & "," & _
'                               cQuote & oRSet("EMPID") & cQuote & "," & _
'                               cQuote & oRSet("FULLNAME") & cQuote & "," & _
'                               cQuote & oRSet("TCID") & cQuote & "," & _
'                               cQuote & oRSet("POSNAME") & cQuote & "," & _
'                               cQuote & oRSet("EMP_STAT") & cQuote & "," & _
'                               cQuote & oRSet("ACTIVE") & cQuote & "," & _
'                               cQuote & oRSet("PAYSTATUS") & cQuote & "," & _
'                               aStatTot(0) & "," & aStatTot(1) & "," & aStatTot(2) & "," & _
'                               aStatTot(3) & "," & aStatTot(4) & "," & aStatTot(5) & "," & _
'                               aStatTot(6) & "," & aStatTot(7) & ")"
                               
                    cSqlStmt = " insert into tmpDTCons([date],DEPNAME,Manpower,reg_hr,reg_hr1, " & _
                               " hour0 , hour_5, hour1, hour1_5, hour2, hour2_5, hour3, hour3_5, hour4, hour4_5, hour5, hour5_5, hour6, hour6_5, hour7, hour7_5,  hour8,  hour8_5, hour9, hour9_5, hour10, hour10_5, hour11, hour11_5, hour12, hour12_5, hour13, hour13_5, hour14, hour14_5, hour15, hour15_5, hour16, hour16_5, hour17,  " & _
                               " absent," & _
                               " EMPID,FULLNAME,TCID,POSNAME,EMP_STAT,ACTIVE,PAYSTATUS,ST_R, ST_R_M, ST_C, ST_C_M, ST_W, ST_W_C, ST_C_E, ST_W_E)values(" & _
                               cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & oRecordSet("linename") & cQuote & "," & _
                               nManPower & "," & nRegHour0 & "," & nRegHour1 & "," & _
                               nOtHour0 & "," & nOtHour0_5 & "," & nOtHour1 & "," & nOtHour1_5 & "," & nOtHour2 & "," & nOtHour2_5 & "," & _
                               nOtHour3 & "," & nOtHour3_5 & "," & nOtHour4 & "," & nOtHour4_5 & "," & nOtHour5 & "," & nOtHour5_5 & "," & _
                               nOtHour6 & "," & nOtHour6_5 & "," & nOtHour7 & "," & nOtHour7_5 & "," & nOtHour8 & "," & nOtHour8_5 & "," & _
                               nOtHour9 & "," & nOtHour9_5 & "," & _
                               nOtHour10 & "," & nOtHour10_5 & "," & nOtHour11 & "," & nOtHour11_5 & "," & _
                               nOtHour12 & "," & nOtHour12_5 & "," & nOtHour13 & "," & nOtHour13_5 & "," & _
                               nOtHour14 & "," & nOtHour14_5 & "," & nOtHour15 & "," & nOtHour15_5 & "," & _
                               nOtHour16 & "," & nOtHour16_5 & "," & nOtHour17 & "," & nAbsent & "," & _
                               cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                               cQuote & oRSet("TCID") & cQuote & "," & cQuote & oRSet("POSNAME") & cQuote & "," & _
                               cQuote & oRSet("EMP_STAT") & cQuote & "," & cQuote & oRSet("ACTIVE") & cQuote & "," & _
                               cQuote & oRSet("PAYSTATUS") & cQuote & "," & _
                               aStatTot(0) & "," & aStatTot(1) & "," & aStatTot(2) & "," & _
                               aStatTot(3) & "," & aStatTot(4) & "," & aStatTot(5) & "," & _
                               aStatTot(6) & "," & aStatTot(7) & ")"
                               
                    QueryTemp cSqlStmt, objdbRs, True
            
                    
                    oRSet.MoveNext
                Wend
            End If
                    
            oRecordSet.MoveNext
        Wend

        ShowProgress 3
        If Check4.Value = vbUnchecked Then
            If Combo1.ListIndex = 0 Then
                ' 1-10.5
                GenerateReport "Employee Daily Time Consumption Summary Report", "PRV3820S.rpt"
            Else
                ' 11-16.5
                GenerateReport "Employee Daily Time Consumption Summary Report", "PRV3820S_1117.rpt"
            End If
        Else
            If Combo1.ListIndex = 0 Then
                ' 1-10.5
                GenerateReport "Employee Daily Time Consumption Summary Report", "PRV3820D.rpt"
            Else
                ' 11-16.5
                GenerateReport "Employee Daily Time Consumption Detail Report", "PRV3820D_1117.rpt"
            End If
        End If

        ShowProgress 4
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub
Sub GenActMan(ByVal nMode As Integer, ByVal cParam As String)
    Dim cSqlStmt As String, _
        cParam2 As String, _
        cDepid As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        nRegular, nContrac, nWap, nWapC, nRemark As Integer
        
    CreateTemp 3
    
    ShowProgress 0

    If Trim(cDepid) <> "" Then cDepid = " where (lineid in " & cDepid & ")"

    cSqlStmt = "select lineid,linename from di5463 " & cDepid
    OpenQueryDNS cSqlStmt, objdbRs, False
    
    cSqlStmt = " SELECT a.EMPID, a.DATE, " & _
               " concat(b.lastname,', ',b.firstname,' ',left(b.mname,1),'. ') as fullname, " & _
               " b.FIRSTNAME, b.MNAME, b.LASTNAME, " & _
               " b.DEPID , ifnull(c.LINENAME,'Undefined Department') as LINENAME, b.emp_stat, b.wap, b.paystatus,a.remark " & _
               " FROM di36770 a " & _
               " left join di3670 b on a.empid=b.empid  " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " where a.date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
               IIf(cDepid <> "", " and b.depid = " & cQuote & objdbRs("lineid") & cQuote, "")
                       
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
        
            nRegular = 0
            nContrac = 0
            nWap = 0
            nWapC = 0
            
            Select Case oRecordSet("emp_stat")
                Case 0 ' wap
                    nWap = 1
                Case 1 'contractual
                    If oRecordSet("wap") = 1 Then
                        nWapC = 1
                    Else
                        nContrac = 1
                    End If
                Case 2 ' regular
                    nRegular = 1
            End Select
            
            nRemark = IIf(Trim(UCase(oRecordSet("remark"))) = UCase("Incomplete entry"), 1, 0)

           cSqlStmt = " insert into tmpDActMan([date],DEPNAME,EMPID,FULLNAME,REG,CONT,WAP,WAPC,REMARK)values(" & _
                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & oRecordSet("linename") & cQuote & "," & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & oRecordSet("fullname") & cQuote & "," & _
                       nRegular & "," & nContrac & "," & nWap & "," & nWapC & "," & nRemark & ")"
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, oRecordSet, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        Select Case nMode
            Case 0
                GenerateReport "Daily Actual Manpower Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy"), "PRVActManS.rpt"
            Case 1
                GenerateReport "Daily Actual Manpower Detailed Report for " & Format(DTPicker1.Value, "mmm d, yyyy"), "PRVActManD.rpt"
                
        End Select

        ShowProgress 4
    Else
        ShowProgress 4
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If

    Set oRecordSet = Nothing
End Sub

Sub GenActWLCost(ByVal cDepid As String, nFilter As Integer)
    Dim cSqlStmt As String, _
        cParam As String, cParam2 As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oLeaveRSet As New ADODB.Recordset, _
        aTimeInfo As Variant, _
        aPayInfo As Variant, _
        nPayStat As Integer, _
        nPClose As Integer, _
        n13mopay As Double, _
        nIncentive As Double, _
        nTotDay As Double, _
        aLeaveInfo As Variant, _
        lWith13Mo As Boolean
        
    Dim aAdjustment As Variant

    CreateTemp 4
    ShowProgress 0
    
    aAdjustment = Array(0#, 0#)
    ' (0)   -   Regular Adjustment
    ' (1)   -   SA Adjustment
            
    If Trim(cDepid) <> "" Then cDepid = " where (lineid in " & cDepid & ")"
    
    
    cSqlStmt = "select lineid,linename from di5463 " & cDepid
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            

            nCtr = 0
            
            aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
            
            cSqlStmt = " select pclose from pa7730 " & _
                       " where (date_start between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                       " ) or (date_end between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")"
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

            nPClose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), 0)
            

            cSqlStmt = " SELECT b.date_hire,b.date_res,b.date_fin,a.EMPID,if (count(b.empid) > 1 ,1,count(b.empid)) as manpower, a.DATE,b.rate_amt, b.active, b.FIRSTNAME, b.MNAME, b.LASTNAME, " & _
                       " concat(b.LASTNAME,', ', b.FIRSTNAME,' ' ,left(b.MNAME,1),'') as fullname, " & _
                       " b.DEPID , c.LINENAME, b.emp_stat, b.wap, b.paystatus,b.COLA_AMT,b.POS_ALLOW, " & _
                       " b.SL_AVAIL,b.VL_AVAIL,b.SL_USE,b.VL_USE,b.YTD_BASIC, b.YTD_GROSS,b.YTD_GROSS_SA,b.YTD_COLA," & _
                       " ifnull(b.BACCNTNO,'') as BACCNTNO " & _
                       " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & " a " & _
                       " left join di3670 b on a.empid=b.empid  " & _
                       " left join di5463 c on b.depid=c.lineid " & _
                       " where  a.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & _
                       IIf(nFilter = 0, " and (b.emp_stat <> 0)", IIf(nFilter = 1, " and (b.emp_stat<>0)", IIf(nFilter = 2, " and (b.emp_stat=0)", IIf(nFilter = 4, "", " and (b.emp_stat=0)")))) & _
                       " and (b.paystatus " & IIf(nFilter = 4, "=2 )", "=0 )") & _
                       " and b.depid = " & cQuote & oRecordSet("lineid") & cQuote & _
                       " and (a.reg_hr+a.nd_hr+a.sun_hr+a.sun_nd)<>0 " & _
                       " group by a.empid"
'
'IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
'                   " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'
'        cSqlStmt = cSqlStmt & " where (a.active" & IIf(nActive = 0, "=0", "<>0") & ")" & IIf(Trim(cParam) = "", "", " and (" & cParam & ")") & _
'                   IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
'                   " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'                   " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
'                   IIf(Combo1.ListIndex = 2, " and a.emp_stat=0", "") & _
'                   " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
'                   IIf((Check4.Value <> vbChecked), " group by a.depid", "")
                       
            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, oTempADO, False
            If oTempADO.RecordCount > 0 Then
                While Not oTempADO.EOF
                
'                    If oTempADO("empid") = "0229" Then MsgBox "stop"
                
                    nCtr = nCtr + 1
                
                    ' --> For Emergency Manpower...
                    If nPayStat <> oTempADO("paystatus") Then
                        nPayStat = oTempADO("paystatus")
'                        nCtr = 0
                    End If
                
                    n13mopay = 0
                    nTotDay = 0
                    nIncentive = 0
                    
                    aLeaveInfo = Array(0#, 0#)
                    aAdjustment = Array(0#, 0#)
                    ' --> compute leave avail here...
                    cSqlStmt = "select " & _
                               "  if(date_start<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ",date_start) as date_start, " & _
                               "  if(date_end>=" & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ",date_end) as date_end, " & _
                               "  tag , paytag " & _
                               "From pa367583 " & _
                               "where (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (status=1) " & _
                               "  and (paytag=0) and (tag in (0,1" & IIf(gCompanyID = 2, ",6", "") & ")) " & _
                               "  and ((date_start between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") " & _
                               "    or (date_end between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ")) "
    '                Script2File cSqlStmt
                    OpenQueryDNS cSqlStmt, oLeaveRSet, False
                    

                    If oLeaveRSet.RecordCount > 0 Then
                        aLeaveInfo = ChkLeave(oLeaveRSet)
                        nTotDay = aLeaveInfo(0) - aLeaveInfo(1)
                        nIncentive = Round(nTotDay * (oTempADO("RATE_AMT") / IIf(oTempADO("paystatus") = 0, 1, 26.08)), 2)
                    End If
                    
                    ' --> compute "predata" here...
                    If oTempADO("active") > 0 Then
                        nTotDay = (DateDiff("d", "01/01/" & Year(Now), IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")))) / 365
                        nIncentive = Round(nIncentive + ((nTotDay * (oTempADO("SL_AVAIL") + oTempADO("VL_AVAIL") - oTempADO("SL_USE") - oTempADO("VL_USE"))) * (oTempADO("RATE_AMT") / IIf(oTempADO("paystatus") = 0, 1, 26.08))), 2)
                    End If
                    
                    aTimeInfo = CheckDTR(oTempADO("EMPID"), _
                                     Array(Format(DTPicker1.Value, "yyyy-mm-dd"), Format(DTPicker2.Value, "yyyy-mm-dd"), 0), _
                                     Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))
                        
                    aPayInfo(0) = Round(Round(aTimeInfo(0), 2) * oTempADO("RATE_AMT"), 2)                            ' --> reg pay
                    aPayInfo(1) = Round(Round(aTimeInfo(1), 2) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)             ' --> reg ot pay
                    aPayInfo(2) = Round(Round(aTimeInfo(2), 2) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)             ' --> sa reg ot pay

                    aPayInfo(3) = Round(Round(aTimeInfo(3), 2) * (oTempADO("RATE_AMT") * 1.1), 2)                    ' --> ndiff pay
                    aPayInfo(4) = Round(Round(aTimeInfo(4), 2) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)       ' --> ndiff ot pay
                    aPayInfo(12) = Round(Round(aTimeInfo(12), 2) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)     ' --> sa ndiff ot pay

                    aPayInfo(5) = Round(Round(aTimeInfo(5), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3), 2)              ' --> sun pay
                    aPayInfo(6) = Round(Round(aTimeInfo(6), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.3), 2)        ' --> sun ot pay

                    aPayInfo(13) = Round(Round(aTimeInfo(13), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1), 2)      ' --> sun ndiff pay
                    aPayInfo(14) = Round(Round(aTimeInfo(14), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1 * 1.3), 2) ' --> sun ndiff ot pay
'
                    If gCompanyID <> "0004" Or gCompanyID <> "0001" Then
                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)
                    Else
                        aPayInfo(9) = aPayInfo(0) + Round(Round(aTimeInfo(3), 2) * (oTempADO("RATE_AMT")), 2)
                    End If
'                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)

                    If Round(aTimeInfo(7), 2) <> 0 Then
                        aPayInfo(7) = Round((Round(aTimeInfo(7), 2) * oTempADO("RATE_AMT")) + (oTempADO("COLA_AMT") * Round(aTimeInfo(7), 2)), 2)                          ' --> holiday pay
                    End If

                    ' --> addendum, 20080313 Incentive Hour
                    If gCompanyID = "0002" Then
                        aPayInfo(15) = Round(Round(aTimeInfo(15), 2) * (oTempADO("RATE_AMT") / 8), 2)                    ' --> Incentive Pay
                    Else
                        aPayInfo(15) = Round(aTimeInfo(15), 2)                                                            ' --> Incentive Pay
                    End If

                    If lExtension Then
                    
                        If gCompanyID = "0001" Or gCompanyID = "0006" Or gCompanyID = "0005" Then
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '              SARegOTPay +
                            '              SANDiffOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay +
                            '                PosAllow +
                            '                COLA +
                            '                Incentive Leave +
                            '                Adjustment             --> ala p 2 computation d2...
                            
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(7) + _
                                           IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(Round((oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                                           
                            'If gCompanyID = "0005" Then
                            If (gCompanyID = "0005") Or (gCompanyID = "0006") Then
                                aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                                aPayInfo(11) = 0
                            End If
                                                                   
                        Else
                    
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay +
                            '                PosAllow +
                            '                COLA +
                            '                Incentive Leave +
                            '                Adjustment             --> ala p 2 computation d2...
                            
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(7) + _
                                           IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SARegOTPay +
                            '              SANDiffOTPay +
                            '              SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           Round((oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                        End If
                                
                        If nPayStat = 2 Then
                            aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                            aPayInfo(11) = 0
                        End If
                    Else
                    
                        If gCompanyID = "0001" Or gCompanyID = "0006" Or gCompanyID = "0005" Then
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '              SARegOTPay +
                            '              SANDiffOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay +
                            '                PosAllow +
                            '                COLA +
                            '                Incentive Leave +
                            '                Adjustment             --> ala p 2 computation d2...
                            
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(7) + _
                                           IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(Round((oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                                           
                                           
                            If gCompanyID = "0005" Then
                                aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                                aPayInfo(11) = 0
                            Else
                                If gCompanyID = "0001" Then
                                    If nPayStat = 2 Then
                                        aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                                        aPayInfo(11) = 0
                                    End If
                                
                                End If
                            End If
                        
                        Else
                    
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay +
                            '                PosAllow +
                            '                COLA +
                            '                Incentive Leave +
                            '                Adjustment             --> ala p 2 computation d2...
                            
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(7) + _
                                           IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SARegOTPay +
                            '              SANDiffOTPay +
                            '              SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           Round((oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                        End If
                    End If
                             
    
                    ' --> revised 20070105, 20070831 - add emergency here
                    If oTempADO("paystatus") = 2 Then
                        n13mopay = 0
                    Else
                        ' --> 13th month here...
                        If (oTempADO("active") > 0) And (oTempADO("wap") = 0) Then
                            ' --> revised 20070122
                            
                            lWith13Mo = True
'                            If (aPayInfo(0) + aPayInfo(3) + aPayInfo(5) + aPayInfo(13)) = 0 Then
'                                lWith13Mo = False
'                            Else
'                                lWith13Mo = True
'                            End If
                                
                            If lWith13Mo Then
                                If (oTempADO("EMP_STAT") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)) Then
                                    
                                    n13mopay = Round((oTempADO("YTD_BASIC") + aPayInfo(9)) / 12, 2)
                                ElseIf oTempADO("EMP_STAT") = 2 Then
                                
                                    n13mopay = Round((oTempADO("YTD_GROSS") + oTempADO("YTD_GROSS_SA") - oTempADO("YTD_COLA") + aPayInfo(10) + aPayInfo(11) - (oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2)))) / 12, 2)
                                Else
                                    n13mopay = 0
                                End If
                            End If
                        Else
                            n13mopay = 0
                        End If
                    End If

                    If nFilter <> 0 Then
                        
                    End If
                    cSqlStmt = " insert into tmpAWLCost([EMPID],[ACTIVE],[EMP_STAT],[RATE_AMT],[COLA_AMT],[POS_ALLOW],[REG_DAY],[REG_PAY],[REG_OT_HR],[REG_OT_PAY],[NDIFF_DAY],[NDIFF_PAY],[NDIFF_OT_HR],[NDIFF_OT_PAY],[HOLIDAY],[HOL_PAY],[SA_REG_OT],[SA_REG_PAY],[SA_NDIFF_OT],[SA_NDIFF_PAY],[SUN_HR],[SUN_PAY],[SUN_OT],[SUN_OT_PAY], " & _
                               " [ADJ_PAY],[SA_ADJ_PAY],[OTHER_PAY],[LEAVE_PAY],[M13PAY],[GROSS_PAY],[SA_NET_PAY],[PAYSTATUS],[FIRSTNAME],[MNAME],[LASTNAME],[WAP],[date_res],[DATE_HIRE],[SEQ_NO],[COLA],[SUN_COLA],[SUN_ND],[SUN_ND_PAY],[SUN_ND_OT],[SUN_ND_OT_PAY],[BACCNTNO],[date_start],[date_end],[DEPNAME],[fullname])values( " & _
                                cQuote & oTempADO("EMPID") & cQuote & "," & oTempADO("ACTIVE") & "," & oTempADO("EMP_STAT") & "," & IIf(Check4.Value = vbUnchecked, oTempADO("manpower"), oTempADO("RATE_AMT")) & "," & oTempADO("COLA_AMT") & "," & IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) & "," & Round(aTimeInfo(0), 2) & "," & _
                                aPayInfo(0) & "," & Round(aTimeInfo(1), 2) & "," & aPayInfo(1) & "," & Round(aTimeInfo(3), 2) & "," & aPayInfo(3) & "," & Round(aTimeInfo(4), 2) & "," & aPayInfo(4) & "," & Round(aTimeInfo(7), 2) & "," & aPayInfo(7) & "," & Round(aTimeInfo(2), 2) & "," & aPayInfo(2) & "," & Round(aTimeInfo(12), 2) & "," & aPayInfo(12) & "," & _
                                Round(aTimeInfo(5), 2) & "," & aPayInfo(5) & "," & Round(aTimeInfo(6), 2) & "," & aPayInfo(6) & "," & aAdjustment(0) & "," & aAdjustment(1) & "," & aPayInfo(20) & "," & nIncentive & "," & n13mopay & "," & aPayInfo(10) + n13mopay & "," & aPayInfo(11) & "," & _
                                oTempADO("PAYSTATUS") & "," & cQuote & DecodeStr(oTempADO("FIRSTNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("MNAME")) & cQuote & "," & cQuote & oTempADO("LASTNAME") & cQuote & "," & oTempADO("WAP") & "," & _
                                cQuote & Format(IIf((oTempADO("active") = 1) Or (oTempADO("active") = 3), oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & nCtr & "," & _
                                Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) & "," & Round(oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8), 2) & "," & Round(aTimeInfo(13), 2) & "," & aPayInfo(13) & "," & Round(aTimeInfo(14), 2) & "," & aPayInfo(14) & "," & cQuote & oTempADO("BACCNTNO") & cQuote & "," & _
                                cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & DecodeStr(oRecordSet("LINENAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("Fullname")) & cQuote & ")"
        
'                    MsgBox cSqlStmt
'                    Script2File cSqlStmt
                    QueryTemp cSqlStmt, objdbRs, True
                        
                    
                    oTempADO.MoveNext
                Wend
            End If
            
            'd2 ang insert
            
                    
            oRecordSet.MoveNext
        Wend

        ShowProgress 3

        Select Case nFilter
            Case 0
                If Check4.Value = vbUnchecked Then
                    GenerateReport "Actual Labor Cost Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCost_s.rpt"
                Else
                    GenerateReport "Actual Labor Cost Detail Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCost_d.rpt"
                End If
            Case 1, 3
                If Check4.Value = vbUnchecked Then
                    GenerateReport "Actual Labor Cost" & IIf(nFilter = 3, " WAP", "") & " SA Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostSA_s.rpt"
                Else
                    GenerateReport "Actual Labor Cost" & IIf(nFilter = 3, " WAP", "") & " SA Detail Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostSA_d.rpt"
                End If
            Case 2
                If Check4.Value = vbUnchecked Then
                    GenerateReport "Actual Labor Cost WAP Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostWAP_s.rpt"
                Else
                    GenerateReport "Actual Labor Cost WAP Detail Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostWAP_d.rpt"
                End If
                
            Case 4
                If Check4.Value = vbUnchecked Then
                    GenerateReport "Actual Labor Cost Emergency Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostEM_s.rpt"
                Else
                    GenerateReport "Actual Labor Cost Emergency Detail Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvActWLCostEM_d.rpt"
                End If
                
        End Select


        ShowProgress 4
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub

Sub CreateTmpWork(ByVal nMode As Integer)
On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
        cSqlStmt = "create table tmpWork( [compcode] char(4)," & _
                                        " [empid] char(8),              [fullname] char(100), " & _
                                        " [workcenterid] char(20),      [BEPworkid] char(20), " & _
                                        " [costcenterid] char(20),      [linename] char(100), " & _
                                        " [posid] char(8),              [posname] char(50), " & _
                                        " [emp_stat] char(20),          [Active] char(10), " & _
                                        " [ERP_Active] char(10),        [Paystatus] char(10)," & _
                                        " [ERPPOSCODE] char(1),         [WORKDATE] date , " & _
                                        " [Tag] integer,                [Remarks] char(100) )"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpWork"
    QueryTemp cSqlStmt, oTempADO, True
End Sub
Sub GenWorkCenter(ByVal cDepid As String)
    Dim cSqlStmt, _
        cParam As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        nPClose As Integer, _
        sRemarks As String, _
        ntag As String, _
        nERPPOSCODE As String

    CreateTmpWork 0

    ShowProgress 0
        
    If Trim(cDepid) <> "" Then
        cDepid = " and b.depid IN " & cDepid
    End If

    cSqlStmt = "select pclose,periodid from pa7730 " & _
                "where " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " between date_start and date_end"
    OpenQueryDNS cSqlStmt, objdbRs, False
    nPClose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), 0)


'    cSqlStmt = "SELECT a.date,g.compcode as Plant, a.empid , concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'. ') as Workername, ifnull(f.workcenterid,' ') as Work_center," & _
'                "   ifnull(e.costcenterid,' ') as Cost_center, " & _
'                "   ifnull(c.LINENAME,' ') as LINENAME," & _
'                "   d.posid as Pos_ID, d.POSNAME as posname, b.EMP_STAT, b.ACTIVE, b.PAYSTATUS" & _
'                " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & " a " & _
'                "   left join di3670 b on a.empid = b.empid " & _
'                "   left join di5463 c on b.depid = c.lineid " & _
'                "   left join di7670 d on b.posid = d.posid " & _
'                "   left join pa37722 e on b.costcenterid=e.costcenterid " & _
'                "   left join pa97722 f on b.workcenterid=f.workcenterid " & _
'                "   left join pa2660 g on b.cmpid=f.cmpid " & _
'                " Where ((b.Active = 3) And (b.date_fin = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))" & _
'                " or ((b.active = 1) and (b.date_fin = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))" & _
'                " or ((b.active = 2) and (b.date_res = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))" & _
'                " or b.active = 0" & _
'                 cDepid & _
'                "   group by b.empid " & _
'                "   order by c.linename"
                
    cSqlStmt = "SELECT a.date,g.compcode as Plant, a.empid , concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'. ') as Workername, ifnull(f.workcenterid,' ') as Work_center,ifnull(b.bepworkcenterid,' ') as bep_Work_center," & _
                "   ifnull(e.costcenterid,' ') as Cost_center, " & _
                "   ifnull(c.LINENAME,' ') as LINENAME," & _
                "   ifnull(c.ERPPOSCODE,0) as ERPPOSCODE," & _
                "   d.posid as Pos_ID, d.POSNAME as posname, b.EMP_STAT, b.ACTIVE, b.ERP_ACTIVE, b.PAYSTATUS,a.tag" & _
                " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & " a " & _
                "   left join di3670 b on a.empid = b.empid " & _
                "   left join di5463 c on b.depid = c.lineid " & _
                "   left join di7670 d on b.posid = d.posid " & _
                "   left join pa37722 e on b.costcenterid=e.costcenterid " & _
                "   left join pa97722 f on b.workcenterid=f.workcenterid " & _
                "   left join pa2660 g on b.cmpid=f.cmpid " & _
                " where (((b.active=1) or (b.active=3)) and ((b.date_res = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((b.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (b.date_res > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) or " & _
                " ((b.active=2) and ((b.date_fin = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or ((b.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") and (b.date_fin > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))) " & _
                " or ((b.ACTIVE=0) and (b.date_hire<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) " & _
                cDepid & _
                "   group by b.empid " & _
                "   order by b.paystatus,c.linename"
                    
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
         While Not oRecordSet.EOF
         
'            If oRecordSet("empid") = "364" Then MsgBox "stop!!"
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            Select Case oRecordSet("ERPPOSCODE")
                Case 0 'A'
                    nERPPOSCODE = "A"
                Case 1 'B'
                    nERPPOSCODE = "B"
                Case 2 'C'
                    nERPPOSCODE = "C"
                Case 3 'D'
                    nERPPOSCODE = "D"
                Case 4 'E'
                    nERPPOSCODE = "E"
                Case 5 'F'
                    nERPPOSCODE = "F"
                Case 6 'G'
                    nERPPOSCODE = "G"
                Case 7 'Z'
                    nERPPOSCODE = "Z"
            
            End Select
            
            OpenQueryDNS " select tag " & _
                         " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & _
                         " di36770 where date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                         " and empid = " & cQuote & oRecordSet("empid") & cQuote, oRSet, False
                         
            ShowProgress 4
            
            If oRSet.RecordCount > 0 Then
                

                    sRemarks = ""
'                    If oRecordSet("tag") <> oRSet("tag") Then
'                        MsgBox "stop"
'                    End If
                    
                    Select Case oRSet("TAG")
                        Case 0
                            sRemarks = "Complete Entry"
                        Case 1 'B'
                            sRemarks = "No Entry / Absent"
                        Case 2 'C'
                            sRemarks = "On Leave"
                        Case 3 'D'
                            sRemarks = "Incomplete Entry"
                    End Select
                    ntag = oRSet("Tag")
            Else
                    ntag = "0"
                    sRemarks = ""
            End If
            
            
            cSqlStmt = "insert into tmpWork (compcode,empid,fullname,workcenterid,BEPworkid,costcenterid,linename,posid,posname,emp_stat,Active,ERP_ACTIVE,Paystatus,ERPPOSCODE,[Tag],[Remarks],[WORKDATE])values(" & _
                       cQuote & nCompCode & cQuote & "," & _
                       cQuote & IIf(gAgency <> 0, UCase(left(cODBC, 1)), "") & IIf(gAgency <> 0, left(right(nCompCode, 3), 1), left(nCompCode, 2)) & oRecordSet("empid") & cQuote & "," & _
                       cQuote & oRecordSet("workername") & cQuote & "," & _
                       cQuote & oRecordSet("Work_center") & cQuote & "," & _
                       cQuote & oRecordSet("bep_Work_center") & cQuote & "," & _
                       cQuote & oRecordSet("Cost_center") & cQuote & "," & _
                       cQuote & oRecordSet("linename") & cQuote & "," & _
                       cQuote & oRecordSet("Pos_ID") & cQuote & "," & _
                       cQuote & oRecordSet("posname") & cQuote & "," & _
                       cQuote & IIf(oRecordSet("emp_stat") = 0, "Wap", IIf(oRecordSet("emp_stat") = 1, "Contractual", "Regular")) & cQuote & "," & _
                       cQuote & IIf(oRecordSet("Active") = 0, "Active", IIf(oRecordSet("Active") = 1, "Resigned", IIf(oRecordSet("Active") = 2, "Finished", "Terminated"))) & cQuote & "," & _
                       cQuote & IIf(oRecordSet("ERP_Active") = 0, "Active", "In-Active") & cQuote & "," & _
                       cQuote & IIf(oRecordSet("paystatus") = 0, "Daily", IIf(oRecordSet("paystatus") = 1, "Monthly", "Emergency")) & cQuote & "," & _
                       cQuote & nERPPOSCODE & cQuote & "," & _
                       cQuote & ntag & cQuote & "," & _
                       cQuote & sRemarks & cQuote & "," & _
                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")"
'            MsgBox cSqlStmt

            QueryTemp cSqlStmt, objdbRs, True
        
        oRecordSet.MoveNext
        
        Wend
        ShowProgress 3

        GenerateReport "Employee Work Center Listing ", "RPTWCenter.RPT"
        
        ShowProgress 4
        
Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    Set oRecordSet = Nothing
End Sub


Sub CreateTempTMSERP()
On Error GoTo ErrCreate
    Dim cSqlStmt As String
              

    cSqlStmt = " CREATE TABLE TMPTMSERP( [DATE] date,               [WORKDATE] date,                [EMPID] char(8)," & _
               " [FIRSTNAME] char(50),      [MNAME] char(50),       [LASTNAME] char(50),        [FULLNAME] char(100)," & _
               " [EMP_STAT] integer,        [ACTIVE] integer,       [PAYSTATUS] integer,        [LineID] char(3),       [LINENAME] char(100), " & _
               " [POSID] char(3),           [POSNAME] char(50),     [COSTCENTERID] char(10),    [COSTDESC] char(100),   [WORKCENTERID] char(10),    [WORKDESC] char(100), [compcode] char(4),  " & _
               " [reg_hr] double," & _
               " [reg_ot_hr] double,        [sa_reg_ot] double," & _
               " [nd_hr] double,            [nd_ot_hr] double," & _
               " [sa_nd_ot] double,         [sun_hr] double," & _
               " [sun_ot_hr] double,        [sun_nd] double," & _
               " [sun_nd_ot] double,         " & _
               " [rmin] double,             [rotmin] double, " & _
               " [rotmintot] double,        [ndmin] double, " & _
               " [ndotmin] double,          [ndotmintot] double, " & _
               " [sun_min] double,          [sunotmin] double, " & _
               " [suntot] double,           [ERPPOSCODE] char(1), " & _
               " [sunndmin] double,         [sunndotmin] double, " & _
               " [sunndtot] double,         [munitetot] double, " & _
               " [outtrantime] char(15),    [intrantime] char(15))"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPTMSERP"
    QueryTemp cSqlStmt, oTempADO, True
End Sub


Sub GenTMSERP(ByVal cDepid As String)
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        aothertot As Variant, _
        aOtherInfo As Variant, _
        aTrantype As Variant, _
        dLogDate As Date, _
        cPclose As String
        
        
    'revision feb 12,2013
    '0 - A
    '1 - B
    '2 - C
    '3 - D
    '4 - E
    '5 - F
    '6 - G
    '7 - Z

    'aothertot *****************************
    '0 = rmin           reg hour munite
    '1 = rotmin         reg hour ot munite Total
    '2 = rotmintot      reg all Total
    '3 = ndmin          nd reg min
    '4 = ndotmin        nd reg ot munite total
    '5 = ndotmintot     nd all total
    '6 = sun_min        sun hr munite
    '7 = sunotmin       sun ot total
    '8 = suntot         sun all total
    '9 = sunndmin       sun nd minute
    '10 = sunndotmin    sun nd ot minute total
    '11 = sunndtot      sun nt all total
    '12 = munitetot     munite all total
    
    aothertot = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)

    'aotherinfo *************
    '0 = Department name
    '1 = Position name
    '2 = CostCenter Description
    '3 = WorkCenter Description
    '4 = ERPPOSCODE
    aOtherInfo = Array("", "", "", "", "")
    
    aTrantype = Array("", "", "")

    CreateTempTMSERP
    
    
    ShowProgress 0
    
    If Trim(cDepid) <> "" Then
        cDepid = " and b.depid IN " & cDepid
    End If
    
    
    OpenQueryDNS "select PERIODID, PCLOSE from pa7730 where " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " between date_start and  date_end and 13month = 0", objdbRs, False
    cPclose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), "")
    'new revise
    cSqlStmt = " select a.EMPID, a.DATE, a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.nd_hr, a.nd_ot_hr, a.sa_nd_ot, a.sun_hr, a.sun_ot_hr,a.sun_nd, a.sun_nd_ot, " & _
               " b.FIRSTNAME, b.MNAME, b.LASTNAME, concat(b.LASTNAME,', ',b.FIRSTNAME, ' ' ,left(b.MNAME,1),'. ') as fullname,  " & _
               " b.DEPID, b.POSID, b.EMP_STAT, b.ACTIVE, b.PAYSTATUS, b.COSTCENTERID , b.WORKCENTERID " & _
               " from " & IIf(cPclose = 0, "di36770", "dih36770") & " a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " where date = " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
               " and (a.reg_hr<>0 or a.nd_hr<>0 or a.sun_hr<>0 or a.sun_nd<>0) " & _
               cDepid
    'Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
             ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
        
            aothertot = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
            
            aothertot(0) = oRecordSet("reg_hr") * 60
            aothertot(1) = (oRecordSet("reg_ot_hr") + oRecordSet("sa_reg_ot")) * 60
            aothertot(2) = (oRecordSet("reg_hr") + oRecordSet("reg_ot_hr") + oRecordSet("sa_reg_ot")) * 60
            aothertot(3) = oRecordSet("nd_hr") * 60
            aothertot(4) = (oRecordSet("nd_ot_hr") + oRecordSet("sa_nd_ot")) * 60
            aothertot(5) = (oRecordSet("nd_hr") + oRecordSet("nd_ot_hr") + oRecordSet("sa_nd_ot")) * 60
            aothertot(6) = oRecordSet("sun_hr") * 60
            aothertot(7) = oRecordSet("sun_ot_hr") * 60
            aothertot(8) = (oRecordSet("sun_hr") + oRecordSet("sun_ot_hr")) * 60
            aothertot(9) = oRecordSet("sun_nd") * 60
            aothertot(10) = oRecordSet("sun_nd_ot") * 60
            aothertot(11) = (oRecordSet("sun_nd") + oRecordSet("sun_nd_ot")) * 60
            aothertot(12) = (oRecordSet("reg_hr") + oRecordSet("reg_ot_hr") + oRecordSet("sa_reg_ot") + oRecordSet("nd_hr") + oRecordSet("nd_ot_hr") + oRecordSet("sa_nd_ot") + oRecordSet("sun_hr") + oRecordSet("sun_ot_hr") + oRecordSet("sun_nd") + oRecordSet("sun_nd_ot")) * 60
            
            aOtherInfo = Array("", "", "", "", "")

            OpenQueryDNS "select * from di5463 where lineid = " & cQuote & oRecordSet("depid") & cQuote, objdbRs, False
            aOtherInfo(0) = IIf(objdbRs.RecordCount > 0, objdbRs("linename"), "")
                
                Select Case objdbRs("ERPPOSCODE")
                Case 0 'A'
                    aOtherInfo(4) = "A"
                Case 1 'B'
                    aOtherInfo(4) = "B"
                Case 2 'C'
                    aOtherInfo(4) = "C"
                Case 3 'D'
                    aOtherInfo(4) = "D"
                Case 4 'E'
                    aOtherInfo(4) = "E"
                Case 5 'F'
                    aOtherInfo(4) = "F"
                Case 6 'G'
                    aOtherInfo(4) = "G"
                Case 7 'Z'
                    aOtherInfo(4) = "Z"
                Case "" ' '
                    aOtherInfo(4) = "Z"

            End Select
                
            OpenQueryDNS "select * from di7670 where posid = " & cQuote & oRecordSet("posid") & cQuote, objdbRs, False
            aOtherInfo(1) = IIf(objdbRs.RecordCount > 0, objdbRs("posname"), "")
            OpenQueryDNS "select * from pa37722 where COSTCENTERID = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
            aOtherInfo(2) = IIf(objdbRs.RecordCount > 0, objdbRs("description"), "")
            OpenQueryDNS "select * from pa97722 where WORKCENTERID = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
            aOtherInfo(3) = IIf(objdbRs.RecordCount > 0, objdbRs("description"), "")
            
            aTrantype = Array("", "", "")
            
            cSqlStmt = " select TRAN_NO, EMPID, LOGDATE, TRANSDATE, TRANTIME, TRANTYPE from pa84650 " & _
                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and (logdate = " & _
                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") order by logdate,transdate, trantime"
            OpenQueryDNS cSqlStmt, oRSet, False
            If oRSet.RecordCount > 0 Then
                While Not oRSet.EOF
                    
                    If oRSet("trantype") = 0 Then

                        If Trim(aTrantype(1)) = "" Then
                            aTrantype(0) = oRSet("trantype")
                            aTrantype(1) = oRSet("trantime")
                            dLogDate = oRSet("logdate")
                        End If

                    Else
                        aTrantype(0) = oRSet("trantype")
                        aTrantype(2) = oRSet("trantime")
                        dLogDate = oRSet("logdate")
                    End If
                    
                    If Not oRSet.EOF Then
                        If dLogDate = oRSet("logdate") Then
                            If (oRSet("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                    aTrantype = Array("", "", "", "")
                            End If
                        Else
                            aTrantype = Array("", "", "", "")
                        End If
                    Else
                        aTrantype = Array("", "", "", "")
                    End If
                    
                    oRSet.MoveNext
                Wend
                
            End If
            
            
            cSqlStmt = "insert into TMPTMSERP ([DATE],[WORKDATE],EMPID,FIRSTNAME,MNAME,LASTNAME,FULLNAME,EMP_STAT,[ACTIVE],PAYSTATUS,LineID,LINENAME,POSID,POSNAME,COSTCENTERID,COSTDESC,WORKCENTERID,WORKDESC, " & _
                       " reg_hr,reg_ot_hr,sa_reg_ot,nd_hr,nd_ot_hr,sa_nd_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot,rmin,rotmin,rotmintot,ndmin,ndotmin,ndotmintot,sun_min,sunotmin,suntot,sunndmin,sunndotmin,sunndtot,munitetot,intrantime,outtrantime,compcode,ERPPOSCODE)values(" & _
                       cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & IIf(gAgency <> 0, UCase(left(cODBC, 1)), "") & IIf(gAgency <> 0, left(right(nCompCode, 3), 1), left(nCompCode, 2)) & oRecordSet("empid") & cQuote & "," & _
                       cQuote & oRecordSet("FIRSTNAME") & cQuote & "," & cQuote & oRecordSet("MNAME") & cQuote & "," & cQuote & oRecordSet("LASTNAME") & cQuote & "," & cQuote & oRecordSet("FULLNAME") & cQuote & "," & _
                       oRecordSet("emp_stat") & "," & oRecordSet("ACTIVE") & "," & oRecordSet("PAYSTATUS") & "," & _
                       cQuote & oRecordSet("depid") & cQuote & "," & cQuote & aOtherInfo(0) & cQuote & "," & _
                       cQuote & oRecordSet("POSID") & cQuote & "," & cQuote & aOtherInfo(1) & cQuote & "," & _
                       cQuote & oRecordSet("COSTCENTERID") & cQuote & "," & cQuote & aOtherInfo(2) & cQuote & "," & _
                       cQuote & oRecordSet("WORKCENTERID") & cQuote & "," & cQuote & EncodeStr(DecodeStr(aOtherInfo(3))) & cQuote & "," & _
                       oRecordSet("reg_hr") & "," & oRecordSet("reg_ot_hr") & "," & oRecordSet("sa_reg_ot") & "," & oRecordSet("nd_hr") & "," & oRecordSet("nd_ot_hr") & "," & oRecordSet("sa_nd_ot") & "," & oRecordSet("sun_hr") & "," & oRecordSet("sun_ot_hr") & "," & oRecordSet("sun_nd") & "," & oRecordSet("sun_nd_ot") & "," & _
                       aothertot(0) & "," & aothertot(1) & "," & aothertot(2) & "," & aothertot(3) & "," & aothertot(4) & "," & aothertot(5) & "," & aothertot(6) & "," & aothertot(7) & "," & aothertot(8) & "," & aothertot(9) & "," & aothertot(10) & "," & aothertot(11) & "," & aothertot(12) & "," & cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                       nCompCode & "," & cQuote & aOtherInfo(4) & cQuote & ")"
        

'            Script2File cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
        
            oRecordSet.MoveNext
            
        Wend
        ShowProgress 3

        GenerateReport "ERP - TMS Listing ", "rptERPTMS.RPT"
        
        ShowProgress 4
        
    Else
        ShowProgress 4
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    Set oRecordSet = Nothing
End Sub

Sub GenDLCost(ByVal cDepid As String, nFilter As Integer)
    Dim cSqlStmt As String, _
        cParam As String, cParam2 As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset, _
        oLeaveRSet As New ADODB.Recordset, _
        aTimeInfo As Variant, _
        aPayInfo As Variant, _
        nPayStat As Integer, _
        nPClose As Integer, _
        n13mopay As Double, _
        nIncentive As Double, _
        nTotDay As Double, _
        aLeaveInfo As Variant, _
        lWith13Mo As Boolean
        
    Dim aAdjustment As Variant

    CreateTemp 5
    ShowProgress 0
    
    aAdjustment = Array(0#, 0#)
'    ' (0)   -   Regular Adjustment
'    ' (1)   -   SA Adjustment
            
    If Trim(cDepid) <> "" Then cDepid = " where (lineid in " & cDepid & ")"
    
    cSqlStmt = "select lineid,linename from di5463 " & cDepid
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            nCtr = 0
            
            aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
            
            cSqlStmt = " select pclose from pa7730 " & _
                       " where (date_start between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                       " ) or (date_end between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")"
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

            nPClose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), 0)
            
            cSqlStmt = " SELECT b.date_hire,b.date_res,b.date_fin,a.EMPID,if (count(b.empid) > 1 ,1,count(b.empid)) as manpower, a.DATE,b.rate_amt, b.active, b.FIRSTNAME, b.MNAME, b.LASTNAME, " & _
                       " concat(b.LASTNAME,', ', b.FIRSTNAME,' ' ,left(b.MNAME,1),'') as fullname, " & _
                       " b.DEPID , c.LINENAME, b.emp_stat, b.wap, b.paystatus,b.COLA_AMT,b.POS_ALLOW, " & _
                       " b.SL_AVAIL,b.VL_AVAIL,b.SL_USE,b.VL_USE,b.YTD_BASIC, b.YTD_GROSS,b.YTD_GROSS_SA,b.YTD_COLA," & _
                       " ifnull(b.BACCNTNO,'') as BACCNTNO " & _
                       " FROM " & IIf(nPClose = 0, " di36770 ", " dih36770 ") & " a " & _
                       " left join di3670 b on a.empid=b.empid  " & _
                       " left join di5463 c on b.depid=c.lineid " & _
                       " where  a.date between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & _
                       " and b.depid = " & cQuote & oRecordSet("lineid") & cQuote & _
                       " and (a.reg_hr+a.nd_hr+a.sun_hr+a.sun_nd)<>0 " & _
                       " group by a.empid"
'
'IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
'                   " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'
'        cSqlStmt = cSqlStmt & " where (a.active" & IIf(nActive = 0, "=0", "<>0") & ")" & IIf(Trim(cParam) = "", "", " and (" & cParam & ")") & _
'                   IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
'                   " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'                   " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
'                   IIf(Combo1.ListIndex = 2, " and a.emp_stat=0", "") & _
'                   " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
'                   IIf((Check4.Value <> vbChecked), " group by a.depid", "")
                       
            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, oTempADO, False
            If oTempADO.RecordCount > 0 Then
                While Not oTempADO.EOF
                
'                    If oTempADO("empid") = "0229" Then MsgBox "stop"
                
                    nCtr = nCtr + 1
                
''                     --> For Emergency Manpower...
''                    If nPayStat <> oTempADO("paystatus") Then
''                        nPayStat = oTempADO("paystatus")
''                        nCtr = 0
''                    End If
                
                    n13mopay = 0
                    nTotDay = 0
                    nIncentive = 0
                    
                    aLeaveInfo = Array(0#, 0#)
                    aAdjustment = Array(0#, 0#)
                    ' --> compute leave avail here...
                    cSqlStmt = "select " & _
                               "  if(date_start<=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ",date_start) as date_start, " & _
                               "  if(date_end>=" & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ",date_end) as date_end, " & _
                               "  tag , paytag " & _
                               "From pa367583 " & _
                               "where (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (status=1) " & _
                               "  and (paytag=0) and (tag in (0,1" & IIf(gCompanyID = 2, ",6", "") & ")) " & _
                               "  and ((date_start between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ") " & _
                               "    or (date_end between " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & ")) "
    '                Script2File cSqlStmt
                    OpenQueryDNS cSqlStmt, oLeaveRSet, False
                    

                    If oLeaveRSet.RecordCount > 0 Then
                        aLeaveInfo = ChkLeave(oLeaveRSet)
                        nTotDay = aLeaveInfo(0) - aLeaveInfo(1)
                        nIncentive = Round(nTotDay * (oTempADO("RATE_AMT") / IIf(oTempADO("paystatus") = 0, 1, 26.08)), 2)
                    End If
                    
                    ' --> compute "predata" here...
                    If oTempADO("active") > 0 Then
                        nTotDay = (DateDiff("d", "01/01/" & Year(Now), IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")))) / 365
                        nIncentive = Round(nIncentive + ((nTotDay * (oTempADO("SL_AVAIL") + oTempADO("VL_AVAIL") - oTempADO("SL_USE") - oTempADO("VL_USE"))) * (oTempADO("RATE_AMT") / IIf(oTempADO("paystatus") = 0, 1, 26.08))), 2)
                    End If
                    
                    aTimeInfo = CheckDTR(oTempADO("EMPID"), _
                                     Array(Format(DTPicker1.Value, "yyyy-mm-dd"), Format(DTPicker2.Value, "yyyy-mm-dd"), 0), _
                                     Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))
                        
                    aPayInfo(0) = Round(Round(aTimeInfo(0), 2) * oTempADO("RATE_AMT"), 2)                            ' --> reg pay
                    aPayInfo(1) = Round(Round(aTimeInfo(1), 2) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)             ' --> reg ot pay
                    aPayInfo(2) = Round(Round(aTimeInfo(2), 2) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)             ' --> sa reg ot pay

                    aPayInfo(3) = Round(Round(aTimeInfo(3), 2) * (oTempADO("RATE_AMT") * 1.1), 2)                    ' --> ndiff pay
                    aPayInfo(4) = Round(Round(aTimeInfo(4), 2) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)       ' --> ndiff ot pay
                    aPayInfo(12) = Round(Round(aTimeInfo(12), 2) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)     ' --> sa ndiff ot pay

                    aPayInfo(5) = Round(Round(aTimeInfo(5), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3), 2)              ' --> sun pay
                    aPayInfo(6) = Round(Round(aTimeInfo(6), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.3), 2)        ' --> sun ot pay

                    aPayInfo(13) = Round(Round(aTimeInfo(13), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1), 2)      ' --> sun ndiff pay
                    aPayInfo(14) = Round(Round(aTimeInfo(14), 2) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1 * 1.3), 2) ' --> sun ndiff ot pay
'
                    If gCompanyID <> "0004" Or gCompanyID <> "0001" Then
                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)
                    Else
                        aPayInfo(9) = aPayInfo(0) + Round(Round(aTimeInfo(3), 2) * (oTempADO("RATE_AMT")), 2)
                    End If
'                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)

                    If Round(aTimeInfo(7), 2) <> 0 Then
                        aPayInfo(7) = Round((Round(aTimeInfo(7), 2) * oTempADO("RATE_AMT")) + (oTempADO("COLA_AMT") * Round(aTimeInfo(7), 2)), 2)                          ' --> holiday pay
                    End If

                    ' --> addendum, 20080313 Incentive Hour
                    If gCompanyID = "0002" Then
                        aPayInfo(15) = Round(Round(aTimeInfo(15), 2) * (oTempADO("RATE_AMT") / 8), 2)                    ' --> Incentive Pay
                    Else
                        aPayInfo(15) = Round(aTimeInfo(15), 2)                                                            ' --> Incentive Pay
                    End If

                    If lExtension Then
                    
                        If gCompanyID = "0001" Or gCompanyID = "0006" Then
                    
'                        If gCompanyID = "0001" Or gCompanyID = "0006" Or gCompanyID = "0005" Then
                           
                            
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(5) + aPayInfo(6) + aPayInfo(13) + aPayInfo(14) + _
                                           aPayInfo(20))
                         
'                            aPayInfo(11) = Round(aPayInfo(5) + _
'                                           aPayInfo(6) + _
'                                           aPayInfo(13) + _
'                                           aPayInfo(14) + _
'                                           aPayInfo(20))
'
                          
'                            If gCompanyID = "0006" Then
'                                aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                                aPayInfo(11) = 0
'                            End If
                                                                   
                        Else
                        
                           aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(5) + aPayInfo(6) + aPayInfo(13) + aPayInfo(14) + _
                                           aPayInfo(20))
                                                       
'                            aPayInfo(10) = Round(aPayInfo(0) + _
'                                           aPayInfo(1) + _
'                                           aPayInfo(3) + _
'                                           aPayInfo(4) + _
'                                           aPayInfo(7))
'
'
'                            aPayInfo(11) = Round(aPayInfo(2) + _
'                                           aPayInfo(12) + _
'                                           aPayInfo(5) + _
'                                           aPayInfo(6) + _
'                                           aPayInfo(13) + _
'                                           aPayInfo(14) + _
'                                           aPayInfo(20))
'
                                           
                        End If
                                
'                        If nPayStat = 2 Then
'                            aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                            aPayInfo(11) = 0
'                        End If
                    Else
                    
                        If gCompanyID = "0001" Or gCompanyID = "0006" Then
                        
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(5) + aPayInfo(6) + aPayInfo(13) + aPayInfo(14) + _
                                           aPayInfo(20))
                        
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '              SARegOTPay +
                            '              SANDiffOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay +
                           '         --> ala p 2 computation d2...
                            
'                            aPayInfo(10) = Round(aPayInfo(0) + _
'                                           aPayInfo(1) + _
'                                           aPayInfo(2) + _
'                                           aPayInfo(12) + _
'                                           aPayInfo(3) + _
'                                           aPayInfo(4) + _
'                                           aPayInfo(7))
'
'                            '    nNetAmt = SunCola +
'                            '              SunPay +
'                            '              SunOTPay +
'                            '              SunNDPay +
'                            '              SunNDOTPay +
'                            '                     --> ala p 2 computation d2...
'                            aPayInfo(11) = Round(aPayInfo(5) + _
'                                           aPayInfo(6) + _
'                                           aPayInfo(13) + _
'                                           aPayInfo(14) + _
'                                           aPayInfo(20))
'
                                           
'                            If gCompanyID = "0005" Then
'                                aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                                aPayInfo(11) = 0
'                            Else
'                                If gCompanyID = "0001" Then
'                                    If nPayStat = 2 Then
'                                        aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                                        aPayInfo(11) = 0
'                                    End If
'
'                                End If
'                            End If
                        
                        Else
                            aPayInfo(10) = Round(aPayInfo(0) + _
                                           aPayInfo(1) + _
                                           aPayInfo(2) + _
                                           aPayInfo(12) + _
                                           aPayInfo(3) + _
                                           aPayInfo(4) + _
                                           aPayInfo(5) + aPayInfo(6) + aPayInfo(13) + aPayInfo(14) + _
                                           aPayInfo(20))
                        
                    
                            '    nGrossAmt = RegPay +
                            '                RegOTPay +
                            '                NDiffPay +
                            '                NDiffOTPay +
                            '                HolPay
                            '                          --> ala p 2 computation d2...
                            
'                            aPayInfo(10) = Round(aPayInfo(0) + _
'                                           aPayInfo(1) + _
'                                           aPayInfo(3) + _
'                                           aPayInfo(4) + _
'                                           aPayInfo(7))
'
'                            '    nNetAmt = SARegOTPay +
'                            '              SANDiffOTPay +
'                            '              SunCola +
'                            '              SunPay +
'                            '              SunOTPay +
'                            '              SunNDPay +
'                            '              SunNDOTPay
'                            '                     --> ala p 2 computation d2...
'                            aPayInfo(11) = Round(aPayInfo(2) + _
'                                           aPayInfo(12) + _
'                                           aPayInfo(5) + _
'                                           aPayInfo(6) + _
'                                           aPayInfo(13) + _
'                                           aPayInfo(14) + _
'                                           aPayInfo(20))
                        End If
                    End If
                             
'
'                    ' --> revised 20070105, 20070831 - add emergency here
'                    If oTempADO("paystatus") = 2 Then
'                        n13mopay = 0
'                    Else
'                        ' --> 13th month here...
'                        If (oTempADO("active") > 0) And (oTempADO("wap") = 0) Then
'                            ' --> revised 20070122
'
'                            lWith13Mo = True
''                            If (aPayInfo(0) + aPayInfo(3) + aPayInfo(5) + aPayInfo(13)) = 0 Then
''                                lWith13Mo = False
''                            Else
''                                lWith13Mo = True
''                            End If
'
'                            If lWith13Mo Then
'                                If (oTempADO("EMP_STAT") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)) Then
'
'                                    n13mopay = Round((oTempADO("YTD_BASIC") + aPayInfo(9)) / 12, 2)
'                                ElseIf oTempADO("EMP_STAT") = 2 Then
'
'                                    n13mopay = Round((oTempADO("YTD_GROSS") + oTempADO("YTD_GROSS_SA") - oTempADO("YTD_COLA") + aPayInfo(10) + aPayInfo(11) - (oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2)))) / 12, 2)
'                                Else
'                                    n13mopay = 0
'                                End If
'                            End If
'                        Else
'                            n13mopay = 0
'                        End If
'                    End If

                    If nFilter <> 0 Then
                        
                    End If
                    
                    
                    cSqlStmt = " insert into tmpDaiLCost([EMPID],[ACTIVE],[EMP_STAT],[RATE_AMT],[COLA_AMT],[POS_ALLOW],[REG_DAY],[REG_PAY],[REG_OT_HR],[REG_OT_PAY],[NDIFF_DAY],[NDIFF_PAY],[NDIFF_OT_HR],[NDIFF_OT_PAY],[HOLIDAY],[HOL_PAY],[SA_REG_OT],[SA_REG_PAY],[SA_NDIFF_OT],[SA_NDIFF_PAY],[SUN_HR],[SUN_PAY],[SUN_OT],[SUN_OT_PAY], " & _
                               " [ADJ_PAY],[SA_ADJ_PAY],[OTHER_PAY],[LEAVE_PAY],[M13PAY],[GROSS_PAY],[SA_NET_PAY],[PAYSTATUS],[FIRSTNAME],[MNAME],[LASTNAME],[WAP],[date_res],[DATE_HIRE],[SEQ_NO],[COLA],[SUN_COLA],[SUN_ND],[SUN_ND_PAY],[SUN_ND_OT],[SUN_ND_OT_PAY],[BACCNTNO],[date_start],[date_end],[DEPNAME],[fullname])values( " & _
                                cQuote & oTempADO("EMPID") & cQuote & "," & oTempADO("ACTIVE") & "," & oTempADO("EMP_STAT") & "," & IIf(Check4.Value = vbUnchecked, oTempADO("manpower"), oTempADO("RATE_AMT")) & "," & oTempADO("COLA_AMT") & "," & IIf(Round(aTimeInfo(0), 2) + Round(aTimeInfo(3), 2) + Round(aTimeInfo(5), 2) > 0, oTempADO("POS_ALLOW"), 0) & "," & Round(aTimeInfo(0), 2) & "," & _
                                aPayInfo(0) & "," & Round(aTimeInfo(1), 2) & "," & aPayInfo(1) & "," & Round(aTimeInfo(3), 2) & "," & aPayInfo(3) & "," & Round(aTimeInfo(4), 2) & "," & aPayInfo(4) & "," & Round(aTimeInfo(7), 2) & "," & aPayInfo(7) & "," & Round(aTimeInfo(2), 2) & "," & aPayInfo(2) & "," & Round(aTimeInfo(12), 2) & "," & aPayInfo(12) & "," & _
                                Round(aTimeInfo(5), 2) & "," & aPayInfo(5) & "," & Round(aTimeInfo(6), 2) & "," & aPayInfo(6) & "," & aAdjustment(0) & "," & aAdjustment(1) & "," & aPayInfo(20) & "," & nIncentive & "," & n13mopay & "," & aPayInfo(10) + n13mopay & "," & aPayInfo(11) & "," & _
                                oTempADO("PAYSTATUS") & "," & cQuote & DecodeStr(oTempADO("FIRSTNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("MNAME")) & cQuote & "," & cQuote & oTempADO("LASTNAME") & cQuote & "," & oTempADO("WAP") & "," & _
                                cQuote & Format(IIf((oTempADO("active") = 1) Or (oTempADO("active") = 3), oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & nCtr & "," & _
                                Round((oTempADO("COLA_AMT") * (Round(aTimeInfo(3), 2) + Round(aTimeInfo(0), 2))), 2) & "," & Round(oTempADO("COLA_AMT") * ((Round(aTimeInfo(5), 2) + Round(aTimeInfo(13), 2)) / 8), 2) & "," & Round(aTimeInfo(13), 2) & "," & aPayInfo(13) & "," & Round(aTimeInfo(14), 2) & "," & aPayInfo(14) & "," & cQuote & oTempADO("BACCNTNO") & cQuote & "," & _
                                cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & Format(DTPicker2.Value, "yyyy-mm-dd") & cQuote & "," & cQuote & DecodeStr(oRecordSet("LINENAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("Fullname")) & cQuote & ")"
                    
        
'                    MsgBox cSqlStmt
'                    Script2File cSqlStmt
                    QueryTemp cSqlStmt, objdbRs, True
                        
                    
                    oTempADO.MoveNext
                Wend
            End If
            
            'd2 ang insert
            
                    
            oRecordSet.MoveNext
        Wend

        ShowProgress 3

        Select Case nFilter
        
               
            Case 0
               
                    GenerateReport "Daily Labor Cost Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvDailCost_d.rpt"
               
            Case 1
            
                    GenerateReport "Daily Labor Cost Summary Report for " & Format(DTPicker1.Value, "mmm d, yyyy") & " up to " & Format(DTPicker2.Value, "mmm d, yyyy"), "prvDailCost_s.rpt"
        
        End Select


        ShowProgress 4
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub
Private Sub Check1_Click()
    DTPicker2.Visible = Check1.Value = vbChecked
    Label1.Visible = Check1.Value = vbChecked
End Sub

Private Sub Check2_Click()
    Dim nCtr As Integer
    
    ListView1.Enabled = Check2.Value <> 1
    For nCtr = 1 To ListView1.ListItems.Count
        ListView1.ListItems(nCtr).Checked = Check2.Value = vbChecked
    Next nCtr

End Sub

Private Sub Command11_Click()
    Unload Me
End Sub

Private Sub Command13_Click()
    cmdClick Text6, Label6
    Text6.SetFocus
End Sub

Private Sub Command14_Click()
    cmdClick Text5, Label4
    Text5.SetFocus
End Sub

Private Sub Command15_Click()
    cmdClick Text7, Label15
    Text7.SetFocus
End Sub

Private Sub Command16_Click()
    cmdClick Text8, Label16
    Text8.SetFocus
End Sub

Private Sub Command5_Click()
    cmdClick Text1, Label3
    Text1.SetFocus
End Sub

Private Sub Command6_Click()
    Dim cParam As String, _
        nCtr As Integer
        
    If Tag <> 3 Then
    
        If Combo1.ListIndex = -1 Then
            MsgBox "Please select type of report to generate!", vbInformation, App.Title
            Combo1.SetFocus
            Exit Sub
        End If
    End If
    
    For nCtr = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(nCtr).Checked Then cParam = cParam & cQuote & ListView1.ListItems(nCtr).Text & cQuote & ","
    Next nCtr
    
    If Trim(cParam) <> "" Then cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
    
    Select Case Tag
        Case 1      ' -->
            GenManpower Combo1.ListIndex, cParam
        Case 2
            GenLateAbsent Combo1.ListIndex, cParam
            
        Case 3
            GenEmpOt cParam
            
        Case 4   ' --> Daily Time Consumption Report 20080520
            GenDTCons cParam
            
        Case 5   ' --> actual manpower
            GenActMan Combo1.ListIndex, cParam
            
        Case 6   ' --> actual manpower
            GenActWLCost cParam, Combo1.ListIndex
            
        Case 7
            GenWorkCenter cParam
            
        Case 8
            GenTMSERP cParam
            
        Case 9   ' --> Daily Labor
        
            GenDLCost cParam, Combo1.ListIndex
            
        Case 10  ' --> Weekly Consumption Report (IAY) 2016-10-27)
            GenWklyCon cParam
            
            
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim cParam As String, _
        nCtr As Integer
    
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    Combo1.ListIndex = 0
    SSTab1.Tab = 0
    
    Check1_Click
    
    'revise 20080707
    'OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINENAME", objdbRs, False
    
    OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 " & _
        IIf(gDepid <> "", IIf(gCompanyID <> "0003", "", " WHERE LINEID <> " & gDepid), "") & _
        " ORDER BY LINENAME", objdbRs, False
    add2LstBox objdbRs, ListView1, Array("LINENAME", "LINEID")
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Sub add2LstBox(ByVal oRecordSet As ADODB.Recordset, ByVal oListBox As ListView, ByVal aField As Variant)
    Dim lstItem As ListItem
    
    If oRecordSet.RecordCount > 0 Then
        oListBox.ListItems.Clear
        While Not oRecordSet.EOF
            Set lstItem = oListBox.ListItems.Add()
            lstItem.Text = objdbRs(aField(1))
            lstItem.SubItems(1) = objdbRs(aField(0))
            oRecordSet.MoveNext
        Wend
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 3, Text1.Text, Label3
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 2, Text5.Text, Label4
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text6.Text, Label6
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 4, Text7.Text, Label15
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 5, Text8.Text, Label16
End Sub
