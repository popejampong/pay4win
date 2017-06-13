VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Begin VB.Form frmEmpIncentive2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incentive Entry"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12495
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   5595
      TabIndex        =   38
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command19 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   43
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
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "TXT:APPR_BY"
         Top             =   1545
         Width           =   660
      End
      Begin VB.CommandButton Command18 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   42
         Top             =   1230
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
         Left            =   1455
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "TXT:NOTED_BY"
         Top             =   1245
         Width           =   660
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
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   945
         Width           =   660
      End
      Begin VB.CommandButton Command17 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   41
         Top             =   930
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
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "TXT:PREP_BY"
         Top             =   645
         Width           =   660
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   300
         Left            =   2145
         TabIndex        =   40
         Top             =   630
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Details >>"
         Height          =   375
         Left            =   5745
         TabIndex        =   39
         Top             =   165
         Width           =   975
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
         TabIndex        =   52
         Top             =   1590
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
         TabIndex        =   51
         Top             =   1290
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
         TabIndex        =   50
         Top             =   990
         Width           =   4215
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
         TabIndex        =   49
         Top             =   690
         Width           =   4215
      End
      Begin VB.Label Label11 
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
         TabIndex        =   48
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label10 
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
         TabIndex        =   47
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label9 
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
         TabIndex        =   46
         Top             =   1290
         Width           =   1350
      End
      Begin VB.Label Label8 
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
         TabIndex        =   45
         Top             =   1590
         Width           =   1350
      End
      Begin VB.Label Label7 
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
         TabIndex        =   44
         Top             =   690
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Import"
      Height          =   405
      Left            =   3105
      TabIndex        =   31
      Top             =   45
      Width           =   1140
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
      Left            =   1275
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "TXT:DEPID"
      Top             =   795
      Width           =   585
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   300
      Left            =   1890
      TabIndex        =   27
      Top             =   795
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
      Left            =   1275
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:PERIODID"
      Top             =   1110
      Width           =   585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   300
      Left            =   1890
      TabIndex        =   24
      Top             =   1110
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
      Left            =   1290
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "TXT:INC_NO"
      Top             =   180
      Width           =   1515
   End
   Begin VB.TextBox txtFlex 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2295
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   4755
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6240
      Left            =   45
      TabIndex        =   8
      Top             =   1515
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   11007
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
      Left            =   1545
      TabIndex        =   22
      Top             =   7785
      Width           =   10890
      Begin VB.CommandButton Command13 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4515
         Picture         =   "frmEmpIncentive2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "16"
         Top             =   135
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Appl&y"
         Height          =   660
         Left            =   8940
         Picture         =   "frmEmpIncentive2.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "22"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7980
         Picture         =   "frmEmpIncentive2.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   7140
         Picture         =   "frmEmpIncentive2.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   6300
         Picture         =   "frmEmpIncentive2.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2640
         Picture         =   "frmEmpIncentive2.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1800
         Picture         =   "frmEmpIncentive2.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   960
         Picture         =   "frmEmpIncentive2.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   9915
         Picture         =   "frmEmpIncentive2.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   5460
         Picture         =   "frmEmpIncentive2.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3600
         Picture         =   "frmEmpIncentive2.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   120
         Picture         =   "frmEmpIncentive2.frx":11896
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "DAT:INC_DATE"
      Top             =   480
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56688640
      CurrentDate     =   38623
   End
   Begin ciaXPPanel.XPPanel XPPanel5 
      Height          =   645
      Left            =   5595
      TabIndex        =   32
      Top             =   735
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1138
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   315
         Left            =   1845
         TabIndex        =   34
         Top             =   60
         Width           =   405
      End
      Begin VB.TextBox Text11 
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
         Left            =   1215
         TabIndex        =   33
         Tag             =   "1"
         ToolTipText     =   "TXT:SHIFTID"
         Top             =   75
         Width           =   630
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Id"
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
         Left            =   45
         TabIndex        =   37
         Top             =   105
         Width           =   1545
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Regular Shift"
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
         Left            =   2325
         TabIndex        =   36
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "08:00 AM - 05:00 PM"
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
         Left            =   1350
         TabIndex        =   35
         Top             =   390
         Width           =   3180
      End
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
      TabIndex        =   30
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Deparment"
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
      TabIndex        =   29
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2295
      TabIndex        =   28
      Top             =   855
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
      Left            =   90
      TabIndex        =   26
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2295
      TabIndex        =   25
      Top             =   1155
      Width           =   4005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Incentive No"
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
      TabIndex        =   23
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmEmpIncentive2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll
' module        :   frmEmpIncentive
' description   :   Employee Incentive  Module
' programmer    :   _-=[ srm ]=-_
' date          :   13 mar 2008

Option Explicit
    Dim nAdd As Integer, _
        nLastRow As Integer, _
        cSeries As String, _
        cParam As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant

Sub ShowData2(cString As String, oLabel As Label)
    OpenQueryDNS "SELECT USERID,CONCAT(FIRSTNAME," & cQuote & " " & cQuote & ",LASTNAME) AS FULLNAME FROM pa2360 WHERE USERID=" & cQuote & cString & cQuote, objdbRs, False
    oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("FULLNAME"), "")
End Sub

Sub txtKeyDown2(nMode As Integer, cString As String, oLabel As Label)
    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command15_Click
                Case 2
                    Command17_Click
                Case 3
                    Command18_Click
                Case 4
                    Command19_Click
            End Select
        Else
            ShowData2 cString, oLabel
        End If
    End If
End Sub


Sub cmdClick(ByVal oTxtBox As TextBox, ByVal oLabel As Label)
    frmLookup.showPopup 1   ', " where sysuser = 1"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTxtBox.Text = cResult
        ShowData2 cResult, oLabel
    End If
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmp4620( " & _
               " [INC_NO] char(10),             [INC_DATE] date," & _
               " [PERIODID] char(5),            [DATE_START] date," & _
               " [DATE_END] date,               [DURATION] char(100)," & _
               " [DEPID] char(3),               [LINENAME] char(100)," & _
               " [SHIFTID] char(5),             [DESCRIPTION] char(100)," & _
               " [TIMEDESC] char(100),          [EMPID] char(6), " & _
               " [FULLNAME] char(100),          [POSNAME] char(100), " & _
               " [EMP_STAT] char(100),          [Inc_hr] double, " & _
               " [SEQ_NO] integer,              [REMARK] char(100), " & _
               " [status] integer,              [prep_by] char(6)," & _
               " [check_by] char(6),            [note_by] char(6)," & _
               " [appr_by] char(6),             [prep_pos] char(100)," & _
               " [check_pos] char(100),         [note_pos] char(100)," & _
               " [appr_pos] char(100),          [prep_name] char(100)," & _
               " [check_name] char(100),        [note_name] char(100)," & _
               " [appr_name] char(100))"

               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmp4620", oTempADO, True
End Sub

Sub txtKeyDown(ByVal nMode As Integer, cString As String, oLabel As Label, Optional ByVal oLabel2 As Label)
    If nAdd <> 0 Then
        If Trim(cString) = "" Then
            Select Case nMode
                Case 1
                    Command3_Click
                Case 2
                    Command2_Click
            End Select
        Else
            ShowData nMode, cString, oLabel, oLabel2
        End If
    End If
End Sub

Sub ShowData(ByVal nMode As Integer, cString As String, oLabel As Label, Optional ByVal oLabel2 As Label)
    Select Case nMode
        Case 1      ' --> department
            OpenQueryDNS "SELECT linename as `description` FROM DI5463 WHERE lineid=" & cQuote & cString & cQuote, objdbRs, False
        Case 2      ' --> period
            OpenQueryDNS "SELECT duration as `description` FROM PA7730 WHERE PERIODID=" & cQuote & cString & cQuote, objdbRs, False
        Case 3
            OpenQueryDNS "SELECT `DESCRIPTION`,CONCAT(TIME_FORMAT(TIME1,'%h:%i %p'),' - ',TIME_FORMAT(TIME2,'%h:%i %p')) AS `TIME` FROM PA74380 WHERE SHIFTID=" & cQuote & cString & cQuote, objdbRs, False
    End Select
    If nMode <> 3 Then
        oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("description"), "")
    Else
        oLabel.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("description"), "")
        oLabel2.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TIME"), "")
    End If
End Sub

'Sub InsertToGrid cResult, nRow, MSHFlexGrid1
Sub InsertToGrid(ByVal cString As String, ByVal nRowPos As Integer, ByVal oFlexGrid As MSHFlexGrid)
    Dim cSqlStmt As String
    With MSHFlexGrid1
        If Trim(cString) <> "" Then
            .TextMatrix(nRowPos, 1) = cString
            
            cSqlStmt = "select a.empid, " & _
                       "  replace(concat(a.firstname,' ',a.lastname),CHAR(22),'" & cQuote & "') as fullname, " & _
                       "  replace(ifnull(b.posname,''),CHAR(22),'" & cQuote & "') as posname, " & _
                       "  if(a.emp_stat = 0,'Wap',if(a.emp_stat=1,'Contractual','Regular')) as emp_stat " & _
                       "from di3670 a " & _
                       "  left join di7670 b on a.posid=b.posid " & _
                       "where a.empid=" & cQuote & cString & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nRowPos, 2) = objdbRs("fullname")
                .TextMatrix(nRowPos, 3) = objdbRs("posname")
                .TextMatrix(nRowPos, 4) = objdbRs("emp_stat")
            End If
        End If
    End With
End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then CtrlPanel Me, nAdd, oTempADO("STATUS") <> 1

    ShowData 1, Text3.Text, Label2
    ShowData 2, Text2.Text, Label4
    ShowData 3, Text11.Text, Label32, Label33
    
    Command3.Enabled = nAdd <> 0
    Command2.Enabled = nAdd <> 0
    
    Command6.Enabled = nAdd <> 0
    
    ShowData2 Text5.Text, Label15
    ShowData2 Text7.Text, Label17
    ShowData2 Text8.Text, Label18
    ShowData2 Text9.Text, Label19
    
    Command15.Enabled = nAdd <> 0
    Command17.Enabled = nAdd <> 0
    Command18.Enabled = nAdd <> 0
    Command19.Enabled = nAdd <> 0

    Command12.Enabled = nAdd <> 0
    

    cSqlStmt = " SELECT a.empid, " & _
               " concat(b.firstname,' ',b.lastname) as fullanme, " & _
               " replace(ifnull(c.posname,''),CHAR(22),'" & cQuote & "') as posname, " & _
               " if(b.emp_stat = 0,'Wap',if(b.emp_stat=1,'Contractual','Regular')) as emp_stat, " & _
               " a.Inc_hr, " & _
               " replace(ifnull(a.REMARK,''),CHAR(22),'" & cQuote & "') as REMARK, " & _
               " a.SEQ_NO,a.status " & _
               " FROM pa4623 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di7670 c on b.posid=c.posid " & _
               " where a.inc_no=" & cQuote & Text1.Text & cQuote & _
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
    If Not ChkPersonnel(Text7) Then Exit Sub
    If Not ChkPersonnel(Text8) Then Exit Sub
    If Not ChkPersonnel(Text9) Then Exit Sub
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Incentive entry?", vbYesNoCancel, "Incentive Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA4620", "INC_NO=" & cQuote & Text1.Text & cQuote) Then
                    MsgBox "Incentive Reference Number already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA4620"), oTempADO, True
                    Script2File InsertFields(Me, "PA4620")
                    
                    Log2Audit Name, "ADD INC_NO -->" & Trim(Text1.Text)
                
                    ShowProgress 0
                    
                    With MSHFlexGrid1
                        For nCtr = 1 To .Rows - 1
                        
                            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                            
                            If Trim(.TextMatrix(nCtr, 5)) <> "" And Trim(.TextMatrix(nCtr, 5)) <> 0 Then
                                cSqlStmt = "insert into PA4623(INC_NO,INC_DATE,EMPID,Inc_hr,SEQ_NO,REMARK)values(" & _
                                           cQuote & Text1.Text & cQuote & "," & _
                                           cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                           Val(.TextMatrix(nCtr, 5)) & "," & _
                                           nCtr & "," & _
                                           cQuote & .TextMatrix(nCtr, 6) & cQuote & ")"
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                            
                        Next nCtr
                    End With
                    
                    ShowProgress 4
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "PA4620", "INC_NO=" & cQuote & Text1.Text & cQuote), oTempADO, True
                Script2File EditField(Me, "PA4620", "INC_NO=" & cQuote & Text1.Text & cQuote)
                
                Log2Audit Name, "EDIT INC_NO -->" & Trim(Text1.Text)
            
                cSqlStmt = "delete from PA4623 where INC_NO=" & cQuote & Text1.Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 0
                
                With MSHFlexGrid1
                    For nCtr = 1 To .Rows - 1
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                        
                        If Trim(.TextMatrix(nCtr, 5)) <> "" And Trim(.TextMatrix(nCtr, 5)) <> 0 Then
                            cSqlStmt = "insert into PA4623(INC_NO,INC_DATE,EMPID,Inc_hr,SEQ_NO,REMARK)values(" & _
                                       cQuote & Text1.Text & cQuote & "," & _
                                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                       Val(.TextMatrix(nCtr, 5)) & "," & _
                                       nCtr & "," & _
                                       cQuote & .TextMatrix(nCtr, 6) & cQuote & ")"
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
    
    If Text1.Text <> cSeries Then ResetSeries "INCENTIVE", cSeries
    
    Frame1.Height = 615
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "INC_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
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
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
            
            If Text1.Text <> cSeries Then ResetSeries "INCENTIVE", cSeries
            
            Frame1.Height = 615
            nAdd = 0
            
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "INC_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            
            ShowRecords
        End If
    End If
End Sub


Private Sub Command12_Click()
    frmLookup.showPopup 9
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text11.Text = cResult
        ShowData 3, cResult, Label32, Label33
'        OpenQueryDNS "SELECT `DESCRIPTION`,CONCAT(TIME_FORMAT(TIME1,'%h:%i %p'),' - ',TIME_FORMAT(TIME2,'%h:%i %p')) AS `TIME` FROM PA74380 WHERE SHIFTID=" & cQuote & cResult & cQuote, objdbRs, False
'        Label32.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
'        Label33.Caption = IIf(objdbRs.RecordCount > 0, objdbRs("TIME"), "")
    End If
End Sub

Private Sub Command13_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        aUserInfo As Variant


    aUserInfo = Array("", "", "", "", "", "")
    
    CreateTemp

    ' --> for user info
    OpenQueryDNS "SELECT * FROM DI2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aUserInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text8.Text & "'"
        aUserInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text9.Text & "'"
        aUserInfo(3) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If
    '--> END USER

    CreateTemp
    
    With MSHFlexGrid1
    
        ShowProgress 0
        
        For nCtr = 1 To (.Rows - 1)
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            cSqlStmt = " insert into tmp4620(INC_NO,INC_DATE,[PERIODID],[DURATION],DEPID,LINENAME,SHIFTID,[DESCRIPTION],TIMEDESC,EMPID,FULLNAME,POSNAME,EMP_STAT,Inc_hr,SEQ_NO, " & _
                       " REMARK,prep_by,check_by,note_by,appr_by,prep_pos,check_pos,note_pos,appr_pos,prep_name,check_name,note_name,appr_name)values(" & _
                       cQuote & Text1.Text & cQuote & "," & _
                       cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Text2.Text & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & _
                       cQuote & Text3.Text & cQuote & "," & cQuote & Label2.Caption & cQuote & "," & _
                       cQuote & Text11.Text & cQuote & "," & cQuote & Label32.Caption & cQuote & "," & _
                       cQuote & Label33.Caption & cQuote & "," & cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 2))) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 3))) & cQuote & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 4))) & cQuote & "," & _
                       .TextMatrix(nCtr, 5) & "," & nCtr & "," & _
                       cQuote & EncodeStr(DecodeStr(.TextMatrix(nCtr, 6))) & cQuote & "," & _
                       cQuote & Text5.Text & cQuote & "," & cQuote & Text7.Text & cQuote & "," & _
                       cQuote & Text8.Text & cQuote & "," & cQuote & Text9.Text & cQuote & "," & _
                       cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & _
                       cQuote & aUserInfo(2) & cQuote & "," & cQuote & aUserInfo(3) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label15.Caption)) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label17.Caption)) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label18.Caption)) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label19.Caption)) & cQuote & ")"
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
        Next nCtr
        
        ShowProgress 3
        
        GenerateReport "Incentive Report", "PRV4620.RPT"
        
        ShowProgress 4
        
    End With

End Sub

Private Sub Command14_Click()
    Frame1.Height = IIf(Frame1.Height = 615, 2220, 615)
    Command14.Caption = IIf(Frame1.Height = 615, "Detail >>", "<< Hide")
End Sub

Private Sub Command15_Click()
    cmdClick Text5, Label15
    Text7.SetFocus
End Sub

Private Sub Command17_Click()
    cmdClick Text7, Label17
    Text8.SetFocus
End Sub

Private Sub Command18_Click()
    cmdClick Text8, Label18
    Text9.SetFocus
End Sub

Private Sub Command19_Click()
    cmdClick Text9, Label19
    Text9.SetFocus
End Sub

Private Sub Command2_Click()
    frmLookup.showPopup 5, " where (pclose=0) and (isprocess=0)"
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text2.Text = cResult
        ShowData 2, cResult, Label4
    End If
    MSHFlexGrid1.SetFocus
End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 2
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text3.Text = cResult
        ShowData 1, cResult, Label2
    End If
    Text2.SetFocus
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
        If MsgBox("Apply this Incentive entry?", vbYesNo, App.Title) = vbYes Then
        
            cString = Text1.Text
            
            ShowProgress 0
            
            With MSHFlexGrid1
                For nCtr = 1 To .Rows - 1
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    
                    If Val(.TextMatrix(nCtr, 8)) <> 1 Then
                        
                        cSqlStmt = "update PA4623 set status=1, " & _
                                   " date_post=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                                   " where INC_NO=" & cQuote & Text1.Text & cQuote & _
                                   " and seq_no=" & .TextMatrix(nCtr, 7)
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                    End If
                    
                    ' update employee 36770
                        cSqlStmt = "update di36770 set inc_hr=" & cQuote & .TextMatrix(nCtr, 5) & cQuote & _
                                   " where periodid = " & cQuote & Text2.Text & cQuote & _
                                   " and empid = " & cQuote & .TextMatrix(nCtr, 1) & cQuote & _
                                   " and date =" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote
                                   
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                        
                    nCount = nCount + 1
                    
                Next nCtr
                
                If nCount = .Rows - 1 Then
                    cSqlStmt = "update PA4620 set status=1," & _
                               " date_post=" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
                               " where INC_NO=" & cQuote & Text1.Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
                End If
                
                ShowProgress 4
                
                oTempADO.Requery adAsyncFetch
                If Trim(cString) <> "" Then oTempADO.Find "INC_NO='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
                GetFields Me, oTempADO
            
                
                ShowRecords
                
            End With
        End If
    Else
        cString = "Warning!" & vbCrLf & "You do not have permission to apply this Incentive entry!" & vbCrLf & vbCrLf & _
                  "Please contact your supervisor or your System Administrator for more information..."
        MsgBox cString, vbCritical, App.Title
    End If
    
    Exit Sub
    
ErrApply:
    ErrorMsg Err.Number, Err.Description, "Apply Incentive #" & Text1.Text, Name
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 16
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "INC_NO='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        nCtr As Integer, _
        dStart_date As String, _
        dEnd_date As String
    
    If Text2.Text = "" Then Exit Sub
    
    SetGridColumn myArray, MSHFlexGrid1
    
    OpenQueryDNS " select date_format(date_start,'%Y-%m-%d') as date_start, " & _
                 " date_format(date_end,'%Y-%m-%d') as date_end from pa7730 where periodid = " & cQuote & Text2.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        dStart_date = objdbRs("date_start")
        dEnd_date = objdbRs("date_end")
    End If
    
    cSqlStmt = " SELECT a.EMPID, concat(a.FIRSTNAME, ' ', a.LASTNAME) as fullname, " & _
               " b.posname, if(a.EMP_STAT=0,'Wap',if(a.EMP_STAT=1,'Contractual','Regular')) as emp_stat, a.depid " & _
               " FROM di3670 a " & _
               " left join di7670 b on a.posid=b.posid " & _
               " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & dStart_date & cQuote & " and " & cQuote & dEnd_date & cQuote & ") or ((a.date_hire<=" & cQuote & dEnd_date & cQuote & ") and (a.date_res > " & cQuote & dEnd_date & cQuote & ")))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & dStart_date & cQuote & " and " & cQuote & dEnd_date & cQuote & ") or ((a.date_hire<=" & cQuote & dEnd_date & cQuote & ") and (a.date_fin > " & cQuote & dEnd_date & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & dEnd_date & cQuote & ")) " & _
               " order by a.emp_stat desc, fullname "
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        ShowProgress 0
        While Not oRecordSet.EOF
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("empid")
                                                        
            If oRecordSet("depid") = Text3.Text Then
                With MSHFlexGrid1
                
                    If Trim(.TextMatrix(.Rows - 1, 1)) <> "" Then
                        .AddItem "", .Rows
                    End If
                
                    .RowSel = .Rows - 1
                    
                    .TextMatrix(.RowSel, 1) = oRecordSet("empid")
                    .TextMatrix(.RowSel, 2) = oRecordSet("fullname")
                    .TextMatrix(.RowSel, 3) = oRecordSet("posname")
                    .TextMatrix(.RowSel, 4) = oRecordSet("emp_stat")
                End With
            End If
                            
            
            oRecordSet.MoveNext
        Wend
        ShowProgress 4
    End If
               
End Sub

Private Sub Command7_Click()
    SetGridColumn myArray, MSHFlexGrid1
    
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command3.Enabled = True
    Command2.Enabled = True
    Command6.Enabled = True
    
    Label2.Caption = ""
    Label4.Caption = ""
    
    Command15.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = True
    Command18.Enabled = True
    Command19.Enabled = True
    
    Command12.Enabled = True
    
    Label15.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    
    DTPicker1.Value = Now
    
    cSeries = GenerateSeries("INCENTIVE")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA4620", "PA4620.INC_NO=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("INCENTIVE")
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
        
        Command3.Enabled = True
        Command2.Enabled = True
        Command6.Enabled = True
        
        
        Command15.Enabled = True
        Command17.Enabled = True
        Command18.Enabled = True
        Command19.Enabled = True
        
        Command12.Enabled = True
        
        Text1.Enabled = False
        DTPicker1.SetFocus
    End If


End Sub

Private Sub Command9_Click()
    On Error GoTo ErrDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, App.Title) = vbYes Then
        OpenQueryDNS "DELETE FROM PA4620 WHERE INC_NO=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM PA4620 WHERE INC_NO=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE " & Trim(Text1.Text) & "-" & Trim(EncodeStr2(DecodeStr(Text2.Text)))
        
        OpenQueryDNS "DELETE FROM PA4623 WHERE INC_NO=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM PA4623 WHERE INC_NO=" & cQuote & Text1.Text & cQuote
        
        Lock2User Me.Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050321
        
        nAdd = 0
        ClearAll Me, False, True
        CtrlPanel Me, nAdd
        
'        OpenQueryDNS "SELECT * FROM PA74380 ORDER BY SHIFTID", oTempADO, False
        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        
        ShowRecords
    End If
Exit Sub
ErrDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Emp ID]:8:True", _
                    "TXT:[Full Name]:50:True", _
                    "TXT:[Position]:30:True", _
                    "TXT:[Status]:30:True", _
                    "NUM:[Incentive]:10:True", _
                    "TXT:[Remark]:50:True", _
                    "NUM:[Seq No]:2:False", _
                    "NUM:[Status]:1:True")
                    
    Tag = nAccess_Tag
    nAdd = 0
    Frame1.Height = 615
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
        
    OpenQueryDNS "SELECT * FROM PA4620 ORDER BY INC_NO", oTempADO, False
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
                    Case 1, 5, 6
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
        KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "txtFlex")
    End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 3, Text11.Text, Label32, Label33
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 2, Text2.Text, Label4
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text3.Text, Label2
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown2 1, Text5.Text, Label15
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown2 2, Text7.Text, Label17
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown2 3, Text8.Text, Label18
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown2 4, Text9.Text, Label19
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
                                       " ((a.active=1) and (a.date_res =" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) or " & _
                                       " ((a.active=2) and (a.date_fin =" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))" & _
                                       IIf(Trim(cParam) = "", "", " and a.empid not in " & cParam) & " and depid = " & cQuote & Text3.Text & cQuote
'                            MsgBox cSqlStmt
                            frmLookup.showPopup 3, cSqlStmt
                            frmLookup.Show 1
                            If Trim(cResult) <> "" Then InsertToGrid cResult, .Row, MSHFlexGrid1
                        Else
                            nCtr = .Row
                            If Not ChkDupInGrid(txtFlex.Text, 1, MSHFlexGrid1) Then
                                cSqlStmt = "select a.empid from di3670 a " & _
                                           " WHERE ((a.ACTIVE=0) or " & _
                                           " ((a.active=1) and (a.date_res =" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")) or " & _
                                           " ((a.active=2) and (a.date_fin =" & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))" & _
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
                        .Col = 5
                        
                    Case 5
                        .TextMatrix(.Row, 5) = txtFlex.Text
                    Case 6
                        .TextMatrix(.Row, 6) = txtFlex.Text
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

