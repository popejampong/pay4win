VERSION 5.00
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaxppanel.ocx"
Begin VB.Form frmBlockList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Block List"
   ClientHeight    =   5415
   ClientLeft      =   7500
   ClientTop       =   2355
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9855
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   315
      Left            =   4140
      TabIndex        =   33
      Top             =   615
      Width           =   495
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
      Left            =   1890
      TabIndex        =   32
      Tag             =   "1"
      ToolTipText     =   "TXT:SSNUM"
      Top             =   645
      Width           =   2235
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
      Left            =   1890
      TabIndex        =   31
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   345
      Width           =   1200
   End
   Begin VB.CommandButton Command12 
      Caption         =   "..."
      Height          =   315
      Left            =   3120
      TabIndex        =   30
      Top             =   315
      Width           =   495
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
      Left            =   1890
      TabIndex        =   26
      Tag             =   "1"
      ToolTipText     =   "TXT:BLKID"
      Top             =   45
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   0
      TabIndex        =   14
      Top             =   4575
      Width           =   9855
      Begin VB.CommandButton Command13 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4485
         Picture         =   "frmBlockList.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "16"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmBlockList.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmBlockList.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         Height          =   660
         Left            =   5385
         Picture         =   "frmBlockList.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "17"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   8865
         Picture         =   "frmBlockList.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmBlockList.frx":7F8A
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmBlockList.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmBlockList.frx":B28E
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   6225
         Picture         =   "frmBlockList.frx":CC10
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Delete"
         Height          =   660
         Left            =   7065
         Picture         =   "frmBlockList.frx":E592
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "19"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   7905
         Picture         =   "frmBlockList.frx":FF14
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.TextBox Text33 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   1890
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "TXT:S_REMARK"
      Top             =   3300
      Width           =   5100
   End
   Begin ciaXPPanel.XPPanel XPPanel8 
      Height          =   2325
      Left            =   1875
      TabIndex        =   34
      Top             =   945
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   4101
      LicValid        =   -1  'True
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   43
         Top             =   1260
         Width           =   7815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   42
         Top             =   75
         Width           =   7815
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   41
         Top             =   300
         Width           =   7815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   40
         Top             =   540
         Width           =   7815
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   39
         Top             =   780
         Width           =   7815
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   38
         Top             =   1020
         Width           =   7815
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   37
         Top             =   1500
         Width           =   7815
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   36
         Top             =   1755
         Width           =   7815
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   35
         Top             =   1995
         Width           =   7815
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   825
      Left            =   6210
      TabIndex        =   44
      Top             =   90
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   1455
      LicValid        =   -1  'True
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmBlockList.frx":11896
         Left            =   90
         List            =   "frmBlockList.frx":118A3
         TabIndex        =   45
         Text            =   "Combo1"
         Top             =   360
         Width           =   3180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Option"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   90
         TabIndex        =   46
         Top             =   90
         Width           =   3105
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      Height          =   390
      Left            =   45
      TabIndex        =   29
      Top             =   2925
      Width           =   1680
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Block List ID"
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
      Index           =   2
      Left            =   90
      TabIndex        =   27
      Top             =   90
      Width           =   1590
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
      Left            =   150
      TabIndex        =   25
      ToolTipText     =   "TXT:LINENAME"
      Top             =   360
      Width           =   1590
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   90
      TabIndex        =   13
      Top             =   3255
      Width           =   1590
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Status"
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
      Left            =   135
      TabIndex        =   12
      Top             =   2670
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Left            =   150
      TabIndex        =   11
      Top             =   2430
      Width           =   1590
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      Left            =   150
      TabIndex        =   10
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday"
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
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1470
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Hire"
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
      Height          =   285
      Left            =   165
      TabIndex        =   8
      Top             =   1005
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fullname"
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
      Left            =   150
      TabIndex        =   7
      Top             =   1230
      Width           =   1590
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      TabIndex        =   6
      Top             =   1725
      Width           =   1590
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No / CP No"
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
      Left            =   135
      TabIndex        =   5
      Top             =   1950
      Width           =   1590
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Number"
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   675
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   7755
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   4815
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Empl. Status"
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
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Status"
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
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "frmBlockList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim nAdd As Integer, _
        cSeries As String, _
        cParam As String, _
        oTempADO As New ADODB.Recordset, _
        myArray As Variant, _
        cBLKID As String

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
        dbNavigator Screen.ActiveControl, Me, oTempADO
        ShowRecords
    End If
End Sub
Sub ShowRecords()
    Dim cSqlStmt As String
    Dim sAddress As String
    Dim sLinename As String
    Dim sPosname As String
    Dim oRSet As New ADODB.Recordset
    
    Command12.Enabled = nAdd <> 0
    Command4.Enabled = nAdd <> 0
    
    Combo1.ListIndex = 0
    Combo1.Enabled = nAdd <> 0
    
    cSqlStmt = " select a.EMPID,a.SSNUM,b.DATE_HIRE,ifnull(concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'. '),'') as fullname, " & _
               " b.BIRTHDAY,b.ADD_NO,b.ADD_BRGY,b.ADD_CITY, " & _
               " ifnull(b.TEL_NUM,'') as TEL_NUM,b.DEPID,b.POSID, if(b.EMP_STAT=0,'Wap',if(b.EMP_STAT=1,'Conmtractual','Regular')) as emp_stat, " & _
               " ifnull(concat(if(b.ACTIVE=0,'Active',if(b.ACTIVE=1,'Resigned',if(b.ACTIVE=2,'Finished','Terminated'))), ' - ' , b.DATE_RES ),'') as date_Term, a.s_remark " & _
               " from PA255578 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " Where a.blkid = " & cQuote & Text1.Text & cQuote & " and a.empid = " & cQuote & Text2.Text & cQuote

    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then

        OpenQueryDNS "select * from di5463 where lineid = " & cQuote & objdbRs("depid") & cQuote, oRSet, False
        sLinename = IIf(oRSet.RecordCount > 0, oRSet("linename"), "")

        OpenQueryDNS "select * from di7670 where posid = " & cQuote & objdbRs("posid") & cQuote, oRSet, False
        sPosname = IIf(oRSet.RecordCount > 0, oRSet("posname"), "")

        Text3.Text = objdbRs("SSNUM")
        Label14.Caption = Format(objdbRs("DATE_HIRE"), "yyyy-mm-dd")
        Label18.Caption = objdbRs("fullname")
        Label21.Caption = Format(objdbRs("BIRTHDAY"), "yyyy-mm-dd")
        
        sAddress = IIf(objdbRs("ADD_NO") = "", "", objdbRs("ADD_NO") & " ")
        sAddress = sAddress & IIf(objdbRs("ADD_BRGY") = "", "", objdbRs("ADD_BRGY") & " ")
        sAddress = sAddress & IIf(objdbRs("ADD_CITY") = "", "", objdbRs("ADD_CITY") & " ")
        Label22.Caption = IIf(sAddress = "", "Bataan", sAddress & "Bataan")
        
        
        Label23.Caption = objdbRs("TEL_NUM")
        Label10.Caption = sLinename
        Label24.Caption = sPosname
        Label25.Caption = objdbRs("emp_stat")
        Label26.Caption = objdbRs("date_Term")
        Text33.Text = objdbRs("s_remark")
    Else
        Text33.Text = ""
        Label14.Caption = ""
        Label18.Caption = ""
        Label21.Caption = ""
        Label22.Caption = ""
        Label23.Caption = ""
        Label10.Caption = ""
        Label24.Caption = ""
        Label25.Caption = ""
        Label26.Caption = ""
        Text33.Text = ""
    End If
End Sub
Private Sub Command10_Click()
    On Error GoTo ErrBLKENTRYSave
    Dim cString As String
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Black List file entry?", vbYesNoCancel, "Black List file entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("PA255578", "BLKID='" & Text1.Text & "'") Then
                    MsgBox "Company ID already exists!", vbOKOnly, App.Title
                    Text1.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "PA255578"), oTempADO, True
                    Script2File InsertFields(Me, "PA255578")
                    
                    Log2Audit Name, "ADD " & "BLK ID -->" & Trim(Text1.Text)
                    Log2Audit Name, "ADD " & "EMPID NAME -->" & Trim(Text2.Text)
                    Log2Audit Name, "ADD " & "SSS NO -->" & Trim(Text3.Text)
                    
                End If
            Else
                OpenQueryDNS EditField(Me, "PA255578", "PA255578.BLKID='" & Text1.Text & "'"), oTempADO, True
                Script2File EditField(Me, "PA255578", "PA255578.BLKID='" & Text1.Text & "'")
                
                Log2Audit Name, "EDIT " & "BLK ID -->" & Trim(Text1.Text)
                Log2Audit Name, "EDIT " & "EMPID -->" & Trim(Text2.Text)
                Log2Audit Name, "EDIT " & "SSS NO -->" & Trim(Text3.Text)
            End If
        Case vbCancel
            GoTo endsave
    End Select

    Lock2User Name, Text1.ToolTipText, Text1.Text, False     ' --> 20050322
    
    If Text1.Text <> cSeries Then ResetSeries "BLKID", cSeries
    
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "BLKID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    GetFields Me, oTempADO
    ShowRecords
    
endsave:
    Exit Sub
    
ErrBLKENTRYSave:
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
            
            If Text1.Text <> cSeries Then ResetSeries "BLKID", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "BLKID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command12_Click()
    Dim cSqlStmt As String, _
        sLinename As String, _
        sPosname As String, _
        sActive As String, _
        sAddress As String
        
    Dim oRSet As New ADODB.Recordset
    
    Frame2.Enabled = False
    
     sActive = IIf(Combo1.ListIndex = 0, " where a.active = 0 ", IIf(Combo1.ListIndex = 1, " where a.active <> 0 ", ""))
    
    frmLookup.showPopup 24, sActive

    frmLookup.Show 1
    
    If Trim(cResult) <> "" Then
        
        cSqlStmt = " select EMPID,SSNUM,DATE_HIRE,concat(LASTNAME,', ',FIRSTNAME,' ', left(MNAME,1),'. ') as fullname, " & _
                   " BIRTHDAY,ADD_NO,ADD_BRGY,ADD_CITY, " & _
                   " TEL_NUM,DEPID,POSID, if(EMP_STAT=0,'Wap',if(EMP_STAT=1,'Conmtractual','Regular')) as emp_stat, " & _
                   " concat(if(ACTIVE=0,'Active',if(ACTIVE=1,'Resigned',if(ACTIVE=2,'Finished','Terminated'))), ' - ' , DATE_RES ) as date_Term,S_REMARK " & _
                   " from di3670 " & _
                   " Where empid = " & cQuote & cResult & cQuote
               
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
        
            OpenQueryDNS "select * from di5463 where lineid = " & cQuote & objdbRs("depid") & cQuote, oRSet, False
            sLinename = IIf(oRSet.RecordCount > 0, oRSet("linename"), "")
            
            OpenQueryDNS "select * from di7670 where posid = " & cQuote & objdbRs("posid") & cQuote, oRSet, False
            sPosname = IIf(oRSet.RecordCount > 0, oRSet("posname"), "")
            
            Text2.Text = cResult
            Text3.Text = objdbRs("SSNUM")
            Label14.Caption = Format(objdbRs("DATE_HIRE"), "yyyy-mm-dd")
            Label18.Caption = objdbRs("fullname")
            Label21.Caption = Format(objdbRs("BIRTHDAY"), "yyyy-mm-dd")
            
            sAddress = IIf(objdbRs("ADD_NO") = "", "", objdbRs("ADD_NO") & " ")
            sAddress = sAddress & IIf(objdbRs("ADD_BRGY") = "", "", objdbRs("ADD_BRGY") & " ")
            sAddress = sAddress & IIf(objdbRs("ADD_CITY") = "", "", objdbRs("ADD_CITY") & " ")
            Label22.Caption = IIf(sAddress = "", "Bataan", sAddress & "Bataan")
    
            
            Label23.Caption = objdbRs("TEL_NUM")
            Label10.Caption = sLinename
            Label24.Caption = sPosname
            Label25.Caption = objdbRs("emp_stat")
            Label26.Caption = objdbRs("date_Term")
            Text33.Text = objdbRs("s_remark")
            
        Else
            Text2.Text = ""
            Text3.Text = ""
            Label14.Caption = ""
            Label18.Caption = ""
            Label21.Caption = ""
            Label22.Caption = ""
            Label23.Caption = ""
            Label10.Caption = ""
            Label24.Caption = ""
            Label25.Caption = ""
            Label26.Caption = ""
            Text33.Text = ""
        End If
                
    End If
    Frame2.Enabled = True

End Sub

Private Sub createtmpBlockprev()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE tmpBlockprev ([blkid] char(15), " & _
                     " [empid] char(15), " & _
                     " [ssnum] char(50), " & _
                     " [date_hire] date, " & _
                     " [fullname] char(100)," & _
                     " [birthday] date," & _
                     " [address] char(100)," & _
                     " [tel_num] char(100)," & _
                     " [linename] char(100)," & _
                     " [posname] char(100)," & _
                     " [active] char(100)," & _
                     " [emp_stat] char(100)," & _
                     " [s_remark] char(200))"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
        
ErrCreate:
'    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpBlockprev", oTempADO, True
End Sub


Private Sub Command13_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer
    
    createtmpBlockprev

    cSqlStmt = "insert into tmpBlockprev (blkid,empid, ssnum,date_hire,fullname,birthday,address,tel_num,linename,posname,active, emp_stat,s_remark)values( " & _
               cQuote & Text1.Text & cQuote & "," & _
               cQuote & Text2.Text & cQuote & "," & _
               cQuote & Text3.Text & cQuote & "," & _
               cQuote & Label14.Caption & cQuote & "," & _
               cQuote & Label18.Caption & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label21.Caption)) & cQuote & "," & _
               cQuote & Label22.Caption & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label23.Caption)) & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label10.Caption)) & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label24.Caption)) & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label25.Caption)) & cQuote & "," & _
               cQuote & EncodeStr(DecodeStr(Label26.Caption)) & cQuote & "," & _
               cQuote & Text33.Text & cQuote & ")"
    
'            MsgBox cSqlStmt
'    Script2File cSqlStmt
    QueryTemp cSqlStmt, objdbRs, True
        
    GenerateReport "Block List Preview", "prv859938.RPT", , True

End Sub
Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
    Dim cSqlStmt As String, _
        sLinename As String, _
        sPosname As String, _
        sActive As String, _
        sAddress As String
        
    Dim oRSet As New ADODB.Recordset
    
    Frame2.Enabled = False
    
     sActive = IIf(Combo1.ListIndex = 0, " where a.ssnum <> '' and a.active = 0 ", IIf(Combo1.ListIndex = 1, " where a.ssnum <> '' and a.active <> 0 ", " where a.ssnum <> '' "))
    
    frmLookup.showPopup 25, sActive

    frmLookup.Show 1
    
    If Trim(cResult) <> "" Then
        
        cSqlStmt = " select EMPID,SSNUM,DATE_HIRE,concat(LASTNAME,', ',FIRSTNAME,' ', left(MNAME,1),'. ') as fullname, " & _
                   " BIRTHDAY,ADD_NO,ADD_BRGY,ADD_CITY, " & _
                   " TEL_NUM,DEPID,POSID, if(EMP_STAT=0,'Wap',if(EMP_STAT=1,'Conmtractual','Regular')) as emp_stat, " & _
                   " concat(if(ACTIVE=0,'Active',if(ACTIVE=1,'Resigned',if(ACTIVE=2,'Finished','Terminated'))), ' - ' , DATE_RES ) as date_Term,S_REMARK " & _
                   " from di3670 " & _
                   " Where SSNUM = " & cQuote & cResult & cQuote
               
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
        
            OpenQueryDNS "select * from di5463 where lineid = " & cQuote & objdbRs("depid") & cQuote, oRSet, False
            sLinename = IIf(oRSet.RecordCount > 0, oRSet("linename"), "")
            
            OpenQueryDNS "select * from di7670 where posid = " & cQuote & objdbRs("posid") & cQuote, oRSet, False
            sPosname = IIf(oRSet.RecordCount > 0, oRSet("posname"), "")
            
            Text2.Text = objdbRs("empid")
            Text3.Text = cResult
            Label14.Caption = Format(objdbRs("DATE_HIRE"), "yyyy-mm-dd")
            Label18.Caption = objdbRs("fullname")
            Label21.Caption = Format(objdbRs("BIRTHDAY"), "yyyy-mm-dd")
            
            sAddress = IIf(objdbRs("ADD_NO") = "", "", objdbRs("ADD_NO") & " ")
            sAddress = sAddress & IIf(objdbRs("ADD_BRGY") = "", "", objdbRs("ADD_BRGY") & " ")
            sAddress = sAddress & IIf(objdbRs("ADD_CITY") = "", "", objdbRs("ADD_CITY") & " ")
            Label22.Caption = IIf(sAddress = "", "Bataan", sAddress & "Bataan")
    
            
            Label23.Caption = objdbRs("TEL_NUM")
            Label10.Caption = sLinename
            Label24.Caption = sPosname
            Label25.Caption = objdbRs("emp_stat")
            Label26.Caption = objdbRs("date_Term")
            Text33.Text = objdbRs("s_remark")
            
        Else
            Text2.Text = ""
            Text3.Text = ""
            Label14.Caption = ""
            Label18.Caption = ""
            Label21.Caption = ""
            Label22.Caption = ""
            Label23.Caption = ""
            Label10.Caption = ""
            Label24.Caption = ""
            Label25.Caption = ""
            Label26.Caption = ""
            Text33.Text = ""
        End If
                
    End If
    Frame2.Enabled = True
End Sub

Private Sub Command5_Click()

    Log2Audit Name, "SEARCH"
    Frame2.Enabled = False
        frmLookup.showPopup 26
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            oTempADO.Requery adAsyncFetch
            oTempADO.Find "BLKID='" & PadStr(Trim(cResult), " ", Text1.MaxLength, PadRight) & "'"
            If Not oTempADO.EOF Then GetFields Me, oTempADO
            ShowRecords
        End If
    Frame2.Enabled = True

End Sub

Private Sub Command7_Click()
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Command12.Enabled = nAdd <> 0
    Command4.Enabled = nAdd <> 0
    Combo1.Enabled = nAdd <> 0
    
    Text33.Text = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label21.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
    Label10.Caption = ""
    Label24.Caption = ""
    Label25.Caption = ""
    Label26.Caption = ""
    Text33.Text = ""
    
    cSeries = GenerateSeries("BLKID")
    Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    While IfExists("PA255578", "PA255578.BLKID=" & cQuote & PadStr(cSeries, "0", Text1.MaxLength) & cQuote)
        cSeries = GenerateSeries("BLKID")
        Text1.Text = PadStr(cSeries, "0", Text1.MaxLength)
    Wend
    Text2.SetFocus
End Sub

Private Sub Command8_Click()
    ' --> 20050329
    If Not isDataLock(Me.Name, Text1.ToolTipText, Text1.Text) Then
        Lock2User Name, Text1.ToolTipText, Text1.Text, True
        nAdd = 2
        
        Command12.Enabled = nAdd <> 0
        Command4.Enabled = nAdd <> 0
        Combo1.Enabled = nAdd <> 0

        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        Text1.Enabled = False
        Text2.SetFocus
    End If
End Sub

Private Sub Command9_Click()
    On Error GoTo ErrBLKLSTDelete
    If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Block List Entry...") = vbYes Then
        OpenQueryDNS "DELETE FROM PA255578 WHERE BLKID=" & cQuote & Text1.Text & cQuote, oTempADO, True
        Script2File "DELETE FROM PA255578 WHERE BLKID=" & cQuote & Text1.Text & cQuote
        
        Log2Audit Name, "DELETE " & Text1.Text & " - " & Trim(EncodeStr2(DecodeStr(Text2.Text))) & " " & Trim(EncodeStr2(DecodeStr(Text3.Text)))
        
        nAdd = 0
        ClearAll Me, False, True

        oTempADO.Requery adAsyncFetch
        GetFields Me, oTempADO
        ShowRecords
    End If
    
    Exit Sub
    
ErrBLKLSTDelete:
    ErrorMsg Err.Number, Err.Description, "Delete Button", Name
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    Tag = nAccess_Tag
    nAdd = 0
    ClearAll Me, False, True
    CtrlPanel Me, nAdd

    OpenQueryDNS "SELECT * FROM PA255578", oTempADO, False

    GetFields Me, oTempADO
    ShowRecords
    
End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label46_Click()
End Sub

