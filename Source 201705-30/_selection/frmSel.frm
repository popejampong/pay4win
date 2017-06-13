VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{083C8784-F106-4CC2-9930-876218A6B74C}#1.1#0"; "ciaXPButton.ocx"
Begin VB.Form frmSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select from List..."
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check7 
      BackColor       =   &H00800000&
      Caption         =   "&Audit Report"
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
      Left            =   4920
      TabIndex        =   49
      Top             =   3120
      Visible         =   0   'False
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
      Left            =   4950
      TabIndex        =   48
      Top             =   2370
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00800000&
      Caption         =   "&Tagalog"
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
      Left            =   4950
      TabIndex        =   47
      Top             =   2745
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4770
      Top             =   -150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ciaXPButton.XPButton XPButton3 
      Height          =   600
      Left            =   4965
      TabIndex        =   11
      Top             =   975
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1058
      Caption         =   "&Ok"
      Picture         =   "frmSel.frx":0000
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
   Begin ciaXPButton.XPButton XPButton1 
      Cancel          =   -1  'True
      Height          =   600
      Left            =   4965
      TabIndex        =   9
      Top             =   1590
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1058
      Caption         =   "&Close"
      Picture         =   "frmSel.frx":1992
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
      Left            =   4935
      TabIndex        =   8
      Top             =   3945
      Width           =   1440
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00800000&
      Caption         =   "&Deduction Only"
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
      Left            =   4935
      TabIndex        =   7
      Top             =   3510
      Width           =   1440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3930
      Left            =   90
      TabIndex        =   2
      Top             =   105
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   6932
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Select Period"
      TabPicture(0)   =   "frmSel.frx":3324
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(1)=   "ListView1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Select Department"
      TabPicture(1)   =   "frmSel.frx":3340
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check2"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmSel.frx":335C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text2"
      Tab(2).Control(1)=   "Command1"
      Tab(2).Control(2)=   "Combo3"
      Tab(2).Control(3)=   "Combo2"
      Tab(2).Control(4)=   "Label13"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "Label3"
      Tab(2).Control(7)=   "Label1"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Signatory"
      TabPicture(3)   =   "frmSel.frx":3378
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "Label5"
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(4)=   "Label6"
      Tab(3).Control(5)=   "Label15"
      Tab(3).Control(6)=   "Label10"
      Tab(3).Control(7)=   "Label8"
      Tab(3).Control(8)=   "Label16"
      Tab(3).Control(9)=   "Label12"
      Tab(3).Control(10)=   "Label14"
      Tab(3).Control(11)=   "Text1"
      Tab(3).Control(12)=   "Command5"
      Tab(3).Control(13)=   "Command14"
      Tab(3).Control(14)=   "Text5"
      Tab(3).Control(15)=   "Text7"
      Tab(3).Control(16)=   "Command15"
      Tab(3).Control(17)=   "Text6"
      Tab(3).Control(18)=   "Command13"
      Tab(3).Control(19)=   "Text8"
      Tab(3).Control(20)=   "Command16"
      Tab(3).Control(21)=   "Command2"
      Tab(3).Control(22)=   "Text3"
      Tab(3).ControlCount=   23
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmSel.frx":3394
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label17"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Dir1"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Drive1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Text4"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.TextBox Text4 
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3480
         Width           =   4290
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   135
         TabIndex        =   44
         Top             =   2805
         Width           =   4305
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   135
         TabIndex        =   43
         Top             =   420
         Width           =   4290
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
         Left            =   -74895
         TabIndex        =   20
         Tag             =   "1"
         ToolTipText     =   "TXT:INSP_BY"
         Top             =   3000
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   41
         Top             =   2985
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
         Left            =   -74880
         TabIndex        =   38
         Tag             =   "1"
         ToolTipText     =   "TXT:PREP_BY"
         Top             =   1515
         Width           =   660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   -74190
         TabIndex        =   37
         Top             =   1485
         Width           =   375
      End
      Begin VB.CommandButton Command16 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   26
         Top             =   3285
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
         Left            =   -74895
         TabIndex        =   21
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   3300
         Width           =   660
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   25
         Top             =   750
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
         Left            =   -74895
         TabIndex        =   16
         Tag             =   "1"
         ToolTipText     =   "TXT:PREP_BY"
         Top             =   765
         Width           =   660
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   24
         Top             =   2400
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
         Left            =   -74895
         TabIndex        =   19
         Tag             =   "1"
         ToolTipText     =   "TXT:INSP_BY"
         Top             =   2415
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
         Left            =   -74895
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "TXT:CHK_BY"
         Top             =   1305
         Width           =   660
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   23
         Top             =   1290
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   300
         Left            =   -74205
         TabIndex        =   22
         Top             =   1845
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
         Left            =   -74895
         TabIndex        =   18
         Tag             =   "1"
         ToolTipText     =   "TXT:REC_BY"
         Top             =   1860
         Width           =   660
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmSel.frx":33B0
         Left            =   -74070
         List            =   "frmSel.frx":33BD
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   780
         Width           =   2190
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmSel.frx":33D3
         Left            =   -74070
         List            =   "frmSel.frx":33E3
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   435
         Width           =   2190
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Select &All"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -74880
         TabIndex        =   5
         Top             =   3630
         Width           =   1755
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Select &All"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -74880
         TabIndex        =   3
         Top             =   3630
         Width           =   4275
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3210
         Left            =   -74880
         TabIndex        =   4
         Top             =   405
         Width           =   4320
         _ExtentX        =   7620
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3210
         Left            =   -74880
         TabIndex        =   6
         Top             =   405
         Width           =   4320
         _ExtentX        =   7620
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Path / Location"
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
         Left            =   135
         TabIndex        =   46
         Top             =   3255
         Width           =   4035
      End
      Begin VB.Label Label14 
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
         Left            =   -73785
         TabIndex        =   42
         Top             =   3045
         Width           =   3270
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Signatory"
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
         Left            =   -74865
         TabIndex        =   40
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         Left            =   -73770
         TabIndex        =   39
         Top             =   1530
         Width           =   3270
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
         TabIndex        =   36
         Top             =   2760
         Width           =   4290
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
         Left            =   -73785
         TabIndex        =   35
         Top             =   3360
         Width           =   3270
      End
      Begin VB.Label Label8 
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
         Left            =   -73785
         TabIndex        =   34
         Top             =   810
         Width           =   3270
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
         TabIndex        =   33
         Top             =   555
         Width           =   4290
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
         Left            =   -73785
         TabIndex        =   32
         Top             =   2460
         Width           =   3270
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
         Left            =   -73785
         TabIndex        =   31
         Top             =   1350
         Width           =   3270
      End
      Begin VB.Label Label11 
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
         Left            =   -74880
         TabIndex        =   30
         Top             =   2205
         Width           =   4290
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
         TabIndex        =   29
         Top             =   1095
         Width           =   4290
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
         TabIndex        =   28
         Top             =   1650
         Width           =   4290
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
         Left            =   -73785
         TabIndex        =   27
         Top             =   1905
         Width           =   3270
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   -74895
         TabIndex        =   15
         Top             =   825
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Quarter"
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   -74895
         TabIndex        =   13
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSel.frx":3425
      Left            =   90
      List            =   "frmSel.frx":342F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4290
      Width           =   4215
   End
   Begin ciaXPButton.XPButton XPButton2 
      Height          =   600
      Left            =   4965
      TabIndex        =   10
      Top             =   360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1058
      Caption         =   "&Process"
      Picture         =   "frmSel.frx":3467
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
      Left            =   90
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   4710
      Left            =   4755
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmSel
' description   :   Module for any kind of selection
' programmer    :   _-=[ srm ]=-_
' date          :   23 Oct 2005

Option Explicit
    Dim oTempADO As New ADODB.Recordset, _
        oDBFConn As New ADODB.Connection, _
        nHolWithPay As Integer, _
        nHolWPND As Integer
    Dim nTagSelect As Integer

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
            Case 6
                Command1_Click
            Case 7
                Command2_Click
        End Select
    Else
        ShowData cString, oLabel
    End If
End Sub
Sub SetSelection(ByVal nMode As Integer)
    Dim oLstItem As ListItem, nCtr As Integer
    
    Tag = nMode
    
    Check3.Caption = "&Deduction Only"
    Label10.Caption = "Prepared By"
    
    Check1.Visible = True
    Check3.Visible = True
    Check4.Visible = True
    Check5.Visible = False
    
    Combo1.Visible = True
    Combo2.Visible = True
    Combo3.Visible = True
    
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    
    XPButton2.Visible = True
    
    Command2.Visible = True
    Command5.Visible = True
    Command14.Visible = True
    Command15.Visible = True
    Command16.Visible = True
    
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    
    Text1.Visible = True
    Text3.Visible = True
    Text5.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    
    With SSTab1
        .TabVisible(0) = True
        .TabVisible(1) = True
        .TabVisible(2) = True
        .TabVisible(3) = True
        .TabVisible(4) = True
        .Tab = 0
    End With
    
    ShowData Text1.Text, Label4
    ShowData Text2.Text, Label9
    ShowData Text3.Text, Label14
    ShowData Text5.Text, Label6
    ShowData Text6.Text, Label8
    ShowData Text7.Text, Label15
    ShowData Text8.Text, Label16
    
    Select Case nMode
        Case 1, 2, 7, 8, 23, 35, 41      ' --> Transaction Selection
            Check1.Visible = False
            
            Check3.Visible = False
            Check4.Visible = (nMode = 7) Or (nMode = 23) Or (nMode = 35) Or (nMode = 41)
            Check6.Visible = IIf((nMode = 7) Or (nMode = 23) Or (nMode = 41), True, False)
            
            If nMode = 41 Then
                With Combo1
                    .Clear
                    .AddItem "Emergency"
                    .AddItem "Finish/Resign"
                    .ListIndex = 0
                    .Visible = True
                End With
            Else
                Combo1.Visible = False
            End If
            Label2.Visible = False
            
            XPButton2.Visible = nMode = 1
            
            With SSTab1
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(3) = (nMode = 7) Or (nMode = 23) Or (nMode = 35) Or (nMode = 41)
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
            End With
            
            If (nMode = 7) Or (nMode = 23) Or (nMode = 35) Or (nMode = 41) Then
                Label10.Caption = "From"
                Label5.Caption = "Attention to:"
                Label4.Visible = False
                Label7.Visible = False
                Label11.Visible = False
                Label12.Visible = False
                Label14.Visible = False
                Label15.Visible = False
                Label16.Visible = False
                
                Text1.Visible = False
                Text3.Visible = False
                Text7.Visible = False
                Text8.Visible = False
                
                Command2.Visible = False
                Command5.Visible = False
                Command15.Visible = False
                Command16.Visible = False
            End If
            
            ListView1.CheckBoxes = False
            If (nMode = 2) Then
                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE 13month=0 ", objdbRs, False
            Else
                If nMode = 23 Or nMode = 35 Then
                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=1) and (PCLOSE=0) ORDER BY PERIODID", objdbRs, False
                Else
                    If (nMode = 7) Or (nMode = 41) Then
                        OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
                    Else
                        OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
                        
                    End If
                End If
            End If
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
    
        Case 3, 4, 5, 6, 10, 15, 22, 34, 36, 44, 45
'            Check1.Visible = False

            ' --> addnew 2015-05-07
            If (nMode = 6) Or (nMode = 15) Then
                Check1.Caption = "Exclude Close Period"
                Check1.Visible = True
                Check1.Value = vbUnchecked
            Else
                Check1.Visible = False
            End If
            
            Check3.Visible = IIf((nMode = 6) Or (nMode = 15) Or (nMode = 10), True, False)
            Check4.Visible = IIf((nMode = 6) Or (nMode = 15) Or (nMode = 3), True, False)
            Check4.Value = vbChecked
            
            Check6.Visible = IIf((nMode = 4) Or (nMode = 5) Or (nMode = 6) Or (nMode = 22) Or (nMode = 15) Or (nMode = 44) Or (nMode = 45), True, False)
            
            ' --> 20071207
            If nMode = 10 Then
                Check3.Caption = "Emergency Manpower"
                Check5.Caption = "&Order by Position"
                Check5.Visible = True
            End If
            If nMode = 4 Then
                Check5.Visible = True
            End If
            
            Check4.Caption = "Detailed Report"
            
            
            XPButton2.Visible = False
            
            ListView1.CheckBoxes = False
            If (nMode = 10) Or (nMode = 3) Then
                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE 13month=0 ", objdbRs, False
'                add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            Else
                If nMode = 22 Or nMode = 34 Then
                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=1) and (PCLOSE=0) ORDER BY PERIODID", objdbRs, False
                Else
                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
                End If
            End If
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
            ListView2.CheckBoxes = True
            
            With SSTab1
                .TabVisible(2) = False
                .TabVisible(3) = IIf(nMode = 6, Check4.Value <> vbChecked, False)
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Select Department"
            End With
            
            Check2.Value = vbChecked
            Check2_Click
        
            OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINEID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
                         
            If nMode = 10 Then
                Check4.Caption = "Group Report"
                Combo1.Visible = False
                Label2.Visible = False
            Else
                If (nMode <> 22) And (nMode <> 34) And (nMode <> 36) Then
                    Label2.Caption = "Select filter:"
                    
                    With Combo1
                        .Clear
                        .AddItem "Regular"
                        .AddItem "SA"
                        .AddItem "WAP"
                        .AddItem "WAP SA"
                        .AddItem "Emergency"
                        .ListIndex = 0
                    End With
                Else
                    Label2.Visible = False
                    Combo1.Visible = False
                End If
            End If
            
        Case 9
            
            XPButton2.Visible = False
            
            Check3.Caption = "Emergency Manpower"
            Check4.Caption = "&Order by Sex"
            Check5.Caption = "&Order by Position"
            
            Check5.Visible = True
            
            With SSTab1
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = False
                
                .TabCaption(0) = "Month"
                .TabCaption(1) = "Department"
            End With
            
            With ListView1
                .ListItems.Clear
                For nCtr = 1 To 12
                    Set oLstItem = .ListItems.Add()
                    oLstItem.Text = Trim(Str(nCtr))
                    oLstItem.SubItems(1) = MonthName(nCtr)
                    If nCtr = Month(Now) Then oLstItem.Checked = True
                Next
            End With
            
            OpenQueryDNS "SELECT LINEID, LINENAME FROM di5463 ORDER BY LINEID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
            
            Label2.Caption = "Select Year"
           
'           UPDATE **201410-09**
            With Combo1
               .Clear
               .AddItem Year(Now) - 2
               .AddItem Year(Now) - 1
               .AddItem Year(Now)
               .AddItem Year(Now) + 1
               .AddItem Year(Now) + 2
               .AddItem Year(Now) + 3
            End With
            
            
            MatchCombo Year(Now), Combo1
            
        Case 11 ' --> Leave Report
            Check1.Visible = False
            Check3.Visible = False
            Check4.Visible = False
            
            Command2.Visible = False
            
            Label2.Visible = False
            Label14.Visible = False
            
            Text3.Visible = False
            
            XPButton2.Visible = False
            
            With SSTab1
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(4) = False
                
                .TabCaption(0) = "Select Period"
                .TabCaption(3) = "Signatory"
            End With
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
            With Combo1
                .Clear
                .AddItem "Sick / Vacation Leave"
                .AddItem "Emergency Leave"
                .AddItem "Maternity Leave"
                .AddItem "Paternity Leave"
                .AddItem "Force Leave"
                .AddItem "Union Leave"
                .AddItem "All Leave"
                .ListIndex = 0
            End With
            
        Case 12         ' new medicare
            XPButton2.Visible = False
            
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = False
            
            With SSTab1
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(4) = True
                
                .TabCaption(4) = "Set Path"
                .TabCaption(0) = "Month"
            End With
            
            Label10.Caption = "Certified Correct By:"
            
            Label5.Visible = False
            Text5.Visible = False
            Command14.Visible = False
            Label6.Visible = False
            
            Label7.Visible = False
            Text1.Visible = False
            Command5.Visible = False
            Label4.Visible = False
            
            Label11.Visible = False
            Text7.Visible = False
            Command15.Visible = False
            Label15.Visible = False
            Text3.Visible = False
            Command2.Visible = False
            Label14.Visible = False
            
            Label12.Visible = False
            Text8.Visible = False
            Command16.Visible = False
            Label16.Visible = False
            
            '---> 201209-11 Costcenter
            Check3.Visible = True
            Check3.Caption = "Cost Center"
            Check3.Value = vbChecked

            
            With ListView1
                .ListItems.Clear
                For nCtr = 1 To 12
                    Set oLstItem = .ListItems.Add()
                    oLstItem.Text = Trim(Str(nCtr))
                    oLstItem.SubItems(1) = MonthName(nCtr)
                    If nCtr = Month(Now) Then oLstItem.Checked = True
                Next
            End With
            
            Label2.Caption = "Select Year"
            With Combo1
               .Clear
               .AddItem Year(Now) - 2
               .AddItem Year(Now) - 1
               .AddItem Year(Now)
               .AddItem Year(Now) + 1
               .AddItem Year(Now) + 2
               .AddItem Year(Now) + 3
            End With
            
            MatchCombo Year(Now), Combo1
    
        Case 13, 14, 16, 17, 21, 26, 42
            Command2.Visible = False
            Command5.Visible = False
            Command14.Visible = False
            Command15.Visible = False
            Command16.Visible = False
            
            Label4.Visible = False
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label11.Visible = False
            Label12.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            
            Text1.Visible = False
            Text3.Visible = False
            Text5.Visible = False
            Text7.Visible = False
            Text8.Visible = False
            
            XPButton2.Visible = False
            If nMode = 17 Then
                Check1.Visible = True
                Check1.Caption = "Close Period"
            Else
                Check1.Visible = False
            End If
            Check3.Visible = False
            
            If nMode = 13 Then
                Check4.Visible = False
            Else

         
                Check4.Caption = "Premium only"
                Check4.Value = vbChecked
                Check3.Visible = True
                Check3.Caption = "Cost Center"
'                Check3.Value = vbChecked
                Check7.Visible = False
                Check7.Caption = "Calamity Loan"
                
                
            End If
            
            With SSTab1
                .TabVisible(0) = IIf(Tag = 21, False, True)
                .TabVisible(1) = False
                .TabVisible(2) = False
                
                If (Tag = 17) Or (Tag = 21) Then
                    Check6.Visible = False
                    ListView1.CheckBoxes = False
                    '20081229 for period early backup of payroll
                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
                    'OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0)  ORDER BY PERIODID", objdbRs, False
                    add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
                    
                    .TabVisible(3) = False
                    .TabCaption(4) = "Select a location"
                    
                    Check4.Visible = False
                    Label2.Visible = False
                    Combo1.Visible = False
                    Drive1.Drive = left(cDownloadPath, 2)
                    Dir1.Path = cDownloadPath
                    Text4.Text = cDownloadPath
'                    Text9.Text = "K1PAY"
                Else
                    .TabVisible(4) = False
                    
                    If Tag = 16 Then .TabVisible(3) = False
                    .TabCaption(0) = "Month"
                    .TabCaption(3) = "Signatory"
                    
                End If
            End With
            
            If Tag = 26 Then
                Check4.Visible = False
                Label10.Caption = "Certified Correct:"
            End If
            
            If (Tag <> 17) And (Tag <> 21) Then
                With ListView1
                    .ListItems.Clear
                    For nCtr = 1 To 12
                        Set oLstItem = .ListItems.Add()
                        oLstItem.Text = Trim(Str(nCtr))
                        oLstItem.SubItems(1) = MonthName(nCtr)
                    Next
                End With
                ListView1.CheckBoxes = False
                
                Label2.Caption = "Select Year"
                With Combo1
                   .Clear
                   .AddItem Year(Now) - 2
                   .AddItem Year(Now) - 1
                   .AddItem Year(Now)
                   .AddItem Year(Now) + 1
                   .AddItem Year(Now) + 2
                   .AddItem Year(Now) + 3
                End With
                
                MatchCombo Year(Now), Combo1
            End If
            
        Case 18, 46    ' --> Salary Division (revised 20071020)
            XPButton2.Visible = False
            
            Check1.Visible = False
            Check3.Visible = False
            Check4.Caption = "Detailed Report"
            Check5.Visible = False
            
'            Check6.Visible = IIf((nMode = 18), True, False)
            If gCompanyID <> "0003" Then
                Check6.Visible = IIf((nMode = 18), True, False)
            Else
                Check6.Visible = IIf((nMode = 18) Or (nMode = 46), True, False)
            End If
           
            With SSTab1
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = False
            End With
            
'            Combo1.Visible = True
'            Label2.Visible = True
            Label2.Caption = "Select Filter:"
            With Combo1
                .Clear
                .AddItem "Regular Payroll"
                .AddItem "Emergency Payroll"
                .ListIndex = 0
            End With
    
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
        Case 19, 20, 32, 33
            XPButton2.Visible = False
            
            If gCompanyID = "0002" Then
                Check3.Visible = True
            Else
                Check3.Visible = False
            End If
            
            Check5.Visible = False
            Check4.Visible = nMode = 20 Or nMode = 33
            Check6.Visible = IIf((nMode = 19) Or (nMode = 20) Or (nMode = 33), True, False)
            Check4.Value = vbChecked
            Check4.Caption = "Detailed Report"
    
            With SSTab1
                .TabVisible(0) = False
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = False
            End With
            
            Label2.Visible = False
            Combo1.Visible = False
            
            OpenQueryDNS "SELECT LINEID, LINENAME FROM di5463 ORDER BY LINEID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
    
        Case 24, 25
            With SSTab1
                .TabVisible(0) = True
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabCaption(4) = "Select a location"
            End With
            
            Check1.Visible = False
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = False
            XPButton2.Visible = False
            
            Drive1.Drive = left(cDownloadPath, 2)
            Dir1.Path = cDownloadPath
            Text4.Text = cDownloadPath
            
            Label2.Caption = "Select Year"
            With Combo1
               .Clear
               .AddItem Year(Now) - 2
               .AddItem Year(Now) - 1
               .AddItem Year(Now)
               .AddItem Year(Now) + 1
               .AddItem Year(Now) + 2
               .AddItem Year(Now) + 3
            End With
            
            MatchCombo Year(Now) - 1, Combo1
                
            ListView1.CheckBoxes = False
            'OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (wtax=1) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE year(date_start) = year(CURDATE()) -1  and month(date_start) = 12 and day((date_start)) > 15 ", objdbRs, False
            
            'OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
'            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (PCLOSE=0) and 13month =0", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
        ' --> Incentive Payroll
        Case 27, 28, 29, 30
            Check1.Visible = False
            
            Check3.Visible = False
           
            Check4.Visible = IIf((nMode = 28) Or (nMode = 29), True, False)
            Check4.Value = vbChecked
            
'            Check3.Visible = IIf((nMode = 6) Or (nMode = 15) Or (nMode = 10), True, False)
'            Check4.Visible = IIf((nMode = 6) Or (nMode = 15) Or (nMode = 3), True, False)
            

'            ' --> 20071207
'            If nMode = 10 Then Check3.Caption = "Emergency Manpower"
            
            Check4.Caption = "Detailed Report"
'
'            Check5.Visible = nMode = 4
'            Check5.Value = IIf(nMode = 4, vbChecked, vbUnchecked)
            
            XPButton2.Visible = False
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            
'            If (nMode = 10) Or (nMode = 3) Then
'                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE 13month=0 ", objdbRs, False
''                add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
'            Else
'                If nMode = 22 Then
'                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=1) and (PCLOSE=0) ORDER BY PERIODID", objdbRs, False
'                Else
'                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
'                End If
'            End If
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
            ListView2.CheckBoxes = True
            
            With SSTab1
                .TabVisible(2) = False
                .TabVisible(3) = IIf(nMode = 6, Check4.Value <> vbChecked, False)
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Select Department"
            End With
            
            Check2.Value = vbChecked
            Check2_Click
        
            OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINEID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
             
            Label2.Caption = "Select filter:"

            With Combo1
                .Clear
                .AddItem "Regular"
                .AddItem "SA"
                .AddItem "WAP"
                .AddItem "WAP SA"
                .AddItem "Emergency"
                .ListIndex = 0
            End With
            
        Case 31 ' ---> Load Report 20080805
            
            Check1.Visible = False
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = False
            
            XPButton2.Visible = False
            
            Combo1.Visible = False
            
            Label2.Visible = False
            
            With SSTab1
                .TabVisible(0) = False
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = False
                .TabCaption(1) = "Select Loan"
            End With
            
            OpenQueryDNS "SELECT DEDID, DEDNAME FROM pa3330 ORDER BY DEDID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("DEDNAME", "DEDID")
            
            Check2.Value = vbChecked
            Check2_Click
        
        Case 39
            Command2.Visible = False
            Command5.Visible = False
            Command14.Visible = False
            Command15.Visible = False
            Command16.Visible = False
            
            Label4.Visible = False
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label10.Visible = False
            Label11.Visible = False
            Label12.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            
            Text1.Visible = False
            Text3.Visible = False
            Text5.Visible = False
            Text7.Visible = False
            Text8.Visible = False
            
            XPButton2.Visible = False
            Check1.Visible = False
            Check3.Visible = False
            
            Check4.Visible = True
            Check4.Caption = "NET PAY"
            Check4.Value = vbChecked
            Check4.Enabled = False
           Check6.Visible = IIf((nMode = 39), True, False)
           
            With SSTab1
                
                .TabVisible(0) = True
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabCaption(1) = "Select Department"
                
                ListView1.CheckBoxes = False
                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
                add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
                .TabVisible(3) = False
                .TabVisible(4) = False
                .TabCaption(4) = "Select a location"
                
                Label2.Visible = False
                
                Combo1.Visible = True
                
                OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINEID", objdbRs, False
                add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
                
                
                With Combo1
                    .Clear
                    .AddItem "WAP"
                    .AddItem "Regular"
                    .ListIndex = 0
                End With
                
                Drive1.Drive = left(cDownloadPath, 2)
                Dir1.Path = cDownloadPath
                Text4.Text = cDownloadPath
            End With
            
        Case 40 ' --> GrandOT Report 2009-08-17
            Check1.Visible = False
            Check3.Visible = False
            Check4.Visible = False
            Combo1.Visible = False
            
            
            Command2.Visible = False
            
            Label2.Visible = False
            Label14.Visible = False
            
            Text3.Visible = False
            
            XPButton2.Visible = False
            
            With SSTab1
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(4) = False
                
                .TabCaption(0) = "Select Period"
                .TabVisible(3) = False
                
            End With
            
            ListView1.CheckBoxes = False
            'OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            'OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) ORDER BY PERIODID", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
                    
        Case 43 ' --> GrandOT Report 2009-08-17
            Check1.Visible = False
            Check3.Visible = False
            Check4.Visible = False
            
            
            Command2.Visible = False
            
            Label2.Visible = False
            Label14.Visible = False
            
            Text3.Visible = False
            
            XPButton2.Visible = False
            
            With SSTab1
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Department"
                .TabVisible(3) = False
            End With
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=1) and year(date_end)< year(current_date)-1 ORDER BY PERIODID DESC", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
            ListView2.CheckBoxes = True
            OpenQueryDNS "SELECT LINEID, LINENAME FROM di5463 ORDER BY LINEID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
                    
            Label2.Caption = "Type of Report"
            With Combo1
                .Visible = True
                .Clear
                .AddItem "Regular Report"
                .AddItem "Audit Report"
                .ListIndex = 0
            End With
            
        Case 47, 48, 49 ' --> TMS Report 2011-03-21
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = False
            Check6.Visible = False
            Check7.Visible = True
            Check1.Caption = "Exclude Close Period"
            Check1.Value = vbChecked
            
            Label12.Visible = False
            Command2.Visible = False
            Label14.Visible = False
            Text3.Visible = False
            
            Command16.Visible = False
            Label16.Visible = False
            Text8.Visible = False
            
            Label11.Visible = False
            Label15.Visible = False
            Text7.Visible = False
            Command15.Visible = False
            
            Label5.Caption = "Noted By"
            Label7.Caption = "Approved By"
            
            XPButton2.Visible = False
            
            With SSTab1
                .TabVisible(0) = True
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabVisible(3) = True
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Department"
                
            End With
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 " & IIf(Check1.Value = vbChecked, " WHERE PCLOSE=0 ", "") & " ORDER BY PERIODID DESC", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")

            
            ListView2.CheckBoxes = True
            OpenQueryDNS "SELECT LINEID, LINENAME FROM di5463 ORDER BY LINENAME", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")
                    
            Label2.Caption = "Type of Report"
            With Combo1
                .Visible = True
                .Clear
                .AddItem "Complete Report"
                .AddItem "Regular Report"
                .AddItem "Extension Report"
                .ListIndex = 0
            End With
            
            If nMode = 48 Then
                nTagSelect = 1
            ElseIf nMode = 49 Then
                nTagSelect = 2
            End If
            
        Case 50
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = False
            Check6.Visible = False
            Check7.Visible = False
            Check1.Caption = "Exclude Close Period"
            Check1.Value = vbChecked
            
            Label12.Visible = False
            Command2.Visible = False
            Label14.Visible = False
            Text3.Visible = False
            
            Command16.Visible = False
            Label16.Visible = False
            Text8.Visible = False
            
            Label11.Visible = False
            Label15.Visible = False
            Text7.Visible = False
            Command15.Visible = False
            
            Label5.Caption = "Noted By"
            Label7.Caption = "Approved By"
            
            XPButton2.Visible = False
            
            With SSTab1
                .TabVisible(0) = True
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabVisible(3) = True
                .TabVisible(4) = False
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Cost Center"
                
            End With
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 " & IIf(Check1.Value = vbChecked, " WHERE PCLOSE=0 ", "") & " ORDER BY PERIODID DESC", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")

            ListView2.CheckBoxes = True
            OpenQueryDNS "select COSTCENTERID, DESCRIPTION from pa37722 where compcode = " & cQuote & nCompCode & cQuote & _
                         " ORDER BY COSTCENTERID", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("DESCRIPTION", "COSTCENTERID")
                    
            Label2.Caption = "Type of Report"
            With Combo1
                .Visible = True
                .Clear
                .AddItem "Summary Report"
                .AddItem "Detailed Report"
                .ListIndex = 0
            End With
            
            Combo1.Visible = False
            Label2.Visible = False
            
'            If nMode = 48 Then
'                nTagSelect = 1
'            ElseIf nMode = 49 Then
'                nTagSelect = 2
'            End If

        Case 51 ' --> Employee Master Data Generation
        
            Dir1.Path = cDownloadPath
            Text4.Text = cDownloadPath
        
        
            XPButton2.Visible = False
        
            Check3.Visible = False
            Check4.Visible = False
            Check5.Visible = True
            Check5.Caption = "All employee"
            Check6.Visible = False
            Check7.Visible = False
            Check1.Caption = "Exclude Close Period"
            Check1.Value = vbChecked
            
            Label12.Visible = False
            Command2.Visible = False
            Label14.Visible = False
            Text3.Visible = False
            
            Command16.Visible = False
            Label16.Visible = False
            Text8.Visible = False
            
            Label11.Visible = False
            Label15.Visible = False
            Text7.Visible = False
            Command15.Visible = False
            
            With SSTab1
                .TabVisible(0) = True
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = True
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Department"
                .TabCaption(4) = "Generate File"
                
            End With
            
            ListView1.CheckBoxes = False
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 " & IIf(Check1.Value = vbChecked, " WHERE 13month <> 1 and PCLOSE=0 ", "WHERE 13month <> 1 ") & " ORDER BY PERIODID DESC", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")

            ListView2.CheckBoxes = True
            OpenQueryDNS "SELECT LINEID, LINENAME FROM di5463 ORDER BY LINENAME", objdbRs, False
            add2LstBox objdbRs, ListView2, Array("LINENAME", "LINEID")

            Label2.Caption = "Status Type"
            With Combo1
                .Visible = True
                .Clear
                .AddItem "Active"
                .AddItem "Resigned"
                .AddItem "Finished"
                .AddItem "Terminated"
                .AddItem ""
                .ListIndex = 0
            End With

    End Select
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

' + -->
' |     Procedure Name  :   ProcessTransaction
' |     Description     :   Process Payroll Transaction of a particular period
' |     Date Created    :   xx mar 2006
' + -->

' --> 20070814 - compute adjustment here (08-01 to 08-02) due to salary adjustment effective aug 3, 2007
Function ChkAdjustment(ByVal cEmpID As String, _
                       aRateInfo As Variant, _
                       aPeriodInfo As Variant, _
                       ByVal aEmpStat As Variant) As Variant
    Dim cSqlStmt As String, _
        aTimeInfo As Variant, _
        aPayInfo As Variant

    aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#)
    aPayInfo = Array(0#, 0#, 0#, 0#, 0#)

    ' --> compute working hours for aug 1-2, 2007
'    aTimeInfo = CheckDTR(cEmpID, _
'                         Array("2010-11-22", "2010-11-30", 0), _
'                         aEmpStat)

'    aTimeInfo = CheckDTR(cEmpID, _
'                         Array("2012-10-11", "2012-10-15", 0), _
'                         aEmpStat)

    aTimeInfo = CheckDTR(cEmpID, _
                         Array("2014-11-30", "2014-11-30", 0), _
                         aEmpStat)
                         
                         
                         
                         
'    ChkAdjustment oTempADO("EMPID"), _
'                  Array(oTempADO("PAYSTATUS"), _
                         oTempADO("RATE_AMT"), _
                         oTempADO("OLD_RATE"), _
                         oTempADO("OLD_COLA")), _
                         oTempADO("NEW_OLA")), _
'                  Array(aPeriodInfo(0), aPeriodInfo(1), aPeriodInfo(2) - aLeaveInfo(1)), _
'                  Array(oTempADO("EMP_STAT"), oTempADO("WAP"))

'                    If aTimeInfo(7) <> 0 Then
'                        aPayInfo(7) = Round((aTimeInfo(7) * oTempADO("RATE_AMT")) + (oTempADO("COLA_AMT") * aTimeInfo(7)), 2)                            ' --> holiday pay
'                    End If

'    If aRateInfo(0) = 0 Then
        ' --> gross using d old rate
        '    nGrossAmt = RegPay +
        '                RegOTPay +
        '                NDiffPay +
        '                NDiffOTPay +
        '                COLA
        aPayInfo(0) = Round(aTimeInfo(0) * aRateInfo(2), 2) + _
                      Round(aTimeInfo(1) * ((aRateInfo(2) / 8) * 1.25), 2) + _
                      Round(aTimeInfo(3) * (aRateInfo(2) * 1.1), 2) + _
                      Round(aTimeInfo(4) * ((aRateInfo(2) / 8) * 1.1 * 1.25), 2) + _
                      Round((aRateInfo(4) * (aTimeInfo(3) + aTimeInfo(0))), 2)
        'Holiday
        If aTimeInfo(7) <> 0 Then
            'aPayInfo(4) = Round((2 * aRateInfo(2)) + (aRateInfo(3) * 2), 2)                            ' --> holiday pay
            aPayInfo(4) = Round((aTimeInfo(7) * aRateInfo(2)) + (aRateInfo(4) * aTimeInfo(7)), 2)                          ' --> holiday pay
'        Else
'            MsgBox "stop"
        End If
        
        aPayInfo(0) = (aPayInfo(0) + aPayInfo(4))
        


        ' --> sa net using d old rate
        '    nNetAmt = SARegOTPay +
        '              SANDiffOTPay +
        '              SunCola +
        '              SunPay +
        '              SunOTPay +
        '              SunNDPay +
        '              SunNDOTPay
            aPayInfo(1) = Round(aTimeInfo(2) * ((aRateInfo(2) / 8) * 1.25), 2) + _
                      Round(aTimeInfo(12) * ((aRateInfo(2) / 8) * 1.1 * 1.25), 2) + _
                      Round((aRateInfo(4) * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
                      Round(aTimeInfo(5) * ((aRateInfo(2) / 8) * 1.3), 2) + _
                      Round(aTimeInfo(6) * ((aRateInfo(2) / 8) * 1.3 * 1.3), 2) + _
                      Round(aTimeInfo(13) * ((aRateInfo(2) / 8) * 1.3 * 1.1), 2) + _
                      Round(aTimeInfo(14) * ((aRateInfo(2) / 8) * 1.3 * 1.1 * 1.3), 2)


        ' --> gross using the new rate
        aPayInfo(2) = Round(aTimeInfo(0) * aRateInfo(1), 2) + _
                      Round(aTimeInfo(1) * ((aRateInfo(1) / 8) * 1.25), 2) + _
                      Round(aTimeInfo(3) * (aRateInfo(1) * 1.1), 2) + _
                      Round(aTimeInfo(4) * ((aRateInfo(1) / 8) * 1.1 * 1.25), 2) + _
                      Round((aRateInfo(3) * (aTimeInfo(3) + aTimeInfo(0))), 2)
                      
        If aTimeInfo(7) <> 0 Then
            'aPayInfo(4) = Round((2 * aRateInfo(1)) + (aRateInfo(4) * 2), 2)                            ' --> holiday pay
            aPayInfo(4) = Round((aTimeInfo(7) * aRateInfo(1)) + (aRateInfo(3) * aTimeInfo(7)), 2)                            ' --> holiday pay
'        Else
'            MsgBox "stop"
        End If
        
        aPayInfo(2) = (aPayInfo(2) + aPayInfo(4))

        ' --> sa net using the new rate
        '    nNetAmt = SARegOTPay +
        '              SANDiffOTPay +
        '              SunCola +
        '              SunPay +
        '              SunOTPay +
        '              SunNDPay +
        '              SunNDOTPay
        aPayInfo(3) = Round(aTimeInfo(2) * ((aRateInfo(1) / 8) * 1.25), 2) + _
                      Round(aTimeInfo(12) * ((aRateInfo(1) / 8) * 1.1 * 1.25), 2) + _
                      Round((aRateInfo(3) * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
                      Round(aTimeInfo(5) * ((aRateInfo(1) / 8) * 1.3), 2) + _
                      Round(aTimeInfo(6) * ((aRateInfo(1) / 8) * 1.3 * 1.3), 2) + _
                      Round(aTimeInfo(13) * ((aRateInfo(1) / 8) * 1.3 * 1.1), 2) + _
                      Round(aTimeInfo(14) * ((aRateInfo(1) / 8) * 1.3 * 1.1 * 1.3), 2)
                      
'    End If
    
    
    ChkAdjustment = Array(Round(aPayInfo(0) - aPayInfo(2), 2), _
                          Round(aPayInfo(1) - aPayInfo(3), 2))
End Function


'old backup
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
    If (gCompanyID = "0007") Or (gCompanyID = "0003") Then
        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    Else
        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#)
    End If
    
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

'Function CheckDTR(ByVal cEmpID As String, _
'                  aPeriodInfo As Variant, _
'                  ByVal aEmpStat As Variant) As Variant
'    Dim cSqlStmt As String, _
'        oDTRRSet As New ADODB.Recordset, _
'        aTimeInfo As Variant, _
'        nCount As Integer, _
'        cDateEnd As String, _
'        cParam As String, _
'        lHol1 As Boolean, _
'        lHol2 As Boolean
'
'    Dim oRecordSet As New ADODB.Recordset
'    '2010-04-19
'    '16-regot
'    '17-saregot
'    '18-ndregot
'    '19-ndsaregot
'    If (gCompanyID = "0007") Or (gCompanyID = "0003") Then
'        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    Else
'        'aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#)
'        aTimeInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0, 0, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    End If
'
'    If oTempADO("active") > 0 Then
'        cDateEnd = Format(IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd")
'
'        cSqlStmt = "select count(holidayid) as tot_day  from PA4329 " & _
'                   "where (date between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & cDateEnd & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(cDateEnd) & ") and (fix_day=1))"
'    Else
'        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
'                   "where (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1))"
'    End If
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, objdbRs, False
'        nCount = objdbRs("tot_day")
'
'
'    ' --> for regular employee only
'    cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
'               "where ((date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
'               "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1)))" & _
'                   " and (tag=1)"
''        Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        If aEmpStat(0) <> 2 Then
'            nCount = nCount - objdbRs("tot_day")
'        End If
'    End If
'
'    '20101-01-13
''    If aEmpStat(0) = 2 Then
''        ' --> for regular employee only
''        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
''                   "where ((date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
''                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1)))" & _
''                       " and (tag=1)"
'''        Script2File cSqlStmt
''        OpenQueryDNS cSqlStmt, objdbRs, False
''        If objdbRs.RecordCount > 0 Then
''            If aEmpStat(0) <> 2 Then
''                nCount = nCount - objdbRs("tot_day")
''            End If
''        Else
''            nCount = 0
''        End If
''    End If
'
'    ' hired date between the selected period...
'    If (DateDiff("d", aPeriodInfo(0), oTempADO("date_hire")) >= 0) And (DateDiff("d", oTempADO("date_hire"), aPeriodInfo(1)) >= 0) Then
'        cSqlStmt = "select count(holidayid) as tot_day from PA4329 " & _
'                   "where (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(oTempADO("date_hire")) & ") and (fix_day=1))"
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            nCount = nCount - objdbRs("tot_day")
'        End If
'    End If
'
'    If nCount < 0 Then nCount = 0
''    '2010-04-19
''    cSqlStmt = "select EMPID, PERIODID, DATE, SHIFTID, " & _
''               "  sum(reg_hr/8) as reg_day, sum(reg_ot_hr) as reg_ot, sum(sa_reg_ot) as sa_reg_ot, " & _
''               "  sum(nd_hr/8) as nd_day, sum(nd_ot_hr) as nd_ot, sum(sa_nd_ot) as sa_nd_ot, " & _
''               "  sum(sun_hr) as sun_hr, sum(sun_ot_hr) as sun_ot, " & _
''               "  sum(sun_nd) as sun_nd, sum(sun_nd_ot) as sun_nd_ot, " & _
''               "  sum(inc_hr) as inc_hr " & _
''               "From di36770 " & _
''               "where (empid=" & cQuote & cEmpID & cQuote & ") " & _
''               "  and (date = " & cQuote & "2010-04-09" & cQuote & " ) " & _
''               "group by empid "
'''    Script2File cSqlStmt
''    OpenQueryDNS cSqlStmt, oDTRRSet, False
''    If oDTRRSet.RecordCount > 0 Then
''        If aEmpStat(0) > 0 And aEmpStat(2) = 0 Then
''            aTimeInfo(16) = oDTRRSet("reg_ot")                               ' --> Reg OT
''            aTimeInfo(17) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
''            aTimeInfo(18) = oDTRRSet("nd_ot")                                ' --> NDiff OT
''            aTimeInfo(19) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
'''        Else
'''            MsgBox "stop"
''        End If
''    End If
'
'    cSqlStmt = "select EMPID, PERIODID, DATE, SHIFTID, " & _
'               "  sum(reg_hr/8) as reg_day, sum(reg_ot_hr) as reg_ot, sum(sa_reg_ot) as sa_reg_ot, " & _
'               "  sum(nd_hr/8) as nd_day, sum(nd_ot_hr) as nd_ot, sum(sa_nd_ot) as sa_nd_ot, " & _
'               "  sum(sun_hr) as sun_hr, sum(sun_ot_hr) as sun_ot, " & _
'               "  sum(sun_nd) as sun_nd, sum(sun_nd_ot) as sun_nd_ot, " & _
'               "  sum(inc_hr) as inc_hr " & _
'               "From di36770 " & _
'               "where (empid=" & cQuote & cEmpID & cQuote & ") " & _
'               "  and (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(aPeriodInfo(1), "yyyy-mm-dd") & cQuote & ") " & _
'               "group by empid "
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oDTRRSet, False
'    If oDTRRSet.RecordCount > 0 Then
'
'        aTimeInfo(0) = oDTRRSet("reg_day")      ' --> Reg Day
'        aTimeInfo(3) = oDTRRSet("nd_day")       ' --> NDiff Day
'        aTimeInfo(5) = oDTRRSet("sun_hr")       ' --> Sunday
'        aTimeInfo(6) = oDTRRSet("sun_ot")       ' --> Sunday OT
'        aTimeInfo(13) = oDTRRSet("sun_nd")      ' --> Sunday ND
'        aTimeInfo(14) = oDTRRSet("sun_nd_ot")   ' --> Sunday NDiff OT
'        aTimeInfo(15) = oDTRRSet("inc_hr")   ' --> Incentive Hour
'
'        If aEmpStat(2) > 0 Then
''            aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
''            aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
''            aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
''            aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
''            aTimeInfo(7) = 0                                                ' --> No Holiday
'
'            If gCompanyID = "0002" Then
'
'                If lAudit = 1 Then
'                    aTimeInfo(1) = oDTRRSet("reg_ot") + oDTRRSet("sa_reg_ot")       ' --> Reg OT
'                    aTimeInfo(4) = oDTRRSet("nd_ot") + oDTRRSet("sa_nd_ot")         ' --> NDiff OT
'                    aTimeInfo(2) = 0                                                ' --> SA Reg OT
'                    aTimeInfo(12) = 0                                               ' --> SA NDiff OT
'                Else
'                    aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
'                    aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
'                    aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
'                    aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
'                End If
'            Else
'                aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
'                aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
'                aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
'                aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
'            End If
'                aTimeInfo(7) = 0                                                ' --> No Holiday
'        Else
'            If gCompanyID = "0002" Then
'                If lAudit = 1 Then
'                    aTimeInfo(1) = oDTRRSet("reg_ot") + oDTRRSet("sa_reg_ot")       ' --> Reg OT
'                    aTimeInfo(4) = oDTRRSet("nd_ot") + oDTRRSet("sa_nd_ot")         ' --> NDiff OT
'                    aTimeInfo(2) = 0                                                ' --> SA Reg OT
'                    aTimeInfo(12) = 0                                              ' --> SA NDiff OT
'                Else
'                    aTimeInfo(1) = oDTRRSet("reg_ot")                               ' --> Reg OT
'                    aTimeInfo(4) = oDTRRSet("nd_ot")                                ' --> NDiff OT
'                    aTimeInfo(2) = oDTRRSet("sa_reg_ot")                            ' --> SA Reg OT
'                    aTimeInfo(12) = oDTRRSet("sa_nd_ot")                            ' --> SA NDiff OT
'                End If
'            Else
'                aTimeInfo(1) = oDTRRSet("reg_ot")       ' --> Reg OT
'                aTimeInfo(2) = oDTRRSet("sa_reg_ot")    ' --> SA Reg OT
'                aTimeInfo(4) = oDTRRSet("nd_ot")        ' --> NDiff OT
'                aTimeInfo(12) = oDTRRSet("sa_nd_ot")    ' --> SA NDiff OT
'
'                'aTimeInfo(7) = IIf((aEmpStat(0) <> 0) And (Not ((aEmpStat(0) = 1) And (aEmpStat(1) = 1))), nCount, 0)
'
'            End If
'                aTimeInfo(7) = IIf((aEmpStat(0) <> 0) And (Not ((aEmpStat(0) = 1) And (aEmpStat(1) = 1))), nCount, 0)
'
'        End If
'    End If
'
''    aTimeInfo(16) = aTimeInfo(16) + aTimeInfo(8)                 'sun ot
''    aTimeInfo(20) = aTimeInfo(20) + aTimeInfo(8)                 'sun nd ot
'
'
'    If gCompanyID = "0002" Then
'        cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
'                   " left join di3670 b on a.empid=b.empid " & _
'                   " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & cEmpID & cQuote & " And a.sun_hr <> 0 And b.emp_stat <> 0 "
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            If objdbRs("emp_stat") = 1 Then
'                aTimeInfo(8) = objdbRs("sun_hr")
'            Else
'                aTimeInfo(7) = aTimeInfo(7) - 1
'                aTimeInfo(8) = objdbRs("sun_hr")
'            End If
'        Else
'            If aEmpStat(0) = 2 Then
'                aTimeInfo(7) = aTimeInfo(7) - 1
'                aTimeInfo(8) = 8
'            Else
'                aTimeInfo(8) = 0
'            End If
'
'        End If
'    Else
'
'        cSqlStmt = "select * from di3670 where empid = " & cQuote & cEmpID & cQuote
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            If objdbRs("emp_stat") <> 0 Then
'                cSqlStmt = " select a.sun_hr,a.sun_ot_hr,b.emp_stat from di36770 a " & _
'                           " left join di3670 b on a.empid=b.empid " & _
'                           " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & cEmpID & cQuote & " And a.sun_hr <> 0 And b.emp_stat <> 0 "
'                OpenQueryDNS cSqlStmt, objdbRs, False
'                If objdbRs.RecordCount > 0 Then
'                    lHol1 = True
'                    aTimeInfo(16) = objdbRs("sun_ot_hr")
'                    aTimeInfo(8) = objdbRs("sun_hr")
'                Else
'                    lHol1 = False
'                End If
'
'                cSqlStmt = " select sun_hr,sun_nd,sun_nd_ot,b.emp_stat from di36770 a " & _
'                           " left join di3670 b on a.empid=b.empid " & _
'                           " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & cEmpID & cQuote & " And a.sun_nd <> 0 And b.emp_stat <> 0 "
'                OpenQueryDNS cSqlStmt, objdbRs, False
'                If objdbRs.RecordCount > 0 Then
'                    lHol2 = True
'                    aTimeInfo(20) = objdbRs("sun_nd_ot")
'                    aTimeInfo(8) = aTimeInfo(8) + objdbRs("sun_nd")
'                Else
'                    lHol2 = False
'                End If
'
'                If lHol1 = True Or lHol2 = True Then
'                    aTimeInfo(7) = aTimeInfo(7) - 1
'                    aTimeInfo(8) = aTimeInfo(8)
'                Else
'                    aTimeInfo(7) = aTimeInfo(7) - 1
'                    aTimeInfo(8) = 0
'                End If
'            Else
'                aTimeInfo(8) = 0
'            End If
'        End If
'
'    End If
'
''    If gCompanyID = "0002" Then
''        cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                   " left join di3670 b on a.empid=b.empid " & _
''                   " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & cEmpID & cQuote & " And a.sun_hr <> 0 And b.emp_stat = 1 "
''        OpenQueryDNS cSqlStmt, objdbRs, False
''        If objdbRs.RecordCount > 0 Then
''
''            If objdbRs("emp_stat") = 1 Then
''                aTimeInfo(8) = 1
''
''            Else
''                aTimeInfo(8) = 0
''
''            End If
'''        Else
'''            If aEmpStat(0) = 2 Then
'''                aTimeInfo(8) = 1
'''                aTimeInfo(7) = aTimeInfo(7) - 1
'''            Else
'''                aTimeInfo(8) = 0
'''            End If
''        End If
''    Else
''        cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                   " left join di3670 b on a.empid=b.empid " & _
''                   " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & cEmpID & cQuote & " And a.sun_hr <> 0 And b.emp_stat <> 0 "
''        OpenQueryDNS cSqlStmt, objdbRs, False
''        If objdbRs.RecordCount > 0 Then
''
''            If objdbRs("emp_stat") = 1 Then
''                aTimeInfo(8) = 1
''
''            Else
''                aTimeInfo(8) = 0
''
''            End If
''        Else
''            aTimeInfo(8) = 0
''        End If
''
''    End If
'
'    CheckDTR = aTimeInfo
'
'    Set oDTRRSet = Nothing
'    Set oRecordSet = Nothing
'End Function



'Function ChkLeave(ByVal dStartDate As Date, dEndDate As Date) As Variant
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

Sub ProcessTransaction()
    Dim nCtr, _
        nPeriod As Integer, _
        cDedID, cSqlStmt, cDepid, cParam As String, _
        aTimeInfo As Variant, _
        oRSetDed As New ADODB.Recordset, oRset1 As New ADODB.Recordset, _
        nDedAmt As Double, nDedAmt2 As Double, nDedAmt3 As Double, nTotDed As Double, _
        n13mopay As Double, _
        nIncentive As Double, nTotDay As Double, nTotExempt As Double, _
        aPeriodInfo As Variant, aDedAmt As Variant, aPayInfo As Variant, aLeaveInfo As Variant, _
        lWith13Mo As Boolean, lAssess As Boolean, lIsWap As Boolean, lWithTax As Boolean, _
        aTmpTax As Variant, _
        nPayStat As Integer, _
        lAllDed As Boolean, _
        nHolCola As Double, _
        nBasicpay_tot As Double, _
        nDedTag As Integer, _
        lTrueTag As Boolean
        
    Dim oLeaveRSet As New ADODB.Recordset
        
    ' --> 20070814, for Payroll adjustment only (rate increase)
    Dim aAdjustment As Variant
        
    aAdjustment = Array(0#, 0#)
    ' (0)   -   Regular Adjustment
    ' (1)   -   SA Adjustment
    
    aPeriodInfo = Array("", "", 0, 0, 0, 0)
    ' (0)   -   Start Date
    ' (1)   -   End Date
    ' (2)   -   # of Holiday
    ' (3)   -   Working Days
    ' (4)   -   Annual Withholding Tax tag... 20070105
    ' (5)   -   Special Non-working Holiday....20091002
    
    aDedAmt = Array(0#, 0#, 0#, 0#, 0#, 0#, "")     ' --> for deduction purposes
    ' (0)   -   SSS ER
    ' (1)   -   SSS Premium
    ' (2)   -   Withholding Tax
    ' (3)   -   PhilHealth/Medicare PS
    ' (4)   -   PhilHealth/Medicare ES
    ' (5)   -   Medicare Total
    ' (6)   -   Loan Control Number - 20060816
    
    For nCtr = 0 To UBound(aTaxExempt)
        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
        cDedID = cDedID & aTaxExempt(nCtr) & ","
    Next nCtr
    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
    
    cParam = ListView1.SelectedItem.Text
    
    OpenQueryDNS "SELECT * FROM PA87260 WHERE PERIODID=" & cQuote & cParam & cQuote, objdbRs, False
    If objdbRs.RecordCount = 0 Then
    
        ' --> for special assessment only...
        If (Trim(gAssessID) <> "") Then
            cSqlStmt = "select DEF_AMT from pa3330 where dedid=" & cQuote & gAssessID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            nAssessAmt = IIf(objdbRs.RecordCount > 0, objdbRs("def_amt"), 0)
            If (MsgBox("Is there special assessment?", vbYesNo, App.Title) = vbYes) Then
                frmSpecialAssess.txtAmount.Text = nAssessAmt
                frmSpecialAssess.Show 1
                lAssess = True
            End If
        Else
            lAssess = False
        End If
        
        ' --> retrieve period info here...
        OpenQueryDNS "SELECT * FROM PA7730 WHERE PERIODID=" & cQuote & cParam & cQuote, objdbRs, False
        nPeriod = IIf(objdbRs.RecordCount > 0, objdbRs("STATUS"), 0)
        aPeriodInfo(0) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("DATE_START"), Now), "yyyy-mm-dd")
        aPeriodInfo(1) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("DATE_END"), Now), "yyyy-mm-dd")
        aPeriodInfo(2) = IIf(objdbRs.RecordCount > 0, objdbRs("holidays"), 0)
        aPeriodInfo(3) = IIf(objdbRs.RecordCount > 0, objdbRs("workindays"), 0)
        aPeriodInfo(4) = IIf(objdbRs.RecordCount > 0, objdbRs("wtax"), 0)
        
        ' --> count holiday here...
        cSqlStmt = "select * from PA4329 " & _
                   "where (date between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1))"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then aPeriodInfo(2) = objdbRs.RecordCount
        
'        cSqlStmt = "select * from PA4329 " & _
'                   "where (date between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & ") and (fix_day=1))"
'
''        Script2File cSqlStmt
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            While objdbRs.EOF
'
'                If objdbRs("withpay") = 0 Then
'                    aPeriodInfo(2) = aPeriodInfo(2) + 1
'                Else
'                    aPeriodInfo(5) = aPeriodInfo(5) + 1
'                End If
'
'                objdbRs.MoveNext
'            Wend
'        End If
'        'aPeriodInfo(2) = objdbRs.RecordCount
        
' --> remarked 20070814
'        cSqlStmt = "SELECT EMPID, SHIFTID, TAXID, CONCAT(LASTNAME,', ',FIRSTNAME,if(trim(mname)='','',concat(' ',left(mname,1),'. '))) as FULLNAME, FIRSTNAME, MNAME, LASTNAME, DEPID, POSID, ISUNION, " & _
'                   " RATE_AMT, COLA_AMT, COLA1215, POS_ALLOW, SSER1215, SSPREM1215, PS1215, ES1215, PAYSTATUS, EMP_STAT, WAP, DATE_RES, DATE_FIN, DATE_HIRE, " & _
'                   " MTD_TAXABLE, MTD_BASIC, MTD_GROSS, YTD_BASIC, YTD_GROSS, YTD_GROSS_SA, YTD_COLA, " & _
'                   " if(active>0,1,0) as active2, `ACTIVE`, SL_AVAIL, VL_AVAIL, SL_USE, VL_USE, PAGIBIGNO, SSNUM, PHEALTHNUM, TIN " & _
'                   "FROM DI3670 " & _
'                   "WHERE (paystatus=0) and (((ACTIVE=0) and ((date_hire<=" & cQuote & aPeriodInfo(0) & cQuote & ") or (date_hire between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ")))" & _
'                   " OR ((active=1) and ((date_res between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") or (date_res > " & cQuote & aPeriodInfo(1) & cQuote & "))) " & _
'                   " OR ((active=2) and ((date_fin between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") or (date_fin > " & cQuote & aPeriodInfo(1) & cQuote & ")))) " & _
'                   "order by active2, if(active>0,'',depid), emp_stat desc, fullname"
        
        ' --> retrieve daily wage earner here... 20070118
        cSqlStmt = "SELECT EMPID, SHIFTID, TAXID, CONCAT(LASTNAME,', ',FIRSTNAME,if(trim(mname)='','',concat(' ',left(mname,1),'. '))) as FULLNAME, FIRSTNAME, MNAME, LASTNAME, DEPID, POSID, ISUNION, " & _
                   " RATE_AMT, COLA_AMT, OLD_RATE, OLD_COLA, COLA1215, POS_ALLOW, SSER1215, SSPREM1215, PS1215, ES1215, PAYSTATUS, EMP_STAT, WAP, DATE_RES, DATE_FIN, DATE_HIRE, " & _
                   " MTD_TAXABLE, MTD_BASIC, MTD_GROSS, YTD_BASIC, YTD_GROSS, YTD_GROSS_SA, YTD_COLA, " & _
                   " if(active>0,1,0) as active2, `ACTIVE`, SL_AVAIL, VL_AVAIL, SL_USE, VL_USE, PAGIBIGNO, SSNUM, PHEALTHNUM, TIN,BACCNTNO,COSTCENTERID,WORKCENTERID " & _
                   "FROM DI3670 " & _
                   "WHERE (paystatus<>1) and (((ACTIVE=0) and ((date_hire<=" & cQuote & aPeriodInfo(0) & cQuote & ") or (date_hire between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ")))" & _
                   " OR (((active=1) or (active=3)) and ((date_res between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") or (date_res > " & cQuote & aPeriodInfo(1) & cQuote & "))) " & _
                   " OR ((active=2) and ((date_fin between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") or (date_fin > " & cQuote & aPeriodInfo(1) & cQuote & ")))) " & _
                   "order by paystatus, active2, if(active>0,'',depid), emp_stat desc, fullname"
'        Script2File cSqlStmt
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then
        
            ShowProgress 0
            
            While Not oTempADO.EOF
           
'                If oTempADO("empid") = "352822" Then MsgBox "hinto"
            
                ' --> For Emergency Manpower...
                If nPayStat <> oTempADO("paystatus") Then
                    nPayStat = oTempADO("paystatus")
                    nCtr = 0
                End If
            
                ' --> create seq_no here...
                If (oTempADO("active") = 0) Then
                    If (oTempADO("depid") <> cDepid) Then
                        lIsWap = False
                        cDepid = oTempADO("depid")
                        nCtr = 0
                    ElseIf oTempADO("emp_stat") = 0 Then
                        If Not lIsWap Then
                            lIsWap = True
                            nCtr = 0
                        End If
                    End If
                Else
                    If (cDepid <> "999") Then
                        lIsWap = False
                        cDepid = "999"
                        nCtr = 0
                    ElseIf oTempADO("emp_stat") = 0 Then
                        If Not lIsWap Then
                            lIsWap = True
                            nCtr = 0
                        End If
                    End If
                End If
                nCtr = nCtr + 1
                
                aAdjustment = Array(0#, 0#)
                
                n13mopay = 0
                nTotDay = 0
                nIncentive = 0
                
'                '2010-04-19
'                If (gCompanyID = "0007") Or (gCompanyID = "0003") Then
'
'                    aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'                Else
''                    aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'                    aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'                End If

                '2013-07-11
                aPayInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                
                '    (0) = reg pay
                '    (1) = reg ot pay
                '    (2) = sa reg ot pay
                '    (3) = ndiff pay
                '    (4) = ndiff ot pay
                '    (5) = Sunday pay
                '    (6) = Sunday OT pay
                '    (7) = Legal Holiday pay
                '    (8) = Special Non-working Holiday
                '    (9) = basic pay
                '    (10) = gross pay
                '    (11) = sa net pay
                '    (12) = sa ndiff ot pay
                '    (13) = sun ndiff pay
                '    (14) = sun ndiff ot pay
                '    (15) = Incentive Hour
                '    (16) = Holiday OT hrs
'                '2010-04-19
'                '    (17) = sun ndiff pay
'                '    (18) = sun ndiff ot pay
'                '    (19) = Incentive Hour
'                '    (20) = Holiday OT hrs
                aLeaveInfo = Array(0#, 0#)
                                
                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , , "Retrieving data of " & oTempADO("FULLNAME")
                
                ' --> compute leave avail here...
                cSqlStmt = "select " & _
                           "  if(date_start<=" & cQuote & aPeriodInfo(0) & cQuote & "," & cQuote & aPeriodInfo(0) & cQuote & ",date_start) as date_start, " & _
                           "  if(date_end>=" & cQuote & aPeriodInfo(1) & cQuote & "," & cQuote & aPeriodInfo(1) & cQuote & ",date_end) as date_end, " & _
                           "  tag , paytag " & _
                           "From pa367583 " & _
                           "where (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (status=1) " & _
                           "  and (paytag=0) and (tag in (0,1" & IIf(gCompanyID = 2, ",6", "") & ")) " & _
                           "  and ((date_start between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
                           "    or (date_end between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ")) "
'                Script2File cSqlStmt
                OpenQueryDNS cSqlStmt, oLeaveRSet, False
                If oLeaveRSet.RecordCount > 0 Then
'                    aLeaveInfo = ChkLeave(objdbRs("date_start"), objdbRs("date_end"))
                    aLeaveInfo = ChkLeave(oLeaveRSet)
                    nTotDay = aLeaveInfo(0) - aLeaveInfo(1)
                    nIncentive = Round(nTotDay * ((oTempADO("RATE_AMT") + oTempADO("COLA_AMT")) / IIf(oTempADO("paystatus") = 0, 1, 26.08)), 2)
                End If
                
                ' --> compute "predata" here...
                If oTempADO("active") > 0 Then
                nIncentive = 0
                    nTotDay = (DateDiff("d", "01/01/" & Year(Now), IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")))) / 365
                    '20120616 revise
                    'nIncentive = Round(nIncentive + ((nTotDay * (oTempADO("SL_AVAIL") + oTempADO("VL_AVAIL") - oTempADO("SL_USE") - oTempADO("VL_USE"))) * (oTempADO("RATE_AMT") / IIf(oTempADO("paystatus") = 0, 1, 26.08))), 2)
                    nIncentive = Round(nIncentive + ((((nTotDay * (oTempADO("SL_AVAIL") + oTempADO("VL_AVAIL"))) - (oTempADO("SL_USE") + oTempADO("VL_USE"))) * (oTempADO("RATE_AMT") + oTempADO("COLA_AMT")))), 2)
                    'nIncentive = (nIncentive - (oTempADO("SL_USE") + oTempADO("VL_USE")))
                    'nIncentive = nIncentive * oTempADO("RATE_AMT")
                    
                End If
                
                ' --> 20060919 - use summary instead...
'                ' --> insert data from bio-clock here...
'                aTimeInfo = ComputeDays(oTempADO("EMPID"), _
'                                        Array(aPeriodInfo(0), aPeriodInfo(1), aPeriodInfo(2) - aLeaveInfo(1)), _
'                                        Array(oTempADO("EMP_STAT"), oTempADO("WAP")))
                
               
                ' --> deactivated 20070831
'                aAdjustment = ChkAdjustment(oTempADO("EMPID"), _
'                                            Array(oTempADO("PAYSTATUS"), oTempADO("RATE_AMT"), oTempADO("OLD_RATE"), oTempADO("COLA_AMT"), oTempADO("OLD_COLA")), _
'                                            Array(aPeriodInfo(0), aPeriodInfo(1), aPeriodInfo(2) - aLeaveInfo(1)), _
'                                            Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))

'                ' --> activate 20110627
                aAdjustment = ChkAdjustment(oTempADO("EMPID"), _
                                            Array(oTempADO("PAYSTATUS"), oTempADO("RATE_AMT"), oTempADO("OLD_RATE"), oTempADO("COLA_AMT"), oTempADO("OLD_COLA")), _
                                            Array(aPeriodInfo(0), aPeriodInfo(1), aPeriodInfo(2) - aLeaveInfo(1)), _
                                            Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))
                                            

                aTimeInfo = CheckDTR(oTempADO("EMPID"), _
                                     Array(aPeriodInfo(0), aPeriodInfo(1), aPeriodInfo(2) - aLeaveInfo(1)), _
                                     Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))
'                '2009-10-02
'                aTimeInfo = CheckDTR(oTempADO("EMPID"), _
'                                     Array(aPeriodInfo(0), aPeriodInfo(1), (aPeriodInfo(2) + aPeriodInfo(5)) - aLeaveInfo(1)), _
'                                     Array(oTempADO("EMP_STAT"), oTempADO("WAP"), oTempADO("PAYSTATUS")))


                If oTempADO("paystatus") <> 1 Then   ' --> Daily/Emergency
                
                    aPayInfo(0) = Round(aTimeInfo(0) * oTempADO("RATE_AMT"), 2)                             ' --> reg pay
                    aPayInfo(1) = Round(aTimeInfo(1) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)              ' --> reg ot pay
                    aPayInfo(2) = Round(aTimeInfo(2) * ((oTempADO("RATE_AMT") / 8) * 1.25), 2)              ' --> sa reg ot pay
                    
                    aPayInfo(3) = Round(aTimeInfo(3) * (oTempADO("RATE_AMT") * 1.1), 2)                     ' --> ndiff pay
                    aPayInfo(4) = Round(aTimeInfo(4) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)        ' --> ndiff ot pay
                    aPayInfo(12) = Round(aTimeInfo(12) * ((oTempADO("RATE_AMT") / 8) * 1.1 * 1.25), 2)      ' --> sa ndiff ot pay
                    
                    aPayInfo(5) = Round(aTimeInfo(5) * ((oTempADO("RATE_AMT") / 8) * 1.3), 2)               ' --> sun pay
                    aPayInfo(6) = Round(aTimeInfo(6) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.3), 2)         ' --> sun ot pay
                    
                    aPayInfo(13) = Round(aTimeInfo(13) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1), 2)       ' --> sun ndiff pay
                    aPayInfo(14) = Round(aTimeInfo(14) * ((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1 * 1.3), 2) ' --> sun ndiff ot pay
                    
                    
                                              
                    If gCompanyID <> "0004" Or gCompanyID <> "0001" Then
                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)
                    Else
                        aPayInfo(9) = aPayInfo(0) + Round(aTimeInfo(3) * (oTempADO("RATE_AMT")), 2)
                    End If
'                        aPayInfo(9) = aPayInfo(0) + aPayInfo(3)
                    
                    If aTimeInfo(7) <> 0 Then
                        aPayInfo(7) = Round((aTimeInfo(7) * oTempADO("RATE_AMT")) + (oTempADO("COLA_AMT") * aTimeInfo(7)), 2)                            ' --> holiday pay
                    End If
                    
'                    'special rest holiday 20110821
'                    If aTimeInfo(8) <> 0 Then
'
'                        aPayInfo(8) = Round((aTimeInfo(8) * (oTempADO("RATE_AMT") / 8 * 0.2)), 2)                        ' --> holiday pay
'
'                        If gCompanyID = "0002" Then
'                            If oTempADO("emp_stat") = 2 Then
'                                'aPayInfo(8) = Round(((((aTimeInfo(8) * (oTempADO("RATE_AMT")) / 8) + (oTempADO("COLA_AMT") * aTimeInfo(8) / 8)) * 1.5)), 2)                      ' --> holiday pay
'
'                                aPayInfo(8) = Round(aTimeInfo(8) * ((oTempADO("RATE_AMT") / 8) * 1.5) + oTempADO("COLA_AMT"), 2)               ' --> sun pay
'
''                                aPayInfo(8) = Round(((aTimeInfo(8) *.2), 2)                   ' --> holiday pay
''                                aPayInfo(8) = Round(((aTimeInfo(8) * (oTempADO("RATE_AMT") / 8)) + oTempADO("COLA_AMT")) * 1.5, 2)                   ' --> holiday pay
'                            End If
'                        End If
                                                
'                        aPayInfo(7) = Round((aPayInfo(7) + aPayInfo(8)), 2)
'                        aTimeInfo(7) = aTimeInfo(7) + 1
'
'                        aPayInfo(16) = Round(aTimeInfo(16) * (((oTempADO("RATE_AMT") / 8) * 1.3 * 1.3) * 0.5), 2)               ' --> special holiday sun ot pay
'                        aPayInfo(20) = Round(aTimeInfo(20) * (((oTempADO("RATE_AMT") / 8) * 1.3 * 1.1 * 1.3) * 0.5), 2)         ' --> special holiday sun nd ot
'
'                        aPayInfo(20) = aPayInfo(20) + aPayInfo(16)
'
'                    End If
  
                    ' --> addendum, 20080313 Incentive Hour
                    If gCompanyID = "0002" Then
                        aPayInfo(15) = Round(aTimeInfo(15) * (oTempADO("RATE_AMT") / 8), 2)                     ' --> Incentive Pay
                    Else
                        aPayInfo(15) = aTimeInfo(15)                                                             ' --> Incentive Pay
                    End If
                    
                    If lExtension Then
                    
                        '201307-11
'                        If gCompanyID = "0001" Or gCompanyID = "0006" Or gCompanyID = "0005" Then
                    
                        If gCompanyID = "0001" Or gCompanyID = "0006" Then
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
                                           IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                                           
'                            If gCompanyID = "0005" Then
'                                aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                                aPayInfo(11) = 0
'                            End If
                                                                   
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
                                           IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) + _
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
                                           Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
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
                        '2013-7-11
'                        If gCompanyID = "0001" Or gCompanyID = "0006" Or gCompanyID = "0005" Then
                    
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
                                           IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) + _
                                           nIncentive + _
                                           aAdjustment(0), 2)
                                           
                            '    nNetAmt = SunCola +
                            '              SunPay +
                            '              SunOTPay +
                            '              SunNDPay +
                            '              SunNDOTPay +
                            '              SAAdjPay         --> ala p 2 computation d2...
                            aPayInfo(11) = Round(Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                                           
'                            If gCompanyID = "0005" Or gCompanyID = "0001" Or gCompanyID = "0006" Then
'                                If nPayStat = 2 Then
'                                    aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                                    aPayInfo(11) = 0
'                                End If
'
'                            End If
                            
                            If gCompanyID = "0001" Or gCompanyID = "0006" Then
                                If nPayStat = 2 Then
                                    aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                                    aPayInfo(11) = 0
                                End If
                            Else
                                If gCompanyID = "0005" Then
                                    aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
                                    aPayInfo(11) = 0
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
                                           IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) + _
                                           Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) + _
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
                                           Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2) + _
                                           aPayInfo(5) + _
                                           aPayInfo(6) + _
                                           aPayInfo(13) + _
                                           aPayInfo(14) + _
                                           aPayInfo(20) + _
                                           aAdjustment(1), 2)
                        End If
                    End If
                             
                Else    ' --> monthly
                    If aTimeInfo(0) = 0 Then
                        aPayInfo(0) = 0
                    Else
'                        aPayInfo(0) = Round((oTempADO("rate_amt") / 2) - ((oTempADO("rate_amt") / 26.08) * (aPeriodInfo(3) - aTimeInfo(0))), 2)
                    End If
                    aPayInfo(9) = aPayInfo(0)
                    aPayInfo(10) = Round(aPayInfo(0) + _
                                   aPayInfo(1) + _
                                   aPayInfo(3) + _
                                   aPayInfo(7) + _
                                   nIncentive, 2)
                End If
                         
                ' --> revised 20070105, 20070831 - add emergency here
                If (oTempADO("paystatus") = 2) Or (aPeriodInfo(4) > 0) Then
                    n13mopay = 0
                Else
                    ' --> 13th month here...
                    If (oTempADO("active") > 0) And (oTempADO("wap") = 0) Then
                        ' --> revised 20070122
                            lWith13Mo = True
                        If lWith13Mo Then
                        
                        
                            cSqlStmt = " SELECT empid,periodid,sum(basicpay) as basicpay FROM pah87260 " & _
                                       " where periodid in ( " & _
                                       " SELECT periodid FROM pa7730 where year(date_start) = " & cQuote & Year(Now) & cQuote & " ) and " & _
                                       " empid = " & cQuote & oTempADO("empid") & cQuote & _
                                       " group by empid " & _
                                       " order by empid "
                            Script2File cSqlStmt
                            OpenQueryDNS cSqlStmt, objdbRs, False
                            'iIf(objdbRs.RecordCount > 0,objdbRs("basicpay"),0)   Then
                            
                            
                        
                        
                            If (oTempADO("EMP_STAT") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)) Then
                                ' --> 20060724 n13MOP7ay = (oTempADO("YTD_BASIC") - oTempADO("YTD_COLA") + aPayInfo(9) - (oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0)))) / 12
                                
                                'kung ang period na i proprocess mo ay close na at ni retoke mo lang dapat ito ang gamitin
'                                n13mopay = Round((oTempADO("YTD_BASIC")) / 12, 2)
                                
                                'n13mopay = Round((oTempADO("YTD_BASIC") + aPayInfo(9)) / 12, 2)
                                
                                n13mopay = Round((IIf(objdbRs.RecordCount <> 0, objdbRs("basicpay"), 0) + aPayInfo(9)) / 12, 2)
                           
'                                MsgBox n13mopay
                            ElseIf oTempADO("EMP_STAT") = 2 Then
                            
                                n13mopay = Round((oTempADO("YTD_GROSS") + oTempADO("YTD_GROSS_SA") - oTempADO("YTD_COLA") + aPayInfo(10) + aPayInfo(11) - (oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0)))) / 12, 2)
'                                MsgBox n13mopay
                            Else
                                n13mopay = 0
                            End If
                        End If
                    Else
                        n13mopay = 0
                    End If
                End If
                
''for Suneast audit
'                If gCompanyID = "0006" Then
'                    If oTempADO("wap") <> 1 Then
'                        aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                    Else
'                        If aPayInfo(5) = 0 Or aPayInfo(5) = 0 Then
'                            aPayInfo(11) = 0
'                        Else
'
''                            aPayInfo(11) = aPayInfo(5) + _
''                                       aPayInfo(6) + _
''                                       aPayInfo(13) + _
''                                       aPayInfo(14) + _
''                                       Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2)
'
'                        'suneast audit
'                            aPayInfo(11) = aPayInfo(5) + _
'                                       aPayInfo(6) + _
'                                       aPayInfo(13) + _
'                                       aPayInfo(14) + _
'                                       aPayInfo(4) + _
'                                       aPayInfo(12) + _
'                                       Round((oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8)), 2)
'
'                            aPayInfo(10) = aPayInfo(10) + aPayInfo(11)
'                        End If
'                    End If
'                End If
                
'                 --> days worked transaction saved here...
                cSqlStmt = "INSERT INTO PA87260(PERIODID,PERIOD_STAT,P_DAY,P_HOLIDAY,SEQ_NO,EMPID,FULLNAME,FIRSTNAME,MNAME,LASTNAME,date_hire,DATE_RES," & _
                           "WAP,EMP_STAT,`ACTIVE`,PAYSTATUS,TAXID,DEPID,RATE_AMT,COLA_AMT,COLA,SUN_COLA,POSID,POS_ALLOW," & _
                           "BASIC1215,GROSS16231,TAX1215,SSER1215,SSPREM1215,PS1215,ES1215,COLA1215," & _
                           "REG_DAY,REG_OT_HR,SA_REG_OT,NDIFF_DAY,NDIFF_OT_HR,SA_NDIFF_OT,SUN_HR,SUN_OT,sun_nd,sun_nd_ot,`HOLIDAY`," & _
                           "REG_PAY,REG_OT_PAY,SA_REG_PAY,NDIFF_PAY,NDIFF_OT_PAY,SA_NDIFF_PAY,SUN_PAY,SUN_OT_PAY,sun_nd_pay,sun_nd_ot_pay,HOL_PAY," & _
                           "BASICPAY,GROSS_PAY,SA_NET_PAY,LEAVE_PAY,M13PAY,ADJ_PAY,SA_ADJ_PAY,PAGIBIGNO,SSSNUM,PHEALTHNUM,TINNUM,INC_HR,INC_PAY,OTHER_PAY,BACCNTNO,COSTCENTERID,WORKCENTERID)VALUES(" & _
                           cQuote & cParam & cQuote & "," & nPeriod & "," & aPeriodInfo(3) & "," & aPeriodInfo(2) & "," & _
                           nCtr & "," & cQuote & oTempADO("EMPID") & cQuote & "," & _
                           cQuote & DecodeStr(oTempADO("FULLNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("FIRSTNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("MNAME")) & cQuote & "," & cQuote & oTempADO("LASTNAME") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(IIf((oTempADO("active") = 1) Or (oTempADO("active") = 3), oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd") & cQuote & "," & _
                           oTempADO("WAP") & "," & oTempADO("EMP_STAT") & "," & oTempADO("ACTIVE") & "," & oTempADO("PAYSTATUS") & "," & _
                           cQuote & oTempADO("TAXID") & cQuote & "," & cQuote & oTempADO("DEPID") & cQuote & "," & _
                           oTempADO("RATE_AMT") & "," & _
                           oTempADO("COLA_AMT") & "," & Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) & "," & Round(oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8), 2) & "," & _
                           cQuote & oTempADO("POSID") & cQuote & "," & IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) & "," & _
                           oTempADO("MTD_BASIC") & "," & oTempADO("MTD_GROSS") & "," & oTempADO("MTD_TAXABLE") & "," & oTempADO("SSER1215") & "," & oTempADO("SSPREM1215") & "," & oTempADO("PS1215") & "," & oTempADO("ES1215") & "," & oTempADO("COLA1215") & "," & _
                           aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & aTimeInfo(5) & "," & aTimeInfo(6) & "," & aTimeInfo(13) & "," & aTimeInfo(14) & "," & aTimeInfo(7) & "," & _
                           aPayInfo(0) & "," & aPayInfo(1) & "," & aPayInfo(2) & "," & aPayInfo(3) & "," & aPayInfo(4) & "," & aPayInfo(12) & "," & aPayInfo(5) & "," & aPayInfo(6) & "," & aPayInfo(13) & "," & aPayInfo(14) & "," & aPayInfo(7) & "," & _
                           aPayInfo(9) & "," & aPayInfo(10) + n13mopay & "," & aPayInfo(11) & "," & nIncentive & "," & n13mopay & "," & _
                           aAdjustment(0) & "," & aAdjustment(1) & "," & _
                           cQuote & oTempADO("PAGIBIGNO") & cQuote & "," & cQuote & oTempADO("SSNUM") & cQuote & "," & _
                           cQuote & IIf(Trim(oTempADO("PHEALTHNUM")) = "", oTempADO("SSNUM"), oTempADO("PHEALTHNUM")) & cQuote & "," & cQuote & oTempADO("TIN") & cQuote & "," & _
                           aTimeInfo(15) & "," & aPayInfo(15) & "," & aPayInfo(20) & "," & cQuote & oTempADO("BACCNTNO") & cQuote & "," & cQuote & oTempADO("COSTCENTERID") & cQuote & "," & cQuote & oTempADO("WORKCENTERID") & cQuote & ")"
                           

''******************* for audit purpose pero ibabalik din sa dati - Suneast
'                cSqlStmt = "INSERT INTO PA87260(PERIODID,PERIOD_STAT,P_DAY,P_HOLIDAY,SEQ_NO,EMPID,FULLNAME,FIRSTNAME,MNAME,LASTNAME,date_hire,DATE_RES," & _
'                           "WAP,EMP_STAT,`ACTIVE`,PAYSTATUS,TAXID,DEPID,RATE_AMT,COLA_AMT,COLA,SUN_COLA,POSID,POS_ALLOW," & _
'                           "BASIC1215,GROSS16231,TAX1215,SSER1215,SSPREM1215,PS1215,ES1215,COLA1215," & _
'                           "REG_DAY,REG_OT_HR,SA_REG_OT,NDIFF_DAY,NDIFF_OT_HR,SA_NDIFF_OT,SUN_HR,SUN_OT,sun_nd,sun_nd_ot,`HOLIDAY`," & _
'                           "REG_PAY,REG_OT_PAY,SA_REG_PAY,NDIFF_PAY,NDIFF_OT_PAY,SA_NDIFF_PAY,SUN_PAY,SUN_OT_PAY,sun_nd_pay,sun_nd_ot_pay,HOL_PAY," & _
'                           "BASICPAY,GROSS_PAY,SA_NET_PAY,LEAVE_PAY,M13PAY,ADJ_PAY,SA_ADJ_PAY,PAGIBIGNO,SSSNUM,PHEALTHNUM,TINNUM,INC_HR,INC_PAY,OTHER_PAY)VALUES(" & _
'                           cQuote & cParam & cQuote & "," & nPeriod & "," & aPeriodInfo(3) & "," & aPeriodInfo(2) & "," & _
'                           nCtr & "," & cQuote & oTempADO("EMPID") & cQuote & "," & _
'                           cQuote & DecodeStr(oTempADO("FULLNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("FIRSTNAME")) & cQuote & "," & cQuote & DecodeStr(oTempADO("MNAME")) & cQuote & "," & cQuote & oTempADO("LASTNAME") & cQuote & "," & _
'                           cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(IIf((oTempADO("active") = 1) Or (oTempADO("active") = 3), oTempADO("date_res"), oTempADO("date_fin")), "yyyy-mm-dd") & cQuote & "," & _
'                           oTempADO("WAP") & "," & oTempADO("EMP_STAT") & "," & oTempADO("ACTIVE") & "," & oTempADO("PAYSTATUS") & "," & _
'                           cQuote & oTempADO("TAXID") & cQuote & "," & cQuote & oTempADO("DEPID") & cQuote & "," & _
'                           oTempADO("RATE_AMT") & "," & _
'                           oTempADO("COLA_AMT") & "," & Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2) & "," & Round(oTempADO("COLA_AMT") * ((aTimeInfo(5) + aTimeInfo(13)) / 8), 2) & "," & _
'                           cQuote & oTempADO("POSID") & cQuote & "," & IIf(aTimeInfo(0) + aTimeInfo(3) + aTimeInfo(5) > 0, oTempADO("POS_ALLOW"), 0) & "," & _
'                           oTempADO("MTD_BASIC") & "," & oTempADO("MTD_GROSS") & "," & oTempADO("MTD_TAXABLE") & "," & oTempADO("SSER1215") & "," & oTempADO("SSPREM1215") & "," & oTempADO("PS1215") & "," & oTempADO("ES1215") & "," & oTempADO("COLA1215") & "," & _
'                           aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & aTimeInfo(5) & "," & aTimeInfo(6) & "," & aTimeInfo(13) & "," & aTimeInfo(14) & "," & aTimeInfo(7) & "," & _
'                           aPayInfo(0) & "," & aPayInfo(1) & "," & aPayInfo(2) & "," & aPayInfo(3) & "," & aPayInfo(4) & "," & aPayInfo(12) & "," & aPayInfo(5) & "," & aPayInfo(6) & "," & aPayInfo(13) & "," & aPayInfo(14) & "," & aPayInfo(7) & "," & _
'                           aPayInfo(9) & "," & aPayInfo(10) + n13mopay & "," & aPayInfo(11) & "," & nIncentive & "," & n13mopay & "," & _
'                           aAdjustment(0) & "," & aAdjustment(1) & "," & _
'                           cQuote & oTempADO("PAGIBIGNO") & cQuote & "," & cQuote & oTempADO("SSNUM") & cQuote & "," & _
'                           cQuote & IIf(Trim(oTempADO("PHEALTHNUM")) = "", oTempADO("SSNUM"), oTempADO("PHEALTHNUM")) & cQuote & "," & cQuote & oTempADO("TIN") & cQuote & "," & _
'                           aTimeInfo(15) & "," & aPayInfo(15) & "," & aPayInfo(11) & ")"


                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ' --> deduction transaction saved here...
                ' --> WAP is excluded in deductions...
                ' --> Emergency is excluded - 20070831
                
                
                'If (oTempADO("PAYSTATUS") <> 2) And (oTempADO("EMP_STAT") <> 0) And Not ((oTempADO("WAP") = 1) And (oTempADO("EMP_STAT") = 1)) Then 'And (nGrossAmt > 0) Then
                
                
                If (oTempADO("PAYSTATUS") <> 2) And (oTempADO("EMP_STAT") <> 0) And Not ((oTempADO("WAP") = 1) And (oTempADO("EMP_STAT") = 1)) Then 'And (nGrossAmt > 0) Then
                    lAllDed = True
                Else
                    lAllDed = False
                    nDedTag = 0
                End If
                
'                If (oTempADO("PAYSTATUS") <> 2) And (oTempADO("EMP_STAT") <> 0) And Not ((oTempADO("WAP") = 1) And (oTempADO("EMP_STAT") = 1)) Then 'And (nGrossAmt > 0) Then
                    cSqlStmt = "select DEDID,DEF_AMT,FIX_DED,DEDTAG, DEDTYPE from pa3330 where " & IIf(nPeriod = 1, "PERIOD1=1", "PERIOD2=1")
                    OpenQueryDNS cSqlStmt, oRSetDed, False
                    If oRSetDed.RecordCount > 0 Then
                        aDedAmt = Array(0#, 0#, 0#, 0#, 0#, 0#, "")
                        nTotDed = 0
                        
                        ' --> added 20061003
                        nTotExempt = 0
                        lWithTax = False
                        
                        While Not oRSetDed.EOF
                            aDedAmt(6) = ""
                            nDedAmt = 0
                            nDedAmt2 = 0
                            nDedAmt3 = 0
                            
                            If (aPayInfo(10) + n13mopay) > 0 Then
                                Select Case oRSetDed("DEDID")
                                    Case "001"      ' --> SSS Premium
                                        If lAllDed Then
                                            cSqlStmt = "select ER_SS, EE_SS, ER_EC  from pa7770 where " & (oTempADO("MTD_GROSS") + aPayInfo(10) + n13mopay) & " between range1 and range2"
                                            OpenQueryDNS cSqlStmt, objdbRs, False
                                            nDedAmt = IIf(objdbRs.RecordCount > 0, objdbRs("EE_SS"), 0) - IIf(nPeriod = 2, oTempADO("SSPREM1215"), 0)
                                            nDedAmt2 = IIf(objdbRs.RecordCount > 0, objdbRs("ER_SS"), 0) - IIf(nPeriod = 2, oTempADO("SSER1215"), 0)
                                            nDedAmt3 = IIf(objdbRs.RecordCount > 0, objdbRs("ER_EC"), 0)
                                            aDedAmt(0) = nDedAmt
                                            aDedAmt(1) = nDedAmt2
                                            aDedAmt(5) = IIf(objdbRs.RecordCount > 0, objdbRs("ER_EC"), 0)
                                            
                                            If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                        End If
                                    Case "003"
                                        ' --> Pag-Ibig Premium, revised 20120502
                                        If lAllDed Then
                                            nDedAmt = (oTempADO("MTD_GROSS") + aPayInfo(10) + n13mopay)
                                            
                                            
                                            ' --> employer's contribution
                                            If nDedAmt > 5000 Then
                                                nDedAmt2 = oRSetDed("def_amt")
                                            Else
                                                nDedAmt2 = nDedAmt * 0.02
                                            End If
                                            
                                            ' --> employee contribution
                                            If (nDedAmt * IIf((oTempADO("MTD_GROSS") + aPayInfo(10) + n13mopay) > 1500, 0.02, 0.01)) >= oRSetDed("def_amt") Then
                                                nDedAmt = oRSetDed("def_amt")
                                            Else
                                                nDedAmt = nDedAmt * IIf((oTempADO("MTD_GROSS") + aPayInfo(10) + n13mopay) > 1500, 0.02, 0.01)
                                            End If
                                            
                                            If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                        End If
                                        
                                        ' --> Pag-Ibig Premium, revised 20070705
''                                        If lAllDed Then
''                                            nDedAmt = (oTempADO("MTD_BASIC") + oTempADO("COLA1215") + aPayInfo(9) + Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2))
''
''                                            ' --> employer's contribution
''                                            If nDedAmt > 5000 Then
''                                                nDedAmt2 = oRSetDed("def_amt")
''                                            Else
''                                                nDedAmt2 = nDedAmt * 0.02
''                                            End If
''
''                                            ' --> employee contribution
''                                            If (nDedAmt * IIf((oTempADO("MTD_BASIC") + oTempADO("COLA1215") + aPayInfo(9) + Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2)) > 1500, 0.02, 0.01)) >= oRSetDed("def_amt") Then
''                                                nDedAmt = oRSetDed("def_amt")
''                                            Else
''                                                nDedAmt = nDedAmt * IIf((oTempADO("MTD_BASIC") + oTempADO("COLA1215") + aPayInfo(9) + Round((oTempADO("COLA_AMT") * (aTimeInfo(3) + aTimeInfo(0))), 2)) > 1500, 0.02, 0.01)
''                                            End If
''
''                                            If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
''                                        End If
                                        
                                    Case "005"      ' --> PhilHealth/Medicare
                                        If lAllDed Then
                                            If aDedAmt(0) > 0 Then
                                                cSqlStmt = "select def_amt from di3673 " & _
                                                           " where (empid=" & cQuote & oTempADO("EMPID") & cQuote & ")" & _
                                                           " and (dedid=" & cQuote & oRSetDed("DEDID") & cQuote & ")" & _
                                                           " and (" & IIf(nPeriod = 1, "period1=1", "period2=1") & ")"
                                                OpenQueryDNS cSqlStmt, oRset1, False
                                                
                                                cSqlStmt = "select PS, ES from PA7454 where " & (oTempADO("MTD_GROSS") + aPayInfo(10)) & " between range1 and range2"
                                                OpenQueryDNS cSqlStmt, objdbRs, False
                                                If oRset1.RecordCount > 0 Then
                                                    nDedAmt = oRset1("def_amt") - IIf(nPeriod = 2, oTempADO("PS1215"), 0)
                                                Else
                                                    nDedAmt = IIf(objdbRs.RecordCount > 0, objdbRs("PS"), 0) - IIf(nPeriod = 2, oTempADO("PS1215"), 0)
                                                End If
                                                nDedAmt2 = IIf(objdbRs.RecordCount > 0, objdbRs("ES"), 0) - IIf(nPeriod = 2, oTempADO("ES1215"), 0)
                                                aDedAmt(3) = nDedAmt
                                                aDedAmt(4) = nDedAmt2
                                            Else
                                                aDedAmt(3) = 0
                                                aDedAmt(4) = 0
                                            End If
                                            If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                        End If
                                    Case "006"      ' --> Withholding Tax
                                        If lAllDed Then
                                            If (oTempADO("emp_stat") = 2) And (Trim(oTempADO("TAXID")) <> "") And oTempADO("Rate_amt") > gBasicRate Then lWithTax = True
                                        End If
                                    Case Else

                                        If lAllDed Then
                                            ' --> custom employee deduction
                                            cSqlStmt = "select def_amt, cut_off_amt, acc_amt, ctrl_no from di3673 " & _
                                                       " where (empid=" & cQuote & oTempADO("EMPID") & cQuote & ")" & _
                                                       " and (dedid=" & cQuote & oRSetDed("DEDID") & cQuote & ")" & _
                                                       " and (status=0)" & _
                                                       " and (period" & nPeriod & "=1)"
    '                                        Script2File cSqlStmt
                                            OpenQueryDNS cSqlStmt, oRset1, False
                                            If oRset1.RecordCount > 0 Then
                                                If oRset1("cut_off_amt") > oRset1("acc_amt") Then
                                                    aDedAmt(6) = oRset1("ctrl_no")
                                                    If (oRset1("acc_amt") + oRset1("def_amt")) > oRset1("cut_off_amt") Then
                                                        nDedAmt = oRset1("cut_off_amt") - oRset1("acc_amt")
                                                    Else
                                                        nDedAmt = oRset1("def_amt")
                                                    End If
                                                    If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                                End If
                                            Else
                                                If oRSetDed("dedtype") <> 1 Then
                                                    If oRSetDed("dedtag") = 1 Then
                                                        If oTempADO("emp_stat") = 2 Then
                                                            If oTempADO("isunion") = 0 Then
                                                                If (gAssessID = oRSetDed("DEDID")) Then
                                                                    ' --> special assessment
                                                                    nDedAmt = IIf(lAssess, nAssessAmt, 0)
                                                                Else
                                                                    ' --> for Union Dues...
                                                                    nDedAmt = oRSetDed("DEF_AMT")
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        ' --> default deduction
                                                        nDedAmt = oRSetDed("DEF_AMT")
                                                    End If
                                                    If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                                End If
                                            End If
                                        Else
                                            If oRSetDed("DEDID") = "002" Or oRSetDed("DEDID") = "004" Or oRSetDed("DEDID") = "018" Or oRSetDed("DEDID") = "020" Then
                                                lTrueTag = False
                                            Else
                                                If gCompanyID = "0003" Then
                                                    If oRSetDed("DEDID") = "007" Or oRSetDed("DEDID") = "013" Then
                                                        lTrueTag = False
                                                    Else
                                                        lTrueTag = True
                                                    End If
                                                End If
                                                
                                                If gCompanyID = "0002" Then
                                                    If oRSetDed("DEDID") = "007" Or oRSetDed("DEDID") = "015" Then
                                                        lTrueTag = False
                                                    Else
                                                        lTrueTag = True
                                                    End If
                                                End If
                                                
                                                If gCompanyID <> "0003" And gCompanyID <> "0002" Then
                                                    lTrueTag = True
                                                End If
                                                
                                                
                                                
                                            End If
                                            
                                            
                                            
                                            If lTrueTag Then
                                                ' --> custom employee deduction
                                                cSqlStmt = "select def_amt, cut_off_amt, acc_amt, ctrl_no from di3673 " & _
                                                           " where (empid=" & cQuote & oTempADO("EMPID") & cQuote & ")" & _
                                                           " and (dedid=" & cQuote & oRSetDed("DEDID") & cQuote & ")" & _
                                                           " and (status=0)" & _
                                                           " and (period" & nPeriod & "=1)"
        '                                        Script2File cSqlStmt
                                                OpenQueryDNS cSqlStmt, oRset1, False
                                                If oRset1.RecordCount > 0 Then
                                                    If oRset1("cut_off_amt") > oRset1("acc_amt") Then
                                                        aDedAmt(6) = oRset1("ctrl_no")
                                                        If (oRset1("acc_amt") + oRset1("def_amt")) > oRset1("cut_off_amt") Then
                                                            nDedAmt = oRset1("cut_off_amt") - oRset1("acc_amt")
                                                        Else
                                                            nDedAmt = oRset1("def_amt")
                                                        End If
                                                        If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                                    End If
                                                Else
                                                    If oRSetDed("dedtype") <> 1 Then
                                                        If oRSetDed("dedtag") = 1 Then
                                                            If oTempADO("emp_stat") = 2 Then
                                                                If oTempADO("isunion") = 0 Then
                                                                    If (gAssessID = oRSetDed("DEDID")) Then
                                                                        ' --> special assessment
                                                                        nDedAmt = IIf(lAssess, nAssessAmt, 0)
                                                                    Else
                                                                        ' --> for Union Dues...
                                                                        nDedAmt = oRSetDed("DEF_AMT")
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            ' --> default deduction
                                                            nDedAmt = oRSetDed("DEF_AMT")
                                                        End If
                                                        If InStr(1, cDedID, oRSetDed("dedid")) > 1 Then nTotExempt = nTotExempt + nDedAmt
                                                    End If
                                                End If
                                                nDedTag = 1
                                            Else
                                                nDedTag = 0
                                            End If
                                        End If
                                End Select
                            End If
                            
                            If lAllDed Then
                                cSqlStmt = "INSERT INTO PA87263(PERIODID, PERIOD_STAT, EMPID, DEDID, CTRL_NO, DED_AMT, DED_AMT2, DED_AMT3, COMPUTED)VALUES(" & _
                                           cQuote & cParam & cQuote & "," & nPeriod & "," & _
                                           cQuote & oTempADO("EMPID") & cQuote & "," & _
                                           cQuote & oRSetDed("DEDID") & cQuote & "," & _
                                           cQuote & aDedAmt(6) & cQuote & "," & _
                                           Round(nDedAmt, 2) & "," & _
                                           Round(nDedAmt2, 2) & "," & _
                                           Round(nDedAmt3, 2) & "," & _
                                           oRSetDed("FIX_DED") & ")"
                                OpenQueryDNS cSqlStmt, objdbRs, True
                            Else
                                If nDedTag = 1 Then
                                    cSqlStmt = "INSERT INTO PA87263(PERIODID, PERIOD_STAT, EMPID, DEDID, CTRL_NO, DED_AMT, DED_AMT2, DED_AMT3, COMPUTED)VALUES(" & _
                                               cQuote & cParam & cQuote & "," & nPeriod & "," & _
                                               cQuote & oTempADO("EMPID") & cQuote & "," & _
                                               cQuote & oRSetDed("DEDID") & cQuote & "," & _
                                               cQuote & aDedAmt(6) & cQuote & "," & _
                                               Round(nDedAmt, 2) & "," & _
                                               Round(nDedAmt2, 2) & "," & _
                                               Round(nDedAmt3, 2) & "," & _
                                               oRSetDed("FIX_DED") & ")"
                                    OpenQueryDNS cSqlStmt, objdbRs, True
                                End If
                            End If
                            Script2File cSqlStmt
                            nTotDed = nTotDed + Round(nDedAmt, 2)
                            
                            oRSetDed.MoveNext
                        Wend
                        
                        ' --> added 20061004
                        If lWithTax Then
                            
                            ' --> revised 20070105
                            If aPeriodInfo(4) > 0 Then
                            
                                ' --> annual withholding tax...
                                aTmpTax = ComputeTax(cParam, oTempADO("empid"), cDedID, Year(aPeriodInfo(1)), aPayInfo(10), nTotExempt)
                                nDedAmt = aTmpTax(0)
                                aDedAmt(2) = aTmpTax(0)
                            Else
                            
                                If oTempADO("emp_stat") = 2 Then
                                
                                    cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
                                               " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("MTD_TAXABLE") + (aPayInfo(10) - nTotExempt)) & ">=ded_amt2)" & _
                                               " order by ded_amt2 desc limit 1"
'                                    Script2File cSqlStmt
                                    OpenQueryDNS cSqlStmt, objdbRs, False
                                    If objdbRs.RecordCount > 0 Then
                                        If objdbRs("DED_PCT") > 0 Then
                                            nDedAmt = objdbRs("DED_AMT") + (((oTempADO("MTD_TAXABLE") + (aPayInfo(10) - nTotExempt - nIncentive)) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
                                        Else
                                            nDedAmt = 0
                                        End If
                                    Else
                                        nDedAmt = 0
                                    End If
                                    aDedAmt(2) = nDedAmt
                                End If
                            
                                If IfExists("pa87263", "(periodid=" & cQuote & cParam & cQuote & ") and (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (dedid='006')") Then
                                    cSqlStmt = "update pa87263 set ded_amt = " & Round(nDedAmt, 2) & _
                                               " where (periodid=" & cQuote & cParam & cQuote & ") and (empid=" & cQuote & oTempADO("empid") & cQuote & ") and (dedid='006')"
                                Else
                                    cSqlStmt = "INSERT INTO PA87263(PERIODID, PERIOD_STAT, EMPID, DEDID, CTRL_NO, DED_AMT, DED_AMT2, DED_AMT3, COMPUTED)VALUES(" & _
                                               cQuote & cParam & cQuote & "," & nPeriod & "," & _
                                               cQuote & oTempADO("EMPID") & cQuote & "," & _
                                               cQuote & "006" & cQuote & "," & _
                                               cQuote & aDedAmt(6) & cQuote & "," & _
                                               Round(nDedAmt, 2) & "," & _
                                               Round(nDedAmt2, 2) & "," & _
                                               Round(nDedAmt3, 2) & "," & _
                                               "0" & ")"
                                End If
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                                
                            End If
                            ' --> end of revision 20070105
                            
                            nTotDed = nTotDed + Round(nDedAmt, 2)
                        End If
                        
                    End If
                    
                    If (nTotDed <> 0) Then
                        cSqlStmt = "UPDATE PA87260 SET DED_AMT=" & Round(nTotDed, 2) & "," & _
                                   " SSPREM=" & Round(aDedAmt(0), 2) & "," & _
                                   " SSER=" & Round(aDedAmt(1), 2) & "," & _
                                   " SSS01=" & Round(aDedAmt(0) + aDedAmt(1), 2) & "," & _
                                   " EC001=" & aDedAmt(5) & "," & _
                                   " MEDICARE=" & Round(aDedAmt(3), 2) & "," & _
                                   " MEDICARE2=" & Round(aDedAmt(4), 2) & "," & _
                                   " MED01=" & Round(aDedAmt(3) + aDedAmt(4), 2) & "," & _
                                   " TAXABLE=" & Round(aPayInfo(10) - nTotExempt - nIncentive, 2) & "," & _
                                   " WTAX=" & Round(aDedAmt(2), 2) & "," & _
                                   " NET_PAY=" & Round(aPayInfo(10) + n13mopay - nTotDed, 2) & _
                                   " WHERE PERIODID=" & cQuote & cParam & cQuote & _
                                   " AND EMPID=" & cQuote & oTempADO("EMPID") & cQuote
                    Else
                        cSqlStmt = "UPDATE PA87260 SET NET_PAY=" & Round(aPayInfo(10) + n13mopay, 2) & "," & _
                                   " DED_AMT=0," & _
                                   " SSPREM=0," & _
                                   " SSER=0," & _
                                   " SSS01=0," & _
                                   " EC001=0," & _
                                   " MEDICARE=0," & _
                                   " MEDICARE2=0," & _
                                   " MED01=0," & _
                                   " TAXABLE=" & Round(aPayInfo(10) + n13mopay - nIncentive, 2) & "," & _
                                   " WTAX=0" & _
                                   " WHERE PERIODID=" & cQuote & cParam & cQuote & _
                                   " AND EMPID=" & cQuote & oTempADO("EMPID") & cQuote
                    End If
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
'                Else
'                    cSqlStmt = "UPDATE PA87260 SET NET_PAY=" & Round(aPayInfo(10) + n13mopay, 2) & "," & _
'                               " DED_AMT=0," & _
'                               " SSPREM=0," & _
'                               " SSER=0," & _
'                               " SSS01=0," & _
'                               " EC001=0," & _
'                               " MEDICARE=0," & _
'                               " MEDICARE2=0," & _
'                               " MED01=0," & _
'                               " TAXABLE=" & Round(aPayInfo(10) + n13mopay - nIncentive, 2) & "," & _
'                               " WTAX=0" & _
'                               " WHERE PERIODID=" & cQuote & cParam & cQuote & _
'                               " AND EMPID=" & cQuote & oTempADO("EMPID") & cQuote
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'                    Script2File cSqlStmt
'                End If
                           
                oTempADO.MoveNext
                
            Wend
            
            ShowProgress 4
            
            
            ' --> for security reason... 20070210
            cSqlStmt = "update pa7730 set isprocess=1, date_process=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                       "where periodid=" & cQuote & cParam & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            Log2Audit Name, "Payroll PeriodID#" & cParam & " successfully processed..."
            
            GoTo ShowData
        End If
        
    Else
        GoTo ShowData
    End If
    
    Exit Sub

ShowData:

    Set oRSetDed = Nothing
    Set oRset1 = Nothing

    GetUserRights PadStr(frmMain.mnuTransaction.Name, " ", 100, PadRight), gUserID
    frmTransaction.lblPeriod.Caption = cParam
    frmTransaction.lblDuration.Caption = ListView1.SelectedItem.SubItems(1) & " Payroll"
    frmTransaction.Show
End Sub


'Sub non_work_hol(ByVal cEmpid As String, aPeriodInfo As Variant)
'    Dim cSqlStmt As String, _
'        oRecordSet As New ADODB.Recordset, _
'        dDate_hol As String, _
'        nWith_Pay As Integer
'
'        cSqlStmt = "select * from PA4329 " & _
'                   "where (date between " & cQuote & Format(aPeriodInfo(0), "yyyy-mm-dd") & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & ") " & _
'                   "  or ((month(date)=" & Month(aPeriodInfo(0)) & ") and (day(date) between " & Day(aPeriodInfo(0)) & " and " & Day(aPeriodInfo(1)) & "))" & _
'                   " and withpay = 1 "
'
'    '    Script2File cSqlStmt
'        OpenQueryDNS cSqlStmt, objdbRs, False
'        If objdbRs.RecordCount > 0 Then
'            dDate_hol = Format(objdbRs("date"), "yyyy-mm-dd")
'            nWith_Pay = objdbRs("withpay")
'        End If
'        If nWith_Pay <> 0 Then
'            If dDate_hol <> "" Then
'                OpenQueryDNS " select * from di36770 where empid = " & cQuote & cEmpid & cQuote & _
'                             "  and date = " & cQuote & Format(dDate_hol, "yyyy-mm-dd") & cQuote & " and reg_hr <> 0", oRecordSet, False
'                If oRecordSet.RecordCount > 0 Then
'                    cSqlStmt = "update di36770 set sun_hr=reg_hr, sun_ot_hr=reg_ot_hr+ sa_reg_ot" & _
'                               " where empid = " & cQuote & cEmpid & cQuote & _
'                               "  and date = " & cQuote & Format(dDate_hol, "yyyy-mm-dd") & cQuote & " and reg_hr <> 0 "
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'
'                    cSqlStmt = "update di36770 set reg_hr=0, reg_ot_hr=0, sa_reg_ot=0" & _
'                               " where empid = " & cQuote & cEmpid & cQuote & _
'                               "  and date = " & cQuote & Format(dDate_hol, "yyyy-mm-dd") & cQuote & " and reg_hr <> 0 "
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'                End If
'            End If
'        End If
'End Sub

' + -->
' |     Procedure Name  :   GenEmpList(ByVal cParam As String, nmode As Integer)
' |     Description     :   Utility to generate Employee Masterlist
' |     Date Created    :   xx mar 2006
' + -->z
Sub CreateTmpEmp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE tmpEmpLst(" & _
               "[LINEID] char(3),           [LINENAME] char(100)," & _
               "[EMPID] char(6),            [FULLNAME] char(100)," & _
               "[TCID] char(6),             [ADDRESS] char(100)," & _
               "[date_hire] date,           [date_end] date," & _
               "[birthday] date,            [status] integer," & _
               "[ssnum] char(15),           [pagibigno] char(15),       [tin] char(15)," & _
               "[rate_amt] double,          [cola_amt] double," & _
               "[pos_allow] double,         [emp_stat] char (100)," & _
               "[paystatus] integer,        [active] integer," & _
               "[sex] char(10),             [position] char(100)," & _
               "[taxcode] char(10),         [shiftname] char(100))"
               
              
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpEmpLst", oTempADO, True
End Sub

Sub GenEmpList(ByVal cPeriod As String, cParam As String, nMode As Integer)
    Dim cSqlStmt As String, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        oRSet3 As New ADODB.Recordset, _
        oRSet4 As New ADODB.Recordset, _
        aOtherInfo As Variant, _
        cString As String
        
    aOtherInfo = Array("", "", "", "")
    
    CreateTmpEmp
    
    ShowProgress 0
    
    
    cString = " and a.active " & IIf(nMode <> 0, "> 0", "=0")
    
    If Trim(cParam) <> "" Then
        cParam = " and a.depid IN " & cParam
    End If
    
    cSqlStmt = " SELECT a.date_hire,b.birthday,b.date_fin,a.date_res,a.active,a.empid,b.tcid, b.firstname, b.mname, b.lastname, concat(b.lastname,', ',b.firstname, ', ',b.mname) as fullname, concat(b.ADD_NO, '', b.ADD_BRGY, '', b.ADD_CITY) as ADDRESS, b.sex, b.posid, a.depid, " & _
               " a.date_hire,b.birthday,b.date_fin,a.date_res, b.status, a.rate_amt, a.cola_amt, a.pos_allow, b.isunion, a.emp_stat, a.paystatus, a.active,  " & _
               " b.taxid, d.taxcode, b.shiftid,b.ssnum , b.pagibigno, b.tin "

    
    OpenQueryDNS "select pclose from pa7730 where periodid =" & cQuote & cPeriod & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If objdbRs("pclose") <> 1 Then
            
            cSqlStmt = cSqlStmt & " FROM pa87260 a left join di3670 b on a.empid=b.empid left join pa8290 d on b.taxid=d.taxid "
        
        Else
            cSqlStmt = cSqlStmt & " FROM pah87260 a left join di3670 b on a.empid=b.empid left join pa8290 d on b.taxid=d.taxid"

        End If
    End If
        
    cSqlStmt = cSqlStmt & " where a.periodid=" & cQuote & cPeriod & cQuote & cString & cParam
'    Script2File cSqlStmt
'    MsgBox cSqlStmt
'            cSqlStmt = "select a.date_hire, a.birthday, a.date_fin, a.date_res, a.status, a.rate_amt, a.cola_amt, a.pos_allow, a.isunion, a.emp_stat, a.paystatus, a.active, " & _
'                       " a.empid, a.firstname, a.mname, a.lastname, concat(a.lastname,', ',a.firstname) as fullname, a.sex, a.posid, a.depid, a.taxid, a.shiftid," & _
'                       " a.ssnum, a.pagibigno, a.tin from di3670 a"
               
'    cSqlStmt = cSqlStmt & " where a.active=" & nmode & IIf(Trim(cParam) = "", "", " and " & cParam)
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        OpenQueryDNS "select posid, posname from DI7670 order by posid", oRset1, False
        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRSet2, False
        OpenQueryDNS "select taxid, taxcode, taxname from pa8290 order by taxid", oRSet3, False
        OpenQueryDNS "select shiftid, `description` from pa74380 order by shiftid", oRSet4, False
        
        While Not oTempADO.EOF
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , , "Processing data for " & oTempADO("fullname")
            
            If oRset1.RecordCount > 0 Then
                oRset1.Requery adAsyncFetch
                oRset1.Find "posid='" & oTempADO("posid") & "'"
                aOtherInfo(0) = IIf(oRset1.EOF, "", oRset1("posname"))
            End If
            
            If oRSet2.RecordCount > 0 Then
                oRSet2.Requery adAsyncFetch
                oRSet2.Find "lineid='" & oTempADO("depid") & "'"
                aOtherInfo(1) = IIf(oRSet2.EOF, "", oRSet2("linename"))
            End If
            
            If oRSet3.RecordCount > 0 Then
                oRSet3.Requery adAsyncFetch
                oRSet3.Find "taxid='" & oTempADO("taxid") & "'"
                aOtherInfo(2) = IIf(oRSet3.EOF, "", oRSet3("taxcode"))
            End If
            
            If oRSet4.RecordCount > 0 Then
                oRSet4.Requery adAsyncFetch
                oRSet4.Find "shiftid='" & oTempADO("shiftid") & "'"
                aOtherInfo(3) = IIf(oRSet4.EOF, "", oRSet4("description"))
            End If
                       
                       
                       cSqlStmt = "insert into tmpEmpLst(empid,tcid,fullname,address,lineid,linename,date_hire,date_end,birthday,[status]," & _
                       " rate_amt,cola_amt,pos_allow,emp_stat,paystatus,[active],[sex],[position],taxcode,shiftname," & _
                       " ssnum, pagibigno, [tin])values(" & _
                       cQuote & oTempADO("empid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("tcid"))) & cQuote & "," & _
                       cQuote & oTempADO("fullname") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("address"))) & cQuote & "," & _
                       cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & "," & _
                       cQuote & Format(oTempADO("date_hire"), "mm/dd/yyyy") & cQuote & "," & cQuote & Format(IIf(oTempADO("active") = 0, Now, IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin"))), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & Format(oTempADO("birthday"), "mm/dd/yyyy") & cQuote & "," & oTempADO("status") & "," & _
                       oTempADO("rate_amt") & "," & oTempADO("cola_amt") & "," & oTempADO("pos_allow") & "," & _
                       cQuote & IIf(oTempADO("emp_stat") = 0, "Wap", IIf(oTempADO("emp_stat") = 1, "Contractual", "Regular")) & cQuote & "," & oTempADO("paystatus") & "," & oTempADO("active") & "," & _
                       cQuote & IIf(oTempADO("sex") = 0, "Male", "Female") & cQuote & "," & _
                       cQuote & DecodeStr(EncodeStr2(aOtherInfo(0))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(2))) & cQuote & "," & _
                       cQuote & DecodeStr(EncodeStr2(aOtherInfo(3))) & cQuote & "," & _
                       cQuote & DecodeStr(EncodeStr2(oTempADO("ssnum"))) & cQuote & "," & _
                       cQuote & DecodeStr(EncodeStr2(oTempADO("pagibigno"))) & cQuote & "," & _
                       cQuote & DecodeStr(EncodeStr2(oTempADO("tin"))) & cQuote & ")"
                       
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oTempADO.MoveNext
        Wend
    End If
    
    ShowProgress 3
    
    QueryTemp "select * from tmpEmpLst", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        GenerateReport IIf(nMode = 0, "", "Resigned/Finished") & " Employee Report Listing", "lst3670.rpt"
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    ShowProgress 4
    
EndGenEmp:
    Set oRset1 = Nothing
    Set oRSet2 = Nothing
    Set oRSet3 = Nothing
    Set oRSet4 = Nothing

    Exit Sub
    
ErrGenEmp:
    ErrorMsg Err.Number, Err.Description, "Employee Report Listing", Name
    
    Resume EndGenEmp
End Sub

Sub CreateTmpEmpSum()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
'    cSqlStmt = "CREATE TABLE tmpEmpLstsum(" & _
'               "[Con] char(3),             [Con_SM] double, " & _
'               "[Con_SF] double, " & _
'               "[Wap_UH] char(3),          [Wap_UH_SM] double, " & _
'               "[Wap_UH_SF] double, " & _
'               "[Wap_H] char(3),           [Wap_H_SM] double, " & _
'               "[Wap_H_SF] double)"
              
    cSqlStmt = " CREATE TABLE tmpEmpLstsum(" & _
               " [Name] char(100),           [Name_SM] double, " & _
               " [Name_SF] double,           [emp_stat] double, " & _
               " [Wap] double)"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpEmpLstsum", oTempADO, True
End Sub

Sub GenEmpListSum(ByVal cPeriod As String, cParam As String, nMode As Integer)
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        aOtherInfo As Variant, _
        cString As String, _
        cString2 As String, _
        nCtr As Integer, _
        cPclose As String
        
    aOtherInfo = Array("", 0#, 0#, 0#)
    
    ShowProgress 0

    CreateTmpEmpSum
    
    OpenQueryDNS "SELECT emp_stat,wap FROM di3670 group by emp_stat,wap ", oRSet, False
    If oRSet.RecordCount > 0 Then
    
        While Not oRSet.EOF
            ShowProgress 2, (oRSet.AbsolutePosition / oRSet.RecordCount) * 100, , , "Processing data... "
            If (oRSet("emp_stat") <> 0) Or (oRSet("wap") <> 1) Then
                If (oRSet("emp_stat") = 0) And (oRSet("wap") = 0) Then
                    aOtherInfo(0) = "WAP"
                    aOtherInfo(1) = oRSet("emp_stat")
                    aOtherInfo(2) = oRSet("wap")
                ElseIf (oRSet("emp_stat") = 2) And (oRSet("wap") = 0) Then
                    aOtherInfo(0) = "Regular"
                    aOtherInfo(1) = oRSet("emp_stat")
                    aOtherInfo(2) = oRSet("wap")
                ElseIf (oRSet("emp_stat") = 1) And (oRSet("wap") = 0) Then
                    aOtherInfo(0) = "Contractual Unhide"
                    aOtherInfo(1) = oRSet("emp_stat")
                    aOtherInfo(2) = oRSet("wap")
                ElseIf (oRSet("emp_stat") = 1) And (oRSet("wap") = 1) Then
                    aOtherInfo(0) = "Contractual Hide"
                    aOtherInfo(1) = oRSet("emp_stat")
                    aOtherInfo(2) = oRSet("wap")
                End If
                cSqlStmt = "insert into tmpEmpLstsum(Name,Name_SM,Name_SF,emp_stat,Wap)values(" & _
                           cQuote & DecodeStr(EncodeStr2(aOtherInfo(0))) & cQuote & ",0,0," & _
                           aOtherInfo(1) & "," & aOtherInfo(2) & ")"
                           
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True

            End If
            
            
            oRSet.MoveNext
        Wend
        ShowProgress 4
    End If
    
    ShowProgress 0
    
    cString = " and a.active " & IIf(nMode <> 0, "> 0", "=0")

    If Trim(cParam) <> "" Then
        cParam = " and a.depid IN " & cParam
    End If


    OpenQueryDNS "select pclose from pa7730 where periodid =" & cQuote & cPeriod & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cPclose = IIf(objdbRs("pclose") <> 0, "pah87260", "pa87260")
    Else
        cPclose = ""
    End If

    cSqlStmt = " SELECT a.emp_stat,b.sex,a.wap FROM " & cPclose & " a " & _
               " left join di3670 b on a.empid=b.empid" & _
               " Where (a.periodid = " & cQuote & cPeriod & cQuote & cString & ") " & cParam & _
               " order by a.emp_stat,a.wap "
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Processing data... "
            
            If oRecordSet("sex") = 0 Then
                cString2 = " set Name_SM =  Name_SM + 1 "
            Else
                cString2 = " set Name_SF = Name_SF + 1 "
            End If
            
            cSqlStmt = " update tmpEmpLstsum " & cString2 & _
                       " where emp_stat = " & oRecordSet("emp_stat") & IIf(oRecordSet("emp_stat") = 0, "", " and wap = " & oRecordSet("wap"))
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
    Else
        ShowProgress 3
        MsgBox "No report to generate!", vbCritical, "System Advisory"

        ShowProgress 4
        GoTo EndGenEmpSum
    End If
    
    ShowProgress 3

    GenerateReport IIf(nMode = 0, "", "Resigned/Finished") & " Employee Summary Listing Report ", "lst36703.rpt"

    ShowProgress 4

    
EndGenEmpSum:
    Set oRecordSet = Nothing

    Exit Sub
    
ErrGenEmpSum:
    ErrorMsg Err.Number, Err.Description, "Employee Summary Report Listing", Name
    
    Resume EndGenEmpSum
End Sub


' + -->
' |     Procedure Name  :   GenPayRoll(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
' |     Description     :   Print Utility for Payroll module
' |     Date Created    :   16 mar 2006
' + -->
'       4   Payslip - 3 reports
'               Regular Payslip
'               SA Payslip
'               WAP Payslip
'               WAP SA Payslip
'       5   Acknowledgement (all 3) - 3 reports
'       6   Payroll Sheet (summary/detail) - 8 reports
'               Regular Payroll Sheet
'               Regular Payroll (deduction) Sheet
'               SA Payroll Sheet
'               WAP Payroll Sheet
'               WAP SA Payroll Sheet
Sub CreateTmpPaySlip(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cParam As String
    
    If nMode = 0 Then
        cSqlStmt = " CREATE TABLE tmp7297655([BACCNTNO] CHAR(16), [PERIODNAME] CHAR(100),     [SEQ_NO] INTEGER, " & _
                   " [P_DAY] DOUBLE,            [P_HOLIDAY] DOUBLE, " & _
                   " [DEPID] CHAR(3),           [DEPTNAME] CHAR(100),       [DEPID2] CHAR(3),           [DEPTNAME2] CHAR(100), " & _
                   " [EMPID] CHAR(6),           [ACTIVE] INTEGER,           [EMP_STAT] INTEGER," & _
                   " [FULLNAME] CHAR(100),      [FNAME] CHAR(50),           [LNAME] CHAR(50),           [MNAME] CHAR(50), " & _
                   " [RATE_AMT] DOUBLE,         [COLA_AMT] DOUBLE,          [SUN_COLA] DOUBLE,          [POS_ALLOW] DOUBLE, " & _
                   " [REG_DAY] DOUBLE,          [REG_PAY] DOUBLE,           [REG_OT_HR] DOUBLE,         [REG_OT_PAY] DOUBLE, " & _
                   " [NDIFF_DAY] DOUBLE,        [NDIFF_PAY] DOUBLE,         [NDIFF_OT_HR] DOUBLE,       [NDIFF_OT_PAY] DOUBLE, " & _
                   " [HOLIDAY] DOUBLE,          [HOL_PAY] DOUBLE, " & _
                   " [SA_REG_OT] DOUBLE,        [SA_REG_PAY] DOUBLE,        [SA_NDIFF_OT] DOUBLE,       [SA_NDIFF_PAY] DOUBLE, " & _
                   " [SUN_HR] DOUBLE,           [SUN_PAY] DOUBLE,           [SUN_OT] DOUBLE,            [SUN_OT_PAY] DOUBLE, " & _
                   " [SUN_ND] DOUBLE,           [SUN_ND_PAY] DOUBLE,        [SUN_ND_OT] DOUBLE,         [SUN_ND_OT_PAY] DOUBLE, " & _
                   " [ADJ_PAY] DOUBLE,          [SA_ADJ_PAY] DOUBLE, " & _
                   " [OTHER_PAY] DOUBLE,        [LEAVE_PAY] DOUBLE, " & _
                   " [DED_AMT] DOUBLE,          [GROSS_PAY] DOUBLE, " & _
                   " [NET_PAY] DOUBLE,          [SA_NET_PAY] DOUBLE, " & _
                   " [SIGNATORY1] char(50),     [POSNAME1] char(50)," & _
                   " [SIGNATORY2] char(50),     [POSNAME2] char(50)," & _
                   " [SIGNATORY3] char(50),     [POSNAME3] char(50)," & _
                   " [SIGNATORY4] char(50),     [POSNAME4] char(50)," & _
                   " [SIGNATORY5] char(50),     [POSNAME5] char(50)," & _
                   " [SIGNATORY6] char(50),     [POSNAME6] char(50)," & _
                   " [SIGNATORY7] char(50),     [POSNAME7] char(50)," & _
                   " [13MO_PAY] DOUBLE,         [CMPID] char(4), [REG_OT_RATE] DOUBLE )"
    ElseIf nMode = 1 Then
        cSqlStmt = " CREATE TABLE tmp7297655d(" & _
                   " [BACCNTNO] CHAR(16), [PERIODID] CHAR(5),        [EMPID] CHAR(6)," & _
                   " [DEDID] CHAR(3),           [DEDNAME] CHAR(100)," & _
                   " [AMOUNT] DOUBLE,           [SEQ_NO] INTEGER)"
    Else
        OpenQueryDNS "select dedid from PA3330", objdbRs, False
        If objdbRs.RecordCount > 0 Then
            While Not objdbRs.EOF
                cParam = cParam & "[dedname" & objdbRs("dedid") & "] char(50), [dedamt" & objdbRs("dedid") & "] DOUBLE,"
                objdbRs.MoveNext
            Wend
        End If
        
        cSqlStmt = " CREATE TABLE tmp7297655d2( [PERIODNAME] CHAR(100), [SEQ_NO] INTEGER,  " & _
                   " [P_DAY] DOUBLE,            [P_HOLIDAY] DOUBLE,         " & _
                   " [DEPID] CHAR(3),           [DEPTNAME] CHAR(100),   " & _
                   " [DEPID2] CHAR(3),          [DEPTNAME2] CHAR(100), " & _
                   " [EMPID] CHAR(6),           [ACTIVE] INTEGER,       [EMP_STAT] INTEGER," & _
                   " [FULLNAME] CHAR(100),      [FNAME] CHAR(50),       [LNAME] CHAR(50), " & _
                   cParam & _
                   " [DED_AMT] DOUBLE,          [GROSS_PAY] DOUBLE,     [NET_PAY] DOUBLE)"
    End If
              
    'MsgBox cSqlStmt
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM " & IIf(nMode = 0, "tmp7297655", IIf(nMode = 1, "tmp7297655d", "tmp7297655d2")), oTempADO, True
End Sub

Sub CreateTmpIncPaySlip(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cParam As String
    
    If nMode = 0 Then
        cSqlStmt = " CREATE TABLE tmp7294620(   [PERIODNAME] CHAR(100),     [SEQ_NO] INTEGER, " & _
                   " [DEPID] CHAR(3),           [DEPTNAME] CHAR(100),       [DEPID2] CHAR(3),           [DEPTNAME2] CHAR(100), " & _
                   " [EMPID] CHAR(6),           [ACTIVE] INTEGER,           [EMP_STAT] INTEGER," & _
                   " [FULLNAME] CHAR(100),      [FNAME] CHAR(50),           [LNAME] CHAR(50),           [MNAME] CHAR(50), " & _
                   " [RATE_AMT] DOUBLE,         [INC_HR] DOUBLE,            [INC_PAY] DOUBLE, " & _
                   " [GROSS_PAY] DOUBLE,        [NET_PAY] DOUBLE," & _
                   " [SIGNATORY1] char(50),     [POSNAME1] char(50)," & _
                   " [SIGNATORY2] char(50),     [POSNAME2] char(50)," & _
                   " [SIGNATORY3] char(50),     [POSNAME3] char(50)," & _
                   " [SIGNATORY4] char(50),     [POSNAME4] char(50)," & _
                   " [SIGNATORY5] char(50),     [POSNAME5] char(50)," & _
                   " [SIGNATORY6] char(50),     [POSNAME6] char(50)," & _
                   " [SIGNATORY7] char(50),     [POSNAME7] char(50)," & _
                   " [13MO_PAY] DOUBLE,         [CMPID] char(4), [REG_OT_RATE] DOUBLE )"
    End If
              
    'MsgBox cSqlStmt
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
QueryTemp "DELETE FROM " & IIf(nMode = 0, "tmp7294620", "tmp7294620d"), oTempADO, True
End Sub


Sub GenIncPayroll(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
    Dim cSqlStmt, _
        cPeriodName, _
        cDedParam, _
        cDedValue As String, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        oRSet3 As New ADODB.Recordset, _
        nActive, _
        nCtr As Integer, _
        aOtherInfo As Variant, _
        aPosInfo As Variant, _
        aDedName As Variant, aDedAmt As Variant, _
        aMonthName As Variant

    If SSTab1.TabVisible(3) Then
        If Not ChkPersonnel(Text6) Then Exit Sub
        If Not ChkPersonnel(Text5) Then Exit Sub
        If Not ChkPersonnel(Text1) Then Exit Sub
        If Not ChkPersonnel(Text7) Then Exit Sub
        If Not ChkPersonnel(Text3) Then Exit Sub
        If Not ChkPersonnel(Text8) Then Exit Sub
    End If

    ' --> process active employee first here...
    nActive = IIf(Tag = 15, 1, IIf(Tag = 29, 1, 0))
    
    aPosInfo = Array("", "", "", "", "", "", "")
    
'    aMonthName = Array("Enero", "Pebrero", "Marso", "Abril", "Mayo", "Hunyo", "Hulyo", "Agosto", "Setyembre", "Oktubre", "Nobyembre", "Disyembre")
    
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aPosInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aPosInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text1.Text & "'"
        aPosInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aPosInfo(3) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text3.Text & "'"
        aPosInfo(4) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text8.Text & "'"
        aPosInfo(5) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If

    If Trim(cParam) <> "" Then
        cParam = "a.depid IN " & cParam
    End If
    
    CreateTmpIncPaySlip 0  ' --> header
    
loopd2:

    aOtherInfo = Array("", "", "", "")

'    OpenQueryDNS "select * from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        If Tag = 20 Then
'            cPeriodName = "13th Month Pay " & Year(objdbRs("date_start"))
'        Else
'            If Check5.Value = vbChecked Then
'                If Combo1.ListIndex <> 2 Then
'                    cPeriodName = "Para sa sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & Format(objdbRs("date_start"), " d, yyyy") & " hanggang " & aMonthName(Month(objdbRs("date_end")) - 1) & Format(objdbRs("date_end"), " d, yyyy")
'                Else
'                    cPeriodName = "Sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & " " & Day(objdbRs("date_start")) & "-" & Day(objdbRs("date_end")) & ", " & Year(objdbRs("date_end"))
'                End If
'            Else
'                cPeriodName = "For the " & IIf(Tag = 5, "Payroll", "") & " period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
'            End If
'        End If
'    End If

    OpenQueryDNS "select * from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cPeriodName = "For the " & IIf(Tag = 5, "Payroll", "") & " period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
    End If
    
    If (Check4.Value = vbChecked) Then
        cSqlStmt = "select a.periodid, a.seq_no, a.empid, a.firstname, a.lastname, a.mname, a.emp_stat, " & _
               " a.fullname, a.depid, a.rate_amt, a.inc_hr, a.inc_pay, " & _
               " a.gross_pay, a.net_pay, a.sa_net_pay, a.active" & _
               " from pa87260 a"

    Else
        cSqlStmt = "select a.periodid, a.depid, count(a.empid) as manpower," & _
                   " sum(a.inc_hr) as inc_hr, sum(a.inc_pay) as inc_pay, " & _
                   " sum(a.gross_pay) as gross_pay, sum(a.net_pay) as net_pay " & _
                   " from pa87260 a"
    End If
    
    cSqlStmt = cSqlStmt & " where (a.active" & IIf(nActive = 0, "=0", "<>0") & ")" & IIf(Trim(cParam) = "", "", " and (" & cParam & ")") & _
               IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
               " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
               " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & IIf((Check4.Value <> vbChecked), " group by a.depid", "")
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRset1, False
        
        ShowProgress 0
        
        While Not oTempADO.EOF
        
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            If oRset1.RecordCount > 0 Then
                oRset1.Requery adAsyncFetch
                oRset1.Find "lineid='" & oTempADO("depid") & "'"
                aOtherInfo(1) = IIf(oRset1.EOF, "", oRset1("linename"))
            End If
            
            If oTempADO("inc_Pay") <> 0 Then
            
            
                If Check4.Value <> vbChecked Then
                    cSqlStmt = "insert into tmp7294620(periodname, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & _
                               " rate_amt, inc_hr, inc_pay, gross_pay, net_pay, " & _
                               " signatory1,signatory2,signatory3,signatory4,signatory5,signatory6,signatory7, " & _
                               " posname1,posname2,posname3,posname4,posname5,posname6,posname7)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               oTempADO("manpower") & "," & _
                               oTempADO("inc_hr") & "," & oTempADO("inc_pay") & "," & oTempADO("gross_pay") & "," & _
                               oTempADO("net_pay") & "," & _
                               cQuote & EncodeStr2(Label8.Caption) & cQuote & "," & cQuote & EncodeStr2(Label6.Caption) & cQuote & "," & cQuote & EncodeStr2(Label4.Caption) & cQuote & "," & _
                               cQuote & EncodeStr2(Label15.Caption) & cQuote & "," & cQuote & EncodeStr2(Label14.Caption) & cQuote & "," & cQuote & EncodeStr2(Label16.Caption) & cQuote & "," & cQuote & cQuote & "," & _
                               cQuote & aPosInfo(0) & cQuote & "," & cQuote & aPosInfo(1) & cQuote & "," & cQuote & aPosInfo(2) & cQuote & "," & cQuote & aPosInfo(3) & cQuote & "," & cQuote & aPosInfo(4) & cQuote & "," & cQuote & aPosInfo(5) & cQuote & "," & cQuote & aPosInfo(6) & cQuote & ")"
                Else
                    cSqlStmt = "insert into tmp7294620(periodname, seq_no, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & " empid, emp_stat, [active], fullname, fname, mname, lname, " & _
                               " rate_amt, inc_hr, inc_pay, gross_pay, net_pay)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("seq_no") & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
                               cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("mname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
                               oTempADO("rate_amt") & "," & oTempADO("inc_hr") & "," & oTempADO("inc_pay") & "," & _
                               oTempADO("gross_pay") & "," & oTempADO("net_pay") & ")"
                End If
            
    '            MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
            End If
            oTempADO.MoveNext
        Wend
        ShowProgress 4
    End If
    
    ' --> process resigned/fc employee next...
    If nActive = 0 Then
        nActive = 1
        GoTo loopd2
    End If
    
    ShowProgress 0
    
    QueryTemp "select * from " & IIf(Check3.Value <> vbChecked, "tmp7294620", "tmp72946202"), objdbRs, False
    If objdbRs.RecordCount > 0 Then

        ShowProgress 3

        Select Case Tag
            Case 27 ', 19      ' --> Payslip/13th month
                Select Case nFilter
                    Case 0
                        GenerateReport "INCENTIVE PAYSLIP", "rpt7294620.rpt"
                    Case 30 '3
                        GenerateReport IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(Check5.Value = vbChecked, "rpt727547T.rpt", "rpt727547.rpt")
                    Case 2, 4
                        GenerateReport IIf(nFilter = 4, "EMERGENCY", "WAP") & " INCENTIVE PAYSLIP", "rpt9274620.rpt"
                End Select
'
            Case 30 '  22      ' --> Acknowledgement
                GenerateReport IIf(nFilter = 2, "WAP ", IIf(nFilter = 3, "WAP SA ", IIf(Tag = 22, "13th Month Pay 2006 ", IIf(nFilter = 4, "EMERGENCY ", "")))) & "INCENTIVE Acknowledgement Report", "rpt462734748.rpt"
'
            Case 28, 29      ' --> Payroll Sheet (summary/detail)
                If Check3.Value <> vbChecked Then
                    Select Case nFilter
                        Case 0      ' --> reg payroll sheet
                            GenerateReport "INCENTIVE PAYROLL ", IIf(Check4.Value = vbChecked, "rpt462748d.rpt", "rpt462748s.rpt")
                        Case 2, 4   ' --> wap payroll sheet
                            GenerateReport IIf(nFilter = 4, "EMERGENCY", "WAP") & " INCENTIVE PAYROLL", IIf(Check4.Value = vbChecked, "rpt927748462d.rpt", "rpt927748462s.rpt")
                    End Select
                End If

'            Case 20     ' --> 13th month payroll sheet
'                GenerateReport cPeriodName, IIf(Check4.Value = vbChecked, "rpt13667.rpt", "rpt13667s.rpt")

        End Select
        
    Else
        ShowProgress 3

        MsgBox "No report to generate!", vbCritical, "System Advisory"
    End If
    
    ShowProgress 4
    
    Set oRset1 = Nothing
    Set oRSet2 = Nothing
    Set oRSet3 = Nothing
  
End Sub

Sub GenPayRoll(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
    Dim cSqlStmt, _
        cPeriodName, _
        cDedParam, _
        cDedParam2, _
        cDedValue As String, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        oRSet3 As New ADODB.Recordset, _
        nActive, _
        nCtr As Integer, _
        aOtherInfo As Variant, _
        aPosInfo As Variant, _
        aDedName As Variant, aDedAmt As Variant, _
        aMonthName As Variant
        
    Dim nCloseTag As Integer
        

    If SSTab1.TabVisible(3) Then
        If Not ChkPersonnel(Text6) Then Exit Sub
        If Not ChkPersonnel(Text5) Then Exit Sub
        If Not ChkPersonnel(Text1) Then Exit Sub
        If Not ChkPersonnel(Text7) Then Exit Sub
        If Not ChkPersonnel(Text3) Then Exit Sub
        If Not ChkPersonnel(Text8) Then Exit Sub
    End If

    ' --> process active employee first here...
    nActive = IIf((Tag = 15) Or (Tag = 44) Or (Tag = 45), 1, 0)
    
    aPosInfo = Array("", "", "", "", "", "", "")
    
    aMonthName = Array("Enero", "Pebrero", "Marso", "Abril", "Mayo", "Hunyo", "Hulyo", "Agosto", "Setyembre", "Oktubre", "Nobyembre", "Disyembre")
    
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aPosInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aPosInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text1.Text & "'"
        aPosInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aPosInfo(3) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text3.Text & "'"
        aPosInfo(4) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text8.Text & "'"
        aPosInfo(5) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If

    If Trim(cParam) <> "" Then
        cParam = "a.depid IN " & cParam
    End If
    
    If Check3.Value = vbChecked Then
        CreateTmpPaySlip 2  ' --> payroll sheet deduction
    Else
        CreateTmpPaySlip 0  ' --> header
        If (nFilter = 0) Or (nFilter = 2) Or (nFilter = 4) Then CreateTmpPaySlip 1 ' --> detail
    End If
    
loopd2:

    aOtherInfo = Array("", "", "", "")
    
    If Check3.Value = vbChecked Then
        If Combo1.ListIndex = 2 Or Combo1.ListIndex = 4 Then
            If gCompanyID = "0003" Then
                cDedParam2 = " a.dedid not in(" & cQuote & "001" & cQuote & "," & _
                             cQuote & "002" & cQuote & "," & _
                             cQuote & "003" & cQuote & "," & _
                             cQuote & "004" & cQuote & "," & _
                             cQuote & "005" & cQuote & "," & _
                             cQuote & "006" & cQuote & "," & _
                             cQuote & "007" & cQuote & "," & _
                             cQuote & "013" & cQuote & "," & _
                             cQuote & "020" & cQuote & "," & _
                             cQuote & "021" & cQuote & ")"
            Else
                If gCompanyID = "0002" Then
                    cDedParam2 = " a.dedid not in(" & cQuote & "001" & cQuote & "," & _
                                 cQuote & "002" & cQuote & "," & _
                                 cQuote & "003" & cQuote & "," & _
                                 cQuote & "004" & cQuote & "," & _
                                 cQuote & "005" & cQuote & "," & _
                                 cQuote & "006" & cQuote & "," & _
                                 cQuote & "007" & cQuote & "," & _
                                 cQuote & "015" & cQuote & "," & _
                                 cQuote & "018" & cQuote & "," & _
                                 cQuote & "020" & cQuote & "," & _
                                cQuote & "021" & cQuote & ")"
                Else
                    cDedParam2 = ""
                End If
            End If
        Else
            cDedParam2 = ""
        End If
        
        OpenQueryDNS "select dedid, dedname,dederpid from PA3330 a " & IIf(cDedParam2 = "", "", " where " & cDedParam2), objdbRs, False

        If objdbRs.RecordCount > 0 Then
            cDedParam = ""
            ReDim aDedName(objdbRs.RecordCount + 100)
            ReDim aDedAmt(objdbRs.RecordCount + 100)
            While Not objdbRs.EOF
'                MsgBox (objdbRs("dedname") & " (" & IIf(objdbRs("dederpid") = "", objdbRs("dedid"), objdbRs("dederpid")) & ")")
                cDedParam = cDedParam & "dedname" & objdbRs("dedid") & "," & "dedamt" & objdbRs("dedid") & ","
'                MsgBox aDedName(Val(objdbRs("dedid")))
                aDedName(Val(objdbRs("dedid"))) = (objdbRs("dedname") & " (" & IIf(objdbRs("dederpid") = "", objdbRs("dedid"), objdbRs("dederpid")) & ")")

                objdbRs.MoveNext

            Wend
        End If
        
        
    End If

    OpenQueryDNS "select * from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If Tag = 20 Then
            cPeriodName = "13th Month Pay " & Year(objdbRs("date_start"))
        Else
            If Check5.Value = vbChecked Then
                If gCompanyID = "0002" Then
                    cPeriodName = "For the " & IIf(Tag = 5, "Payroll", "") & " period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
                Else
                    If Combo1.ListIndex <> 2 Then
                        cPeriodName = "Para sa sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & Format(objdbRs("date_start"), " d, yyyy") & " hanggang " & aMonthName(Month(objdbRs("date_end")) - 1) & Format(objdbRs("date_end"), " d, yyyy")
                    Else
                        cPeriodName = "Sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & " " & Day(objdbRs("date_start")) & "-" & Day(objdbRs("date_end")) & ", " & Year(objdbRs("date_end"))
                    End If
                End If
            Else
                cPeriodName = "For the " & IIf(Tag = 5, "Payroll", "") & " period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
            End If
        End If
    
        '---> Add new 2015-05-07
        
        If objdbRs("PCLOSE") = 0 Then
            nCloseTag = 0
        Else
            nCloseTag = 1
        End If
        
    
    End If
    cSqlStmt = ""
    
    If (Check4.Value = vbChecked) Then
'        cSqlStmt = "select a.periodid, a.p_day, a.p_holiday, a.seq_no, a.empid, a.firstname, a.lastname, a.mname, a.emp_stat, " & _
'                   " a.fullname, a.depid, a.rate_amt, a.cola_amt, a.cola, a.sun_cola, a.pos_allow, a.reg_day, a.reg_pay, " & _
'                   " a.reg_ot_hr, a.reg_ot_pay,a.ndiff_day, a.ndiff_pay, a.ndiff_ot_hr, a.ndiff_ot_pay, a.holiday, a.hol_pay, " & _
'                   " a.sa_reg_ot, a.sa_reg_pay, a.sa_ndiff_ot, a.sa_ndiff_pay, a.sun_hr, a.sun_pay, a.sun_ot, a.sun_ot_pay, a.sun_nd, a.sun_nd_pay, a.sun_nd_ot, a.sun_nd_ot_pay, " & _
'                   " a.adj_pay, a.sa_adj_pay, a.other_pay, a.leave_pay, a.m13pay, a.ded_amt, a.gross_pay, a.net_pay, a.sa_net_pay, a.active" & _
'                   " from pa87260 a"
                       
        ' --> Revision 2015-05-07
        cSqlStmt = "select a.periodid, a.p_day, a.p_holiday, a.seq_no, a.empid, a.firstname, a.lastname, a.mname, a.emp_stat, " & _
                   " a.fullname, a.depid, a.rate_amt, a.cola_amt, a.cola, a.sun_cola, a.pos_allow, a.reg_day, a.reg_pay, " & _
                   " a.reg_ot_hr, a.reg_ot_pay,a.ndiff_day, a.ndiff_pay, a.ndiff_ot_hr, a.ndiff_ot_pay, a.holiday, a.hol_pay, " & _
                   " a.sa_reg_ot, a.sa_reg_pay, a.sa_ndiff_ot, a.sa_ndiff_pay, a.sun_hr, a.sun_pay, a.sun_ot, a.sun_ot_pay, a.sun_nd, a.sun_nd_pay, a.sun_nd_ot, a.sun_nd_ot_pay, " & _
                   " a.adj_pay, a.sa_adj_pay, a.other_pay, a.leave_pay, a.m13pay, a.ded_amt, a.gross_pay, a.net_pay, a.sa_net_pay, a.active" & _
                   " from " & IIf(nCloseTag = 0, "pa87260", "pah87260") & " a"
                   
    Else
'        cSqlStmt = "select a.periodid, a.p_day, a.p_holiday, a.depid, count(a.empid) as manpower,sum(truncate(a.cola_amt*(a.reg_day+a.ndiff_day),2)) as cola_amt, " & _
'                   " sum(a.cola) as cola, sum(a.pos_allow) as pos_allow, sum(a.sun_cola) as sun_cola, " & _
'                   " sum(a.reg_day) as reg_day, sum(a.reg_pay) as reg_pay, sum(a.reg_ot_hr) as reg_ot_hr, sum(a.reg_ot_pay) as reg_ot_pay," & _
'                   " sum(a.ndiff_day) as ndiff_day, sum(a.ndiff_pay) as ndiff_pay, sum(a.ndiff_ot_hr) as ndiff_ot_hr, sum(a.ndiff_ot_pay) as ndiff_ot_pay, " & _
'                   " sum(a.holiday) as holiday, sum(a.hol_pay) as hol_pay, sum(a.sa_reg_ot) as sa_reg_ot, sum(a.sa_reg_pay) as sa_reg_pay, " & _
'                   " sum(a.sa_ndiff_ot) as sa_ndiff_ot, sum(a.sa_ndiff_pay) as sa_ndiff_pay, sum(a.sun_hr) as sun_hr, sum(a.sun_pay) as sun_pay, sum(a.sun_nd) as sun_nd, sum(a.sun_nd_pay) as sun_nd_pay, sum(a.sun_nd_ot) as sun_nd_ot, sum(a.sun_nd_ot_pay) as sun_nd_ot_pay, " & _
'                   " sum(a.sun_ot) as sun_ot, sum(a.sun_ot_pay) as sun_ot_pay, sum(a.adj_pay) as adj_pay, sum(a.sa_adj_pay) as sa_adj_pay, " & _
'                   " sum(a.other_pay) as other_pay, sum(a.leave_pay) as leave_pay, sum(a.m13pay) as m13pay, sum(a.ded_amt) as ded_amt, " & _
'                   " sum(a.gross_pay) as gross_pay, sum(a.net_pay) as net_pay, sum(a.sa_net_pay) as sa_net_pay" & _
'                   " from pa87260 a"
                   
        ' --> Revision 2015-05-07
        cSqlStmt = "select a.periodid, a.p_day, a.p_holiday, a.depid, count(a.empid) as manpower,sum(truncate(a.cola_amt*(a.reg_day+a.ndiff_day),2)) as cola_amt, " & _
                   " sum(a.cola) as cola, sum(a.pos_allow) as pos_allow, sum(a.sun_cola) as sun_cola, " & _
                   " sum(a.reg_day) as reg_day, sum(a.reg_pay) as reg_pay, sum(a.reg_ot_hr) as reg_ot_hr, sum(a.reg_ot_pay) as reg_ot_pay," & _
                   " sum(a.ndiff_day) as ndiff_day, sum(a.ndiff_pay) as ndiff_pay, sum(a.ndiff_ot_hr) as ndiff_ot_hr, sum(a.ndiff_ot_pay) as ndiff_ot_pay, " & _
                   " sum(a.holiday) as holiday, sum(a.hol_pay) as hol_pay, sum(a.sa_reg_ot) as sa_reg_ot, sum(a.sa_reg_pay) as sa_reg_pay, " & _
                   " sum(a.sa_ndiff_ot) as sa_ndiff_ot, sum(a.sa_ndiff_pay) as sa_ndiff_pay, sum(a.sun_hr) as sun_hr, sum(a.sun_pay) as sun_pay, sum(a.sun_nd) as sun_nd, sum(a.sun_nd_pay) as sun_nd_pay, sum(a.sun_nd_ot) as sun_nd_ot, sum(a.sun_nd_ot_pay) as sun_nd_ot_pay, " & _
                   " sum(a.sun_ot) as sun_ot, sum(a.sun_ot_pay) as sun_ot_pay, sum(a.adj_pay) as adj_pay, sum(a.sa_adj_pay) as sa_adj_pay, " & _
                   " sum(a.other_pay) as other_pay, sum(a.leave_pay) as leave_pay, sum(a.m13pay) as m13pay, sum(a.ded_amt) as ded_amt, " & _
                   " sum(a.gross_pay) as gross_pay, sum(a.net_pay) as net_pay, sum(a.sa_net_pay) as sa_net_pay" & _
                  " from " & IIf(nCloseTag = 0, "pa87260", "pah87260") & " a"
                  
    End If
    ' for NO ATM

    If (Tag = 19) Or (Tag = 20) Or (Tag = 22) Then
        cSqlStmt = cSqlStmt & " where a.periodid=" & cQuote & cPeriodID & cQuote & " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                   IIf((Check4.Value <> vbChecked), " group by a.depid", "")
    
    Else
        cSqlStmt = cSqlStmt & " where (a.active" & IIf(nActive = 0, "=0", "<>0") & ")" & IIf(Trim(cParam) = "", "", " and (" & cParam & ")") & _
                   IIf(nFilter = 0, " and (a.emp_stat <> 0)", IIf(nFilter = 1, " and (a.sa_net_pay<>0) and (a.emp_stat<>0)", IIf(nFilter = 2, " and (a.emp_stat=0)", IIf(nFilter = 4, "", " and (a.sa_net_pay<>0) and (a.emp_stat=0)")))) & _
                   " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
                   " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                   IIf(Combo1.ListIndex = 2, " and a.emp_stat=0", "") & _
                   " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                   IIf((Check4.Value <> vbChecked), " group by a.depid", "")
    End If
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRset1, False
        
        ShowProgress 0
        
        While Not oTempADO.EOF
        
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
            If oRset1.RecordCount > 0 Then
                oRset1.Requery adAsyncFetch
                oRset1.Find "lineid='" & oTempADO("depid") & "'"
                aOtherInfo(1) = IIf(oRset1.EOF, "", oRset1("linename"))
            End If
            
            If Check3.Value = vbChecked Then
                ' --> Payroll Sheet (deduction)
                For nCtr = 1 To UBound(aDedAmt) - 1
                    aDedAmt(nCtr) = 0
                Next nCtr
                cDedValue = ""

                If Check4.Value = vbChecked Then

'                    cSqlStmt = "select a.periodid, a.empid, a.dedid, a.ded_amt," & _
'                               " ifnull(b.dedname,'') as dedname " & _
'                               " from pa87263 a left join pa3330 b on a.dedid=b.dedid" & _
'                               " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
'                               " and (a.empid=" & cQuote & oTempADO("empid") & cQuote & ")" & _
'                               IIf(cDedParam2 <> "", " and (" & cDedParam2 & ")", "")
                
                    ' --> revision 2015-05-07
                    cSqlStmt = "select a.periodid, a.empid, a.dedid, a.ded_amt," & _
                               " ifnull(b.dedname,'') as dedname " & _
                               " from " & IIf(nCloseTag = 0, "pa87263", "pah87263") & " a " & _
                               " left join pa3330 b on a.dedid=b.dedid" & _
                               " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                               " and (a.empid=" & cQuote & oTempADO("empid") & cQuote & ")" & _
                               IIf(cDedParam2 <> "", " and (" & cDedParam2 & ")", "")
                
                
                Else

'                    cSqlStmt = "select a.periodid, c.depid, a.dedid, sum(a.ded_amt) as ded_amt," & _
'                               "  ifnull(b.dedname,'') as dedname " & _
'                               "from pa87263 a left join pa87260 c on a.empid=c.empid and a.periodid=c.periodid " & _
'                               "  left join di3670 d on a.empid=d.empid" & _
'                               "  left join pa3330 b on a.dedid=b.dedid" & _
'                               " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
'                               " and (c.depid=" & cQuote & oTempADO("depid") & cQuote & ")" & _
'                               IIf((nFilter = 0), " and (d.emp_stat <> 0)", "") & _
'                               " and (d.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'                               " and (c.active" & IIf(nActive = 0, "=0", "<>0") & ")" & _
'                               IIf(Combo1.ListIndex = 2, " and c.emp_stat=0", "") & _
'                               IIf(cDedParam2 <> "", " and (" & cDedParam2 & ")", "") & _
'                               IIf(nFilter = 4, "", IIf(nActive = 0, " and (c.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")", "")) & _
'                               " group by b.dedid, c.depid"
                
'                    cSqlStmt = "select a.periodid, c.depid, a.dedid, sum(a.ded_amt) as ded_amt," & _
'                               "  ifnull(b.dedname,'') as dedname " & _
'                               "from pa87263 a left join pa87260 c on a.empid=c.empid and a.periodid=c.periodid " & _
'                               "  left join di3670 d on a.empid=d.empid" & _
'                               "  left join pa3330 b on a.dedid=b.dedid" & _
'                               " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
'                               " and (c.depid=" & cQuote & oTempADO("depid") & cQuote & ")" & _
'                               IIf((nFilter = 0), " and (d.emp_stat <> 0)", "") & _
'                               " and (a.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
'                               " and (c.active" & IIf(nActive = 0, "=0", "<>0") & ")" & _
'                               IIf(Combo1.ListIndex = 2, " and c.emp_stat=0", "") & _
'                               IIf(cDedParam2 <> "", " and (" & cDedParam2 & ")", "") & _
'                               " and (c.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
'                               " group by b.dedid, c.depid"
                               
'                   ---> newrevision 2015-05-07
                    cSqlStmt = "select a.periodid, c.depid, a.dedid, sum(a.ded_amt) as ded_amt," & _
                               " ifnull(b.dedname,'') as dedname " & _
                               " from " & IIf(nCloseTag = 0, "pa87263", "pah87263") & " a " & _
                               " left join " & IIf(nCloseTag = 0, "pa87260", "pah87260") & " c on a.empid=c.empid and a.periodid=c.periodid " & _
                               " left join pa3330 b on a.dedid=b.dedid" & _
                               " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                               " and (c.depid=" & cQuote & oTempADO("depid") & cQuote & ")" & _
                               IIf((nFilter = 0), " and (c.emp_stat <> 0)", "") & _
                               " and (c.paystatus " & IIf(nFilter = 4, "=", "<>") & " 2)" & _
                               " and (c.active" & IIf(nActive = 0, "=0", "<>0") & ")" & _
                               IIf(Combo1.ListIndex = 2, " and c.emp_stat=0", "") & _
                               IIf(cDedParam2 <> "", " and (" & cDedParam2 & ")", "") & _
                               " and (c.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                               " group by b.dedid, c.depid"
                
                End If
                
'                MsgBox cSqlStmt
'                Script2File cSqlStmt
                OpenQueryDNS cSqlStmt, oRSet2, False
                If oRSet2.RecordCount > 0 Then
                    While Not oRSet2.EOF
                        aDedAmt(Val(oRSet2("dedid"))) = oRSet2("ded_amt")
                        oRSet2.MoveNext
                    Wend
                    For nCtr = 1 To UBound(aDedAmt) - 1
                        If DecodeStr(EncodeStr2(aDedName(nCtr))) <> "" Then
                            cDedValue = cDedValue & cQuote & DecodeStr(EncodeStr2(aDedName(nCtr))) & cQuote & "," & Val(aDedAmt(nCtr)) & ","
                        End If
                    Next nCtr
                Else
                    For nCtr = 1 To UBound(aDedAmt) - 1
                        If DecodeStr(EncodeStr2(aDedName(nCtr))) <> "" Then
                            cDedValue = cDedValue & cQuote & DecodeStr(EncodeStr2(aDedName(nCtr))) & cQuote & ",0,"
                        End If
                    Next nCtr
                End If
                
                If Check4.Value = vbChecked Then
                    cSqlStmt = "insert into tmp7297655d2(periodname, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & _
                               " seq_no,empid, emp_stat, [active], fullname, fname, lname, " & _
                               cDedParam & "ded_amt, gross_pay, net_pay)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               oTempADO("seq_no") & "," & cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
                               cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
                               cDedValue & oTempADO("ded_amt") & "," & oTempADO("gross_pay") & "," & oTempADO("net_pay") & ")"
                Else
                    cSqlStmt = "insert into tmp7297655d2(periodname, p_day, p_holiday, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & _
                               cDedParam & "ded_amt, gross_pay, net_pay)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("p_day") & "," & oTempADO("p_holiday") & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               cDedValue & oTempADO("ded_amt") & "," & oTempADO("gross_pay") & "," & oTempADO("net_pay") & ")"
                End If
                Script2File cSqlStmt
                
                QueryTemp cSqlStmt, objdbRs, True
                
            Else
            
                If Check4.Value <> vbChecked Then
                    cSqlStmt = "insert into tmp7297655(periodname, p_day, p_holiday, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & _
                               " rate_amt, cola_amt, sun_cola, pos_allow, reg_day, reg_pay, reg_ot_hr, reg_ot_pay, " & _
                               " ndiff_day, ndiff_pay, ndiff_ot_hr, ndiff_ot_pay, [holiday], hol_pay, sa_reg_ot, sa_reg_pay, " & _
                               " sa_ndiff_ot, sa_ndiff_pay, sun_hr, sun_pay, sun_ot, sun_ot_pay, sun_nd, sun_nd_pay, sun_nd_ot, sun_nd_ot_pay, " & _
                               " adj_pay, sa_adj_pay, other_pay, leave_pay, 13mo_pay, ded_amt, gross_pay, net_pay, sa_net_pay, " & _
                               " signatory1,signatory2,signatory3,signatory4,signatory5,signatory6,signatory7, " & _
                               " posname1,posname2,posname3,posname4,posname5,posname6,posname7)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("p_day") & "," & oTempADO("p_holiday") & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               oTempADO("manpower") & "," & oTempADO("cola") & "," & oTempADO("sun_cola") & "," & oTempADO("pos_allow") & "," & _
                               oTempADO("reg_day") & "," & oTempADO("reg_pay") & "," & oTempADO("reg_ot_hr") & "," & oTempADO("reg_ot_pay") & "," & _
                               oTempADO("ndiff_day") & "," & oTempADO("ndiff_pay") & "," & oTempADO("ndiff_ot_hr") & "," & oTempADO("ndiff_ot_pay") & "," & _
                               oTempADO("holiday") & "," & oTempADO("hol_pay") & "," & _
                               oTempADO("sa_reg_ot") & "," & oTempADO("sa_reg_pay") & "," & _
                               oTempADO("sa_ndiff_ot") & "," & oTempADO("sa_ndiff_pay") & "," & _
                               oTempADO("sun_hr") & "," & oTempADO("sun_pay") & "," & oTempADO("sun_ot") & "," & oTempADO("sun_ot_pay") & "," & _
                               oTempADO("sun_nd") & "," & oTempADO("sun_nd_pay") & "," & oTempADO("sun_nd_ot") & "," & oTempADO("sun_nd_ot_pay") & "," & _
                               oTempADO("adj_pay") & "," & oTempADO("sa_adj_pay") & "," & _
                               oTempADO("other_pay") & "," & oTempADO("leave_pay") & "," & oTempADO("M13PAY") & "," & oTempADO("ded_amt") & "," & oTempADO("gross_pay") & "," & _
                               IIf((Tag = 5) And ((nFilter = 1) Or (nFilter = 3)), oTempADO("sa_net_pay"), oTempADO("net_pay")) & "," & oTempADO("sa_net_pay") & "," & _
                               cQuote & EncodeStr2(Label8.Caption) & cQuote & "," & cQuote & EncodeStr2(Label6.Caption) & cQuote & "," & cQuote & EncodeStr2(Label4.Caption) & cQuote & "," & _
                               cQuote & EncodeStr2(Label15.Caption) & cQuote & "," & cQuote & EncodeStr2(Label14.Caption) & cQuote & "," & cQuote & EncodeStr2(Label16.Caption) & cQuote & "," & cQuote & cQuote & "," & _
                               cQuote & aPosInfo(0) & cQuote & "," & cQuote & aPosInfo(1) & cQuote & "," & cQuote & aPosInfo(2) & cQuote & "," & cQuote & aPosInfo(3) & cQuote & "," & cQuote & aPosInfo(4) & cQuote & "," & cQuote & aPosInfo(5) & cQuote & "," & cQuote & aPosInfo(6) & cQuote & ")"
                Else
                ' --> REG_OT_RATE add by the auditor of mico... 20080305
'                    MsgBox "d2 sya dumaan"
                
                    cSqlStmt = "insert into tmp7297655(periodname, p_day, p_holiday, seq_no, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & " empid, emp_stat, [active], fullname, fname, mname, lname, " & _
                               " rate_amt, cola_amt, sun_cola, pos_allow, reg_day, reg_pay, reg_ot_hr, reg_ot_pay, " & _
                               " ndiff_day, ndiff_pay, ndiff_ot_hr, ndiff_ot_pay, [holiday], hol_pay, sa_reg_ot, sa_reg_pay, " & _
                               " sa_ndiff_ot, sa_ndiff_pay, sun_hr, sun_pay, sun_ot, sun_ot_pay, sun_nd, sun_nd_pay, sun_nd_ot, sun_nd_ot_pay, " & _
                               " adj_pay, sa_adj_pay, other_pay, leave_pay, 13mo_pay, ded_amt, gross_pay, net_pay, sa_net_pay,REG_OT_RATE)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("p_day") & "," & oTempADO("p_holiday") & "," & oTempADO("seq_no") & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
                               cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("mname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
                               oTempADO("rate_amt") & "," & oTempADO("cola") & "," & oTempADO("sun_cola") & "," & oTempADO("pos_allow") & "," & _
                               oTempADO("reg_day") & "," & oTempADO("reg_pay") & "," & _
                               oTempADO("reg_ot_hr") & "," & oTempADO("reg_ot_pay") & "," & _
                               oTempADO("ndiff_day") & "," & oTempADO("ndiff_pay") & "," & _
                               oTempADO("ndiff_ot_hr") & "," & oTempADO("ndiff_ot_pay") & "," & _
                               oTempADO("holiday") & "," & oTempADO("hol_pay") & "," & _
                               oTempADO("sa_reg_ot") & "," & oTempADO("sa_reg_pay") & "," & _
                               oTempADO("sa_ndiff_ot") & "," & oTempADO("sa_ndiff_pay") & "," & _
                               oTempADO("sun_hr") & "," & oTempADO("sun_pay") & "," & oTempADO("sun_ot") & "," & oTempADO("sun_ot_pay") & "," & _
                               oTempADO("sun_nd") & "," & oTempADO("sun_nd_pay") & "," & oTempADO("sun_nd_ot") & "," & oTempADO("sun_nd_ot_pay") & "," & _
                               oTempADO("adj_pay") & "," & oTempADO("sa_adj_pay") & "," & _
                               oTempADO("other_pay") & "," & oTempADO("leave_pay") & "," & oTempADO("M13PAY") & "," & oTempADO("ded_amt") & "," & oTempADO("gross_pay") & "," & _
                               IIf(((Tag = 5) Or (Tag = 45)) And ((nFilter = 1) Or (nFilter = 3)), oTempADO("sa_net_pay"), oTempADO("net_pay")) & "," & oTempADO("sa_net_pay") & "," & _
                               oTempADO("rate_amt") / 8 * 1.25 & ")"

                End If
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
                
                If ((nFilter = 0) Or (nFilter = 2) Or (nFilter = 4)) And (Tag <> 20) Then
                    If Check4.Value = vbChecked Then
                        cSqlStmt = "select a.periodid, a.empid, a.dedid, a.ded_amt," & _
                                   " ifnull(b.dedname,'') as dedname," & _
                                   " ifnull(b.dedname2,'') as dedname2" & _
                                   " from " & IIf(nCloseTag = 0, "pa87263", "pah87263") & " a " & _
                                   " left join pa3330 b on a.dedid=b.dedid" & _
                                   " where (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                                   " and (a.empid=" & cQuote & oTempADO("empid") & cQuote & ")"
                        OpenQueryDNS cSqlStmt, oRSet2, False
                        If oRSet2.RecordCount > 0 Then
                            While Not oRSet2.EOF
                                cSqlStmt = "insert into tmp7297655d(periodid, empid, dedid, dedname, amount)values(" & _
                                           cQuote & cPeriodID & cQuote & "," & _
                                           cQuote & oTempADO("empid") & cQuote & "," & _
                                           cQuote & oRSet2("dedid") & cQuote & "," & _
                                           cQuote & DecodeStr(EncodeStr2(oRSet2(IIf(Check5.Value = vbChecked, "dedname2", "dedname")))) & cQuote & "," & _
                                           oRSet2("ded_amt") & ")"
                                QueryTemp cSqlStmt, objdbRs, True
                                oRSet2.MoveNext
                            Wend
                        End If
                    End If
                End If
            End If
            
            oTempADO.MoveNext
        Wend
        
        ShowProgress 4
    End If
    
    ShowProgress 0
    
    QueryTemp "select * from " & IIf(Check3.Value <> vbChecked, "tmp7297655", "tmp7297655d2"), objdbRs, False
    If objdbRs.RecordCount > 0 Then
    
        ShowProgress 3
        
         Select Case Tag
            Case 4, 19, 44     ' --> Payslip/13th month
                Select Case nFilter
                    Case 0
                        GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check5.Value <> vbChecked, IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & "PAYSLIP", IIf(Check6.Value = vbChecked, "WALANG ATM ", "") & "TALAAN NG KINITA"), IIf(Check5.Value = vbChecked, "rpt7297547T.rpt", "rpt7297547.rpt")
                    Case 1, 3
                        If gCompanyID <> "0002" Then
                            If (gCompanyID = "0001") Or (gCompanyID = "0006") Then
                                'd2 sa yun SA nila na katulad ng sa mico....
                                GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(Check5.Value = vbChecked, "rpt727547T.rpt", "rpt727547.rpt")
                            Else
                                GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(nFilter = 3, IIf(Check5.Value = vbChecked, "rpt727547TK1.rpt", "rpt727547K1.rpt"), IIf(Check5.Value = vbChecked, IIf(lAudit = 0, "rpt727547TK1.rpt", "rpt727547T.rpt"), IIf(lAudit = 0, "rpt727547K1.rpt", "rpt727547.rpt")))
                            End If
                        Else
                            GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(nFilter = 3, IIf(Check5.Value = vbChecked, "rpt727547TK1.rpt", "rpt727547K1.rpt"), IIf(Check5.Value = vbChecked, IIf(lAudit = 0, "rpt727547TK1.rpt", "rpt727547T.rpt"), IIf(lAudit = 0, "rpt727547K1.rpt", "rpt727547.rpt")))
                        End If
                    Case 2
                        GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check5.Value <> vbChecked, IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & "WAP Payslip", IIf(Check6.Value = vbChecked, "WALANG ATM ", "") & "TALAAN NG KINITA PARA SA WAP"), IIf(Check5.Value = vbChecked, "rpt9277547t.rpt", "rpt9277547.rpt")
                    Case 4
                        GenerateReport IIf(Tag = 44, "Resigned/Finished ", "") & IIf(Check5.Value <> vbChecked, IIf(Check6.Value = vbChecked, "NO ATM ", "") & "EMERGENCY Payslip", "TALAAN NG KINITA PARA SA EMERGENCY"), IIf(Check5.Value = vbChecked, "rpt9277547t.rpt", "rpt9277547_EM.rpt")

                End Select

            Case 5, 22, 45     ' --> Acknowledgement
                GenerateReport IIf(Tag = 45, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & IIf(nFilter = 1, "SA ", IIf(nFilter = 2, "WAP ", IIf(nFilter = 3, "WAP SA ", IIf(Tag = 22, "13th Month Pay ", IIf(nFilter = 4, "EMERGENCY ", ""))))) & "Acknowledgement Report", "rpt734748.rpt"

            Case 6, 15     ' --> Payroll Sheet (summary/detail)
                If Check3.Value <> vbChecked Then
                    Select Case nFilter
                        Case 0      ' --> reg payroll sheet
                            'GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, IIf(gCompanyID <> "0003", IIf(Tag = 15, "", "NO ATM "), "NO ATM "), "WITH ATM ") & "PAYROLL", IIf(Check4.Value = vbChecked, "rpt729748d.rpt", "rpt729748s.rpt")
                            GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & "PAYROLL", IIf(Check4.Value = vbChecked, "rpt729748d.rpt", "rpt729748s.rpt")
                        Case 1, 3  ' --> sa payroll sheet
                            'GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, IIf(gCompanyID <> "0003", IIf(Tag = 15, "", "NO ATM "), "NO ATM "), "WITH ATM ") & IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(Check4.Value = vbChecked, "rpt72748d.rpt", "rpt72748s.rpt")
                            GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & IIf(nFilter = 1, "SA", "WAP SA") & " PAYROLL", IIf(Check4.Value = vbChecked, "rpt72748d.rpt", "rpt72748s.rpt")
                        '20096-08-17
                        Case 2   ' --> wap payroll sheet
                            'GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, IIf(gCompanyID <> "0003", IIf(Tag = 15, "", "NO ATM "), "NO ATM "), "WITH ATM ") & "WAP PAYROLL", IIf(Check4.Value = vbChecked, "rpt927748d.rpt", "rpt927748s.rpt")
                            GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & "WAP PAYROLL", IIf(Check4.Value = vbChecked, "rpt927748d.rpt", "rpt927748s.rpt")
                        Case 4   ' --> emergency payroll sheet
                            'GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, IIf(gCompanyID <> "0003", IIf(Tag = 15, "", "NO ATM "), "NO ATM "), "WITH ATM ") & "EMERGENCY PAYROLL", IIf(Check4.Value = vbChecked, "rpt927748d_EM.rpt", "rpt927748s_EM.rpt")
                            GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & "EMERGENCY PAYROLL", IIf(Check4.Value = vbChecked, "rpt927748d_EM.rpt", "rpt927748s_EM.rpt")

                    End Select
                Else

                    ' --> Payroll Deduction sheet
                    GenerateReport IIf(Tag = 15, "Resigned/Finished ", "") & cPeriodName, IIf(Check4.Value = vbChecked, "rpt729748ded.rpt", "rpt729748sd.rpt")
                End If

            Case 20     ' --> 13th month payroll sheets
                If gCompanyID <> "0002" Then
                    GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & cPeriodName, IIf(Check4.Value = vbChecked, "rpt13667.rpt", "rpt13667s.rpt")
                Else
                    If Check3.Value <> vbChecked Then
                        GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & cPeriodName, IIf(Check4.Value = vbChecked, "rpt13667_k1.rpt", "rpt13667s_k1.rpt")
                    Else
                        GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "WITH ATM ") & cPeriodName, IIf(Check4.Value = vbChecked, "rpt729748ded_k1.rpt", "rpt729748sd_k1.rpt")
                    End If
                End If
        End Select
        
    Else
        ShowProgress 3

        MsgBox "No report to generate!", vbCritical, "System Advisory"
    End If
    
    ShowProgress 4
    
    Set oRset1 = Nothing
    Set oRSet2 = Nothing
    Set oRSet3 = Nothing
End Sub



' + -->
' |     Procedure Name  :   GenDenom(ByVal cPeriodID As String)
' |     Description     :   Generate Denomination Report
' |     Date Created    :   20 mar 2006
' + -->
Sub CreateTmpDenom()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String

    cSqlStmt = " CREATE TABLE TmpDenom(     [PERIODNAME] CHAR(100), " & _
               " [DEPID] CHAR(3),           [DEPTNAME] CHAR(100), " & _
               " [EMPID] CHAR(6),           [ACTIVE] INTEGER,       [EMP_STAT] INTEGER," & _
               " [FULLNAME] CHAR(100),      [FNAME] CHAR(100),      [LNAME] CHAR(100), " & _
               " [FROM_NAME] CHAR(100),     [ATTN_NAME] CHAR(100), " & _
               " [P1000] DOUBLE,            [P500] DOUBLE,          [P100] DOUBLE, " & _
               " [P50] DOUBLE,              [P20] DOUBLE, " & _
               " [P10] DOUBLE,              [P5] DOUBLE, " & _
               " [P1] DOUBLE,               [PCOIN] DOUBLE, " & _
               " [NET_PAY] DOUBLE,          [SA_NET_PAY] DOUBLE, " & _
               " [CMPID] char(4))"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM TmpDenom", oTempADO, True
End Sub

Function GetDenom(nAmount As Double) As Variant
    Dim nRemainder As Double, _
        nCtr As Integer, aOtherInfo As Variant, _
        cString As String
        
    aOtherInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    ' (0)   -   1000
    ' (1)   -   500
    ' (2)   -   100
    ' (3)   -   50
    ' (4)   -   20
    ' (5)   -   10
    ' (6)   -   5
    ' (7)   -   1
    ' (8)   -   cent
    
    aOtherInfo(0) = Int(nAmount) \ 1000
    nRemainder = nAmount - (aOtherInfo(0) * 1000)
    
    If nRemainder > 0 Then
        aOtherInfo(1) = Int(nRemainder) \ 500
        nRemainder = Val(nRemainder) - (aOtherInfo(1) * 500)
    End If
    
    If nRemainder > 0 Then
        aOtherInfo(2) = Int(nRemainder) \ 100
        nRemainder = Val(nRemainder) - (aOtherInfo(2) * 100)
    End If
    
    If nRemainder > 0 Then
        aOtherInfo(3) = Int(nRemainder) \ 50
        nRemainder = Val(nRemainder) - (aOtherInfo(3) * 50)
    End If

    If nRemainder > 0 Then
        aOtherInfo(4) = Int(nRemainder) \ 20
        nRemainder = Val(nRemainder) - (aOtherInfo(4) * 20)
    End If

    If nRemainder > 0 Then
        aOtherInfo(5) = Int(nRemainder) \ 10
        nRemainder = Val(nRemainder) - (aOtherInfo(5) * 10)
    End If

    If nRemainder > 0 Then
        aOtherInfo(6) = Int(nRemainder) \ 5
        nRemainder = Val(nRemainder) - (aOtherInfo(6) * 5)
    End If

    If nRemainder > 0 Then
        aOtherInfo(7) = Int(nRemainder) \ 1
        nRemainder = Val(nRemainder) - (aOtherInfo(7) * 1)
    End If
    
    aOtherInfo(8) = nRemainder
    
    GetDenom = aOtherInfo
End Function

Sub GenDenom(ByVal cPeriodID As String, nFilter As Integer, nMode As Integer)
    Dim cSqlStmt As String, cPeriodName As String, _
        aDenomInfo As Variant, aDenomInfo2 As Variant, aDenomInfo3 As Variant
        
    If Not ChkPersonnel(Text6) Then Exit Sub
    If Not ChkPersonnel(Text5) Then Exit Sub
    
    CreateTmpDenom
    
    aDenomInfo = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aDenomInfo2 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aDenomInfo3 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    
    
    If nFilter <> 2 Then
    'incentive revise 20080319
        
        If nMode <> 41 Then
            
            If nMode <> 23 Then
                cSqlStmt = "select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
                           "  a.depid, ifnull(b.linename,'') as department, " & _
                           "  a.net_pay , a.sa_net_pay,a.inc_pay, a.active, a.emp_stat " & _
                           "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
                           "where a.periodid=" & cQuote & cPeriodID & cQuote & _
                           " and paystatus=0 and a.active = 0 " & _
                           " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")"
            Else
                cSqlStmt = "select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
                           "  a.depid, ifnull(b.linename,'') as department, " & _
                           "  a.net_pay , a.sa_net_pay,a.inc_pay, a.active, a.emp_stat " & _
                           "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
                           "where a.periodid=" & cQuote & cPeriodID & cQuote & _
                           " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")"
            End If
        Else

'            cSqlStmt = "select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
'                       "  a.depid, ifnull(b.linename,'') as department, " & _
'                       "  a.net_pay , a.sa_net_pay,a.inc_pay, a.active, a.emp_stat " & _
'                       "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
'                       "where a.periodid=" & cQuote & cPeriodID & cQuote & _
'                       IIf(Combo1.ListIndex = 0, " and paystatus=2 and a.active=0", " and a.active<>0 ")
            
'            If gCompanyID = "0003" Then
'                cSqlStmt = "select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
'                           "  a.depid, ifnull(b.linename,'') as department, " & _
'                           "  a.net_pay , a.sa_net_pay,a.inc_pay, a.active, a.emp_stat " & _
'                           "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
'                           "where a.periodid=" & cQuote & cPeriodID & cQuote & _
'                           IIf(Combo1.ListIndex = 0, " and paystatus=2 and a.active=0", " and a.active<>0 ")
'
'            Else
                cSqlStmt = "select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
                           "  a.depid, ifnull(b.linename,'') as department, " & _
                           "  a.net_pay , a.sa_net_pay,a.inc_pay, a.active, a.emp_stat " & _
                           "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
                           "where a.periodid=" & cQuote & cPeriodID & cQuote & _
                           IIf(Combo1.ListIndex = 0, " and paystatus=2 and a.active=0", " and a.active<>0 ") & _
                           " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")"
'            End If
        End If
    Else
                  
        If gCompanyID <> "0001" Then
            cSqlStmt = " select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
                       " a.depid, ifnull(b.linename,'') as department, " & _
                       " ((c.sl_avail+c.vl_avail)-(c.sl_use+c.vl_use))* c.rate_amt as net_pay," & _
                       " 0 as sa_net_pay, 0 as inc_pay, a.active, a.emp_stat " & _
                       " from pa87260 a " & _
                       " left join di5463 b on a.depid=b.lineid " & _
                       " left join di3670 c on a.empid=c.empid " & _
                       " where a.periodid=" & cQuote & cPeriodID & cQuote & " and a.emp_stat = 2 " & _
                       " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")"
                       
        Else
            cSqlStmt = " select a.periodid, a.empid, a.fullname, a.firstname, a.mname, a.lastname, " & _
                       " a.depid, ifnull(b.linename,'') as department, " & _
                       " ((c.sl_avail+c.vl_avail)-(c.sl_use+c.vl_use))* c.rate_amt as net_pay," & _
                       " 0 as sa_net_pay, 0 as inc_pay, a.active, a.emp_stat " & _
                       " from pa87260 a " & _
                       " left join di5463 b on a.depid=b.lineid " & _
                       " left join di3670 c on a.empid=c.empid " & _
                       " where a.periodid=" & cQuote & cPeriodID & cQuote & " and a.emp_stat <> 0 and c.sl_avail <> 0 " & _
                       " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")"
        
        End If
    End If
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        ShowProgress 0
        
        OpenQueryDNS "select * from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            cPeriodName = "For the Payroll period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
        End If
        
        While Not oTempADO.EOF
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
            aDenomInfo = GetDenom(oTempADO("net_pay"))
            aDenomInfo2 = GetDenom(oTempADO("sa_net_pay"))
            aDenomInfo3 = GetDenom(oTempADO("inc_pay"))
            
            cSqlStmt = "insert into TmpDenom(periodname, depid, deptname, empid, active, emp_stat," & _
                       " fullname, fname, lname, net_pay, sa_net_pay," & _
                       " p1000,p500,p100,p50,p20,p10,p5,p1,pcoin, from_name, attn_name)values(" & _
                       cQuote & EncodeStr2(DecodeStr(cPeriodName)) & cQuote & "," & _
                       cQuote & oTempADO("depid") & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oTempADO("department"))) & cQuote & "," & _
                       cQuote & oTempADO("empid") & cQuote & "," & oTempADO("active") & "," & oTempADO("emp_stat") & "," & _
                       cQuote & EncodeStr2(DecodeStr(oTempADO("fullname"))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oTempADO("firstname"))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oTempADO("lastname"))) & cQuote & "," & _
                       oTempADO("net_pay") & "," & oTempADO("sa_net_pay") & "," & _
                       aDenomInfo(0) + aDenomInfo2(0) + aDenomInfo3(0) & "," & _
                       aDenomInfo(1) + aDenomInfo2(1) + aDenomInfo3(1) & "," & _
                       aDenomInfo(2) + aDenomInfo2(2) + aDenomInfo3(2) & "," & _
                       aDenomInfo(3) + aDenomInfo2(3) + aDenomInfo3(3) & "," & _
                       aDenomInfo(4) + aDenomInfo2(4) + aDenomInfo3(4) & "," & _
                       aDenomInfo(5) + aDenomInfo2(5) + aDenomInfo3(5) & "," & _
                       aDenomInfo(6) + aDenomInfo2(6) + aDenomInfo3(6) & "," & _
                       aDenomInfo(7) + aDenomInfo2(7) + aDenomInfo3(7) & "," & _
                       aDenomInfo(8) + aDenomInfo2(8) + aDenomInfo3(8) & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(Label6.Caption)) & cQuote & ")"
                       
            QueryTemp cSqlStmt, objdbRs, True
            
            
            oTempADO.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "") & "Denomination Report", IIf(Check4.Value = vbChecked, "rpt33666.rpt", "rpt33666s.rpt")
        
        ShowProgress 4
    End If
End Sub



' + -->
' |     Procedure Name  :   ClosePeriod(ByVal cPeriodID As String)
' |     Description     :   Utility to Close a certain Period
' |     Date Created    :   4 apr 2006
' + -->
Function GenField(aTableDef As Variant) As String
    Dim aFieldStru As Variant, _
        nCtr As Integer, _
        cParam As String
        
    DoEvents
    For nCtr = 0 To UBound(aTableDef)
        aFieldStru = aTableDef(nCtr)
        cParam = cParam & "`" & aFieldStru(0) & "`" & IIf(nCtr <> UBound(aTableDef), ",", "")
    Next nCtr
    GenField = cParam
End Function

Sub ClosePeriod(ByVal cPeriodID As String)
    Dim cSqlStmt As String, _
        lProceed As Boolean, _
        oRecordSet As New ADODB.Recordset, _
        aPeriodInfo As Variant
        
    aPeriodInfo = Array("", "", 0, 0, 0)
'    aPeriodInfo(0)      Date Start
'    aPeriodInfo(1)      Date End
'    aPeriodInfo(2)      status
'    aPeriodInfo(3)      Annual Tax
'    aPeriodInfo(4)      13 month period
    
    cSqlStmt = "Warning!!!" & vbCrLf & vbCrLf & _
               "By clicking the [YES] button, you are agreeing" & vbCrLf & _
               "to the condition that all the transactions for the" & vbCrLf & _
               "selected period will be deemed final and executory." & vbCrLf & _
               "All related transactions for this period will be closed" & vbCrLf & _
               "altogether to avoid any further altercations in the near" & vbCrLf & _
               "future for reference and archival purposes." & vbCrLf & vbCrLf & _
               "Click the [YES] button to proceed now."
               
    If MsgBox(cSqlStmt, vbYesNo + vbCritical, "System Advisory") = vbNo Then Exit Sub
    
    If gUserLevel <> 1 Then
        frmManager.Show 1
        If ModalResult = mrCancel Then Exit Sub
        lProceed = ModalResult = mrOk
    Else
        lProceed = gUserLevel = 1
    End If
    
    
    If lProceed Then
        OpenQueryDNS "select date_start, date_end,`status`, wtax, 13month  from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
        aPeriodInfo(0) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("date_start"), Now), "yyyy-mm-dd")
        aPeriodInfo(1) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("date_end"), Now), "yyyy-mm-dd")
        aPeriodInfo(2) = IIf(objdbRs.RecordCount > 0, objdbRs("status"), 0)
        aPeriodInfo(3) = IIf(objdbRs.RecordCount > 0, objdbRs("wtax"), 0)
        aPeriodInfo(4) = IIf(objdbRs.RecordCount > 0, objdbRs("13month"), 0)
        
        If aPeriodInfo(2) > 0 Then
            
            ShowProgress 0, , 1
            
            
            ShowProgress 2, 11, , , "Updating Employee record..."
            
            If aPeriodInfo(4) = 0 Then
            
                If aPeriodInfo(3) = 1 Then
                    ' --> reset usage of union leave
                    OpenQueryDNS "update PA73887 set ul_use=0", objdbRs, True
                    Script2File cSqlStmt
                
                    cSqlStmt = "update pa87260 a, di3670 b set b.ssprem1215=0, " & _
                               "                               b.sser1215=0, " & _
                               "                               b.ps1215=0," & _
                               "                               b.es1215=0," & _
                               "                               b.cola1215=0," & _
                               "                               b.mtd_gross=0," & _
                               "                               b.mtd_basic=0," & _
                               "                               b.mtd_taxable=0," & _
                               "                               b.ytd_cola=0," & _
                               "                               b.ytd_gross=0," & _
                               "                               b.ytd_basic=0," & _
                               "                               b.ytd_wtax=0," & _
                               "                               b.ytd_gross_sa=0 " & _
                               " where (a.empid=b.empid) and (a.periodid=" & cQuote & cPeriodID & cQuote & ")"
                               
'                    cSqlStmt = "update pa87260 a, di3670 b set b.ssprem1215=0, " & _
'                               "                               b.sser1215=0, " & _
'                               "                               b.ps1215=0," & _
'                               "                               b.es1215=0," & _
'                               "                               b.cola1215=0," & _
'                               "                               b.mtd_gross=0," & _
'                               "                               b.mtd_basic=0," & _
'                               "                               b.mtd_taxable=0," & _
'                               "                               b.ytd_cola=0," & _
'                               "                               b.ytd_gross=0," & _
'                               "                               b.ytd_basic=0," & _
'                               "                               b.ytd_wtax=0," & _
'                               "                               b.ytd_gross_sa=0," & _
'                               "                               b.sl_avail=0," & _
'                               "                               b.vl_avail=0," & _
'                               "                               b.sl_use=0," & _
'                               "                               b.vl_use=0" & _
'                               " where (a.empid=b.empid) and (a.periodid=" & cQuote & cPeriodID & cQuote & ")"
                Else
                    ' --> update employee's MTD/YTD info...
                    Select Case aPeriodInfo(2)
                        Case 1  ' --> 1st Period (1-15)
                        
                            cSqlStmt = "update pa87260 a, di3670 b set b.ssprem1215=a.ssprem, " & _
                                       "                               b.sser1215=a.sser, " & _
                                       "                               b.ps1215=a.medicare," & _
                                       "                               b.es1215=a.medicare," & _
                                       "                               b.cola1215=a.cola," & _
                                       "                               b.mtd_gross=a.gross_pay," & _
                                       "                               b.mtd_basic=a.reg_pay+a.ndiff_pay," & _
                                       "                               b.mtd_taxable=a.taxable," & _
                                       "                               b.ytd_cola=b.ytd_cola+a.cola," & _
                                       "                               b.ytd_gross=b.ytd_gross+a.gross_pay-a.m13pay," & _
                                       "                               b.ytd_basic=b.ytd_basic+(a.reg_pay+a.ndiff_pay)," & _
                                       "                               b.ytd_wtax=b.ytd_wtax+a.wtax," & _
                                       "                               b.ytd_gross_sa=b.ytd_gross_sa+a.sa_net_pay" & _
                                       " where (a.empid=b.empid) and (a.periodid=" & cQuote & cPeriodID & cQuote & ")"
                        Case 2  ' --> 2nd Period (16-end of month)
                            cSqlStmt = "update pa87260 a, di3670 b set b.ssprem1215=0, " & _
                                       "                               b.sser1215=0, " & _
                                       "                               b.ps1215=0," & _
                                       "                               b.es1215=0," & _
                                       "                               b.cola1215=0," & _
                                       "                               b.mtd_gross=0," & _
                                       "                               b.mtd_basic=0," & _
                                       "                               b.mtd_taxable=0," & _
                                       "                               b.ytd_cola=b.ytd_cola+a.cola," & _
                                       "                               b.ytd_gross=b.ytd_gross+a.gross_pay-a.m13pay," & _
                                       "                               b.ytd_basic=b.ytd_basic+(a.reg_pay+a.ndiff_pay)," & _
                                       "                               b.ytd_wtax=b.ytd_wtax+a.wtax," & _
                                       "                               b.ytd_gross_sa=b.ytd_gross_sa+a.sa_net_pay" & _
                                       " where (a.empid=b.empid) and (a.periodid=" & cQuote & cPeriodID & cQuote & ")"
                    End Select
                End If
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
        
                ' --> move daily attendance
                ShowProgress 2, 22, , , "Moving Daily Time Record to archive..."
                cSqlStmt = "insert into pah84650(" & GenField(chkPA84650) & ")" & _
                           "select " & GenField(chkPA84650) & _
                           " from PA84650 where logdate between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 2, 33, , , "Deleting DTR entry..."
                cSqlStmt = "delete from pa84650 where logdate between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                
                ' --> move shifting schedule - 20070113
                cSqlStmt = "insert into dih36770(" & GenField(chkDI36770) & ")" & _
                           "select " & GenField(chkDI36770) & _
                           " from di36770 where `date` between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                cSqlStmt = "delete from di36770 where `date` between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            End If
            
            ' --> move payroll header transaction to history
            ShowProgress 2, 44, , , "Moving Payroll transaction to archive..."
            cSqlStmt = "insert into pah87260(" & GenField(chkPA87260) & ")" & _
                       "select " & GenField(chkPA87260) & _
                       " from pa87260 where periodid=" & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            ShowProgress 2, 55, , , "Deleting Payroll transaction..."
            cSqlStmt = "delete from pa87260 where periodid=" & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            ' --> update custom deduction here
            ShowProgress 2, 66, , , "Updating Custom Deduction..."
            cSqlStmt = "update di3673 a, pa87263 b set a.acc_amt = a.acc_amt + b.ded_amt " & _
                       "where (a.empid=b.empid and a.dedid=b.dedid and a.ctrl_no=b.ctrl_no) " & _
                       " and (b.periodid=" & cQuote & cPeriodID & cQuote & ")"
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            ' --> close finished deduction here
            ShowProgress 2, 71, , , "Closing finished deduction..."
            cSqlStmt = "update di3673 set status=1, date_fin=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                       " where cut_off_amt <= acc_amt and status=0"
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            ' --> move deduction to history
            ShowProgress 2, 77, , , "Moving Deduction transaction to archive..."
            cSqlStmt = "insert into pah87263(" & GenField(chkPA87263) & ")" & _
                       "select " & GenField(chkPA87263) & _
                       " from pa87263 where periodid=" & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            ShowProgress 2, 88, , , "Deleting deduction transaction..."
            cSqlStmt = "delete from pa87263 where periodid=" & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            
            ' --> move attendance and dtr error to history
            If lAttendance Then
                cSqlStmt = "insert into att2000h(" & GenField(chkAtt2000) & ")" & _
                           "select " & GenField(chkAtt2000) & _
                           "from att2000 where transdate between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                cSqlStmt = "delete from att2000 where transdate between " & cQuote & aPeriodInfo(0) & cQuote & _
                           " and " & cQuote & aPeriodInfo(1) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            End If
            
            
            ' --> close period
            ShowProgress 2, 95, , , "Closing Period..."
            cSqlStmt = "update pa7730 set pclose=1, date_close=" & cQuote & Format(Now, "yyyy-mm-dd") & cQuote & _
                       " where periodid=" & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
            
            
            Log2Audit Name, "Payroll PeriodID#" & cPeriodID & " successfully closed..."
            
            ShowProgress 4
            
            SetSelection Tag
        End If
    End If
    
    Set oRecordSet = Nothing
End Sub


' + -->
' |     Procedure Name  :   MFCon(ByVal cParam As String, cParam2 As String)
' |     Description     :   Generate Monthly Finish Contract Report
' |     Date Created    :   5 apr 2006
' + -->
Sub CreateMFCon()
        On Error GoTo ErrCreate
    Dim cSqlStmt As String
    cSqlStmt = " CREATE TABLE tmpMFCon( " & _
               " [DATE_HIRE] date,       [DATE_FIN] date, " & _
               " [EMPID] char(6),        [fullname] char(100), " & _
               " [SEX] CHAR(100),        [EMP_STATName] char(100), " & _
               " [LineName] char(100),   [POSNAME] char(100), " & _
               " [CMPName] char(100),    [Remark] char (100), " & _
               " [tcid] char (5),        [emp_stat] integer, " & _
               " [wap] integer,          [paystatus] integer) "
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpMFCon"
    QueryTemp cSqlStmt, oTempADO, True

End Sub

Sub MFCon(ByVal cParam As String, cParam2 As String, nMode As Integer)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        aOtherInfo As Variant, _
        aPeriodInfo As Variant
        
    aPeriodInfo = Array("", "", "")
    aOtherInfo = Array("", "", "", "", "", "")
    
    CreateMFCon
    
    If nMode = 0 Then
        cSqlStmt = " SELECT b.empid,concat(b.lastname,', ', b.firstname, ' ', left(b.mname,1), if(trim(b.mname)='','','.') ) as fullname, " & _
                   "   if(b.paystatus=0,0,1) as paystatus, " & _
                   "   ifnull(c.linename,'') as department, " & _
                   "   ifnull(d.posname,'') as posname, " & _
                   "   b.depid,b.posid,b.date_fin,b.date_hire,b.cmpid,b.emp_stat,b.sex,b.active,b.tcid,b.wap  " & _
                   "FROM di3670 b" & _
                   "   left join di5463 c on b.depid=c.lineid" & _
                   "   left join di7670 d on b.posid=d.posid" & _
                   " where " & cParam & cParam2 & " And (b.active = 0)" & " and (b.paystatus" & IIf(Check3.Value = vbChecked, "=", "<>") & "2)"
    Else
        ' --> retrieve period info here...
        OpenQueryDNS "SELECT * FROM PA7730 WHERE PERIODID=" & cQuote & cParam & cQuote, objdbRs, False
        aPeriodInfo(0) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("DATE_START"), Now), "yyyy-mm-dd")
        aPeriodInfo(1) = Format(IIf(objdbRs.RecordCount > 0, objdbRs("DATE_END"), Now), "yyyy-mm-dd")
        aPeriodInfo(2) = IIf(objdbRs.RecordCount > 0, objdbRs("Duration"), "")
        
'        cSqlStmt = " SELECT b.empid,concat(b.lastname,', ', b.firstname, ' ', left(b.mname,1), if(trim(b.mname)='','','.') ) as fullname, b.depid,b.posid,b.date_fin,b.date_hire,b.cmpid,b.emp_stat,b.sex,b.active,b.tcid FROM pa87260 a " & _
'                   " left join di3670 b on a.empid=b.empid " & _
'                   " where b.date_fin between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & _
'                   " and a.periodid =" & cQuote & cParam & cQuote & cParam2
'        cSqlStmt = cSqlStmt & " and b.active > 0 "


        cSqlStmt = " SELECT b.empid, concat(b.lastname,', ',b.firstname,' ',if(trim(b.mname)='','',concat(left(b.mname,1),'.'))) as fullname, " & _
                   "   if(b.paystatus=0,0,1) as paystatus, " & _
                   "   ifnull(c.linename,'') as department, " & _
                   "   ifnull(d.posname,'') as posname, " & _
                   "   b.depid,ifnull(b.posid,'') as posid, if((b.active=1) or (b.active=3),b.date_res, b.date_fin) as date_fin,b.date_hire,b.cmpid,b.emp_stat,b.sex,b.active,b.tcid,b.wap " & _
                   "FROM di3670 b " & _
                   "   left join di5463 c on b.depid=c.lineid " & _
                   "   left join di7670 d on b.posid=d.posid " & _
                   "Where (((b.date_fin between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & " ) and (b.active=2)) or ((b.date_res between " & cQuote & aPeriodInfo(0) & cQuote & " and " & cQuote & aPeriodInfo(1) & cQuote & " ) and ((b.active=1) or (b.active=3)))) " & _
                   " " & cParam2 & _
                   " and (b.paystatus" & IIf(Check3.Value = vbChecked, "=", "<>") & "2)" & _
                   " order by b.emp_stat,b.wap desc, fullname "
    End If
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving Employee ID#" & oRecordSet("empid")
            
            aOtherInfo(3) = IIf(oRecordSet("sex") = 0, "Male", "Female")
            aOtherInfo(4) = IIf(oRecordSet("emp_stat") = 0, "Wap", IIf(oRecordSet("emp_stat") = 1, "Contractual", "Regular"))
            aOtherInfo(5) = IIf(oRecordSet("active") = 0, "Active ", IIf(oRecordSet("active") = 1, "Resigned", "Finished"))
                        
                
            cSqlStmt = " INSERT INTO tmpMFCon(DATE_HIRE,DATE_FIN,EMPID,fullname,SEX,EMP_STATname,LineName,POSNAME, " & _
                       " remark,CMPName,tcid,emp_stat,paystatus,wap)VALUES(" & _
                       cQuote & Format(oRecordSet("DATE_HIRE"), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & Format(oRecordSet("DATE_FIN"), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & oRecordSet("EMPID") & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oRecordSet("FULLNAME"))) & cQuote & "," & _
                       cQuote & aOtherInfo(3) & cQuote & "," & _
                       cQuote & aOtherInfo(4) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oRecordSet("department"))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oRecordSet("posname"))) & cQuote & "," & _
                       cQuote & aOtherInfo(5) & cQuote & "," & _
                       cQuote & aOtherInfo(2) & cQuote & "," & _
                       cQuote & oRecordSet("tcid") & cQuote & "," & _
                       oRecordSet("emp_stat") & "," & _
                       oRecordSet("paystatus") & "," & _
                       oRecordSet("wap") & ")"
                       
            QueryTemp cSqlStmt, objdbRs, True

            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport IIf(nMode = 0, "Monthly Finish Contract Listing", "Finish Contract Report for Payroll - " & aPeriodInfo(2)) & IIf(Check3.Value = vbChecked, " [EMERGENCY]", ""), IIf(Check5.Value = vbChecked, "LST3670MFC_P", IIf(Check4.Value = vbChecked, "LST3670MFC_S", "LST3670MFC")) & ".RPT", , True
        
        ShowProgress 4
        
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
End Sub


' + -->
' |     Procedure Name  :   GenPEZA(ByVal cPeriodID As String)
' |     Description     :   Generate PEZA Report
' |     Date Created    :   18 apr 2006
' + -->
Sub CreateTmpPEZA()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    cSqlStmt = " CREATE TABLE tmpPEZA( " & _
               " [DURATION] CHAR(100),  [PERIODID] char(5), " & _
               " [LineName] char(100),  [DEPID] char(3), " & _
               " [MSMCNT] integer,      [MSFCNT] integer, " & _
               " [REMCNT] integer,      [REFCNT] integer, " & _
               " [COMCNT] integer,      [COFCNT] integer, " & _
               " [CAMCNT] integer,      [CAFCNT] integer, " & _
               " [MSMAMT] double,       [MSFAMT] double, " & _
               " [REMAMT] double,       [REFAMT] double, " & _
               " [COMAMT] double,       [COFAMT] double, " & _
               " [CAMAMT] double,       [CAFAMT] double, " & _
               " [TOTCNT] integer,      [TOTAMT] double)"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpPEZA"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Sub Save2PEZA(ByVal nMode As Integer, ByVal cQueryString As String)
    Dim cSqlStmt As String, _
        cParam, cParam1, cParam2 As String
    
    ShowProgress 0
    
    OpenQueryDNS cQueryString, oTempADO, False
    If oTempADO.RecordCount > 0 Then
    
        While Not oTempADO.EOF
        
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
            Select Case nMode
                Case 0      ' --> initial data
                    cSqlStmt = "insert into tmpPEZA(periodid, duration, depid,linename,totcnt,totamt," & _
                               "msmcnt,msfcnt,remcnt,refcnt,comcnt,cofcnt,camcnt,cafcnt," & _
                               "msmamt,msfamt,remamt,refamt,comamt,cofamt,camamt,cafamt)values(" & _
                               cQuote & oTempADO("periodid") & cQuote & "," & _
                               cQuote & "Payroll Period " & EncodeStr2(oTempADO("duration")) & cQuote & "," & _
                               cQuote & oTempADO("depid") & cQuote & "," & _
                               cQuote & EncodeStr2(oTempADO("linename")) & cQuote & "," & _
                               "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
                    QueryTemp cSqlStmt, objdbRs, True
                    
                Case 1, 2, 3, 4
                    cParam = IIf(oTempADO("sex") = 0, "m", "f")
                    Select Case nMode
                        Case 1      ' --> management & staff
                            cParam = "ms" & cParam
                        Case 2      ' --> regular
                            cParam = "re" & cParam
                        Case 3      ' --> contractual
                            cParam = "co" & cParam
                        Case 4      ' --> casual
                            cParam = "ca" & cParam
                    End Select
                    cParam1 = cParam & "cnt"
                    cParam2 = cParam & "amt"
                    cSqlStmt = "update tmpPEZA set " & cParam1 & "=" & cParam1 & "+" & oTempADO("cnt") & "," & _
                               cParam2 & "=" & cParam2 & "+ " & oTempADO("gross_pay") & "," & _
                               "totcnt = totcnt + " & oTempADO("cnt") & "," & _
                               "totamt = totamt + " & oTempADO("gross_pay") & _
                               " where (depid=" & cQuote & oTempADO("depid") & cQuote & ")"
                    QueryTemp cSqlStmt, objdbRs, True
                    
            End Select
            
            oTempADO.MoveNext
            
        Wend
        
    End If
    
    ShowProgress 4
End Sub

Sub GenPEZA(ByVal cPeriodID As String)
    Dim cSqlStmt As String, _
        cPeriodTable As String, _
        aSQLQuery As Variant, _
        nCtr As Integer
    
    CreateTmpPEZA
    
    cSqlStmt = "select pclose from pa7730 where periodid=" & cQuote & cPeriodID & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    cPeriodTable = IIf(objdbRs("pclose") = 0, "pa87260", "pah87260")
    
    aSQLQuery = Array(" and (a.emp_stat <> 0) and c.staff ", _
                      " and (a.emp_stat = 2) and (not c.staff) ", _
                      " and (a.emp_stat = 1) and (not c.staff) and (not a.wap) ", _
                      " and (a.emp_stat = 1) and (not c.staff) and a.wap ")
    
    cSqlStmt = "select distinct a.depid, ifnull(b.linename,'') as linename, a.periodid, ifnull(c.duration,'') as duration " & _
               "from " & cPeriodTable & " a left join di5463 b on a.depid=b.lineid " & _
               "left join pa7730 c on a.periodid=c.periodid " & _
               "where a.periodid=" & cQuote & cPeriodID & cQuote
               
    Save2PEZA 0, cSqlStmt
    
    cSqlStmt = "select a.depid, ifnull(b.sex,0) as sex, count(a.empid) as cnt, sum(a.gross_pay) as gross_pay " & _
               "from " & cPeriodTable & " a left join di3670 b on a.empid=b.empid " & _
               "left join di7670 c on a.posid=c.posid " & _
               "where (a.periodid=" & cQuote & cPeriodID & cQuote & ") and (a.paystatus<>2) "
    For nCtr = 0 To 3
        Save2PEZA nCtr + 1, cSqlStmt & aSQLQuery(nCtr) & "group by a.depid, b.sex"
    Next nCtr
    
    QueryTemp "select * from tmpPEZA", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        ShowProgress 0
        ShowProgress 3
        GenerateReport "PEZA REPORT", "rpt7392.rpt"
        ShowProgress 4
    End If
End Sub


' + -->
' |     Procedure Name  :   GenPhilHealth(ByVal nQuarter As Integer, cYear As String)
' |     Description     :   Generate Philhealth/Medicare Remittance Report
' |     Date Created    :   21 july 2006
' + -->

Sub CreateTMPPHealth()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE TMPPHealth( " & _
               " [Month_INFO] char(50), [TEL_NUM] char(50)," & _
               " [CMP_TIN] char(50),    [CMP_PH] char(50), " & _
               " [SIGNATORY] char(50),  [POSNAME] char(50), " & _
               " [EMPID] char(6),       [PHEALTHNUM] char(50), " & _
               " [FIRSTNAME] CHAR(30),  [LASTNAME] char(30),   [MI] char(1), " & _
               " [BR1] integer,          [PS1] double, " & _
               " [ES1] double,           [REM1] char(2)," & _
               " [DATE_EFC] date,       [TAG] integer," & _
               " [COSTID] char(10),   [CDESC] char(100), " & _
               " [WORKID] char(10),   [WDESC] char(100), " & _
               " [COMPCODE] char(4))"
           

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPPHealth"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Function getBracket(ByVal nAmount As Double) As Integer
    Dim cSqlStmt As String
    
    cSqlStmt = "select msal_brac from pa7454 where mtot_cont=" & nAmount
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        getBracket = objdbRs("msal_brac")
    Else
        getBracket = 0
    End If
End Function


Sub GenPhilHealth(ByVal nMonth As String, ByVal nFilter As Integer)


    Dim cSqlStmt, _
        cParam, _
        cPosName As String, _
        aOtherInfo As Variant, _
        aOtherinfo2 As Variant, _
        nCtr As Integer, _
        lClose As Boolean, _
        oRecordSet As New ADODB.Recordset

        aOtherInfo = Array("", "")
        aOtherinfo2 = Array("", "", "", "")

    If Check3.Value <> vbChecked Then
        If Not ChkPersonnel(Text6, "Please Specify Signatory First!!!") Then
            Text6.SetFocus
        Exit Sub
        End If
    End If
    
    OpenQueryDNS "select position from pa2360 where userid=" & cQuote & Text6.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then cPosName = objdbRs("position")

    CreateTMPPHealth
    
    ShowProgress 0
    
'    cSqlStmt = "select empid, active, date_hire, date_res, date_fin, " & _
'               " month(date_res) - ((quarter(date_res)-1)*3) as month_res, " & _
'               " month(date_fin) - ((quarter(date_fin)-1)*3) as month_fin, " & _
'               " month(date_hire) - ((quarter(date_hire)-1)*3) as month_hire " & _
'               "from di3670 " & _
'               "where (paystatus=0) and ((active=0) and (year(date_hire)=" & Combo1.Text & ") and (Month(date_hire)=" & nMonth & "))" & _
'               " or ((active=1) and (year(date_res)=" & Combo1.Text & ") and (Month(date_res)=" & nMonth & "))" & _
'               " or ((active=2) and (year(date_fin)=" & Combo1.Text & ") and (Month(date_fin)=" & nMonth & "))"
'
    
     cSqlStmt = "select empid, active, date_hire, date_res, date_fin, " & _
               " month(date_res) - ((quarter(date_res)-1)*3) as month_res, " & _
               " month(date_fin) - ((quarter(date_fin)-1)*3) as month_fin, " & _
               " month(date_hire) - ((quarter(date_hire)-1)*3) as month_hire, " & _
               " costcenterid, workcenterid " & _
               "from di3670 " & _
               "where (paystatus=0) and ((active=0) and (year(date_hire)=" & Combo1.Text & ") and (Month(date_hire)=" & nMonth & "))" & _
               " or ((active=1) and (year(date_res)=" & Combo1.Text & ") and (Month(date_res)=" & nMonth & "))" & _
               " or ((active=2) and (year(date_fin)=" & Combo1.Text & ") and (Month(date_fin)=" & nMonth & "))"
    
    
    
    
'    Script2File cSqlStmt
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False

    lClose = False

loopd2:

    cSqlStmt = "select month(b.date_start) - ((quarter(b.date_start)-1)*3) as contri_month,C.COSTCENTERID,c.WORKCENTERID, " & _
               "       a.periodid, " & _
               "       a.empid, " & _
               "       ifnull(c.date_hire,curdate()) as date_hire, " & _
               "       ifnull(c.date_res,curdate()) as date_res, " & _
               "       ifnull(c.firstname,'') as firstname, " & _
               "       ifnull(c.mname,'') as mname, " & _
               "       ifnull(c.lastname,'') as lastname, " & _
               "       ifnull(if(trim(c.phealthnum=''),c.sssnum,c.phealthnum),'') as phealthnum, " & _
               "       sum(a.ded_amt) as ded_amt " & _
               "from ((" & IIf(lClose, "pah", "pa") & "87263 a left join " & IIf(lClose, "pah", "pa") & "87260 c on a.periodid=c.periodid and a.empid=c.empid) " & _
               "  left join pa7730 b on a.periodid=b.periodid) " & _
               "where a.periodid in (select periodid From pa7730 " & _
               "    Where (((month(date_start) = " & nMonth & ") And (Year(date_start) = " & Combo1 & ")) Or " & _
               "           ((month(date_end) = " & nMonth & ") And (Year(date_end) = " & Combo1 & "))) " & _
               "          and pclose=" & IIf(lClose, "1", "0") & ") " & _
               "and (a.dedid='005') and (a.ded_amt>0) " & _
               "group by month(b.date_start), a.empid " & _
               "order by a.empid, a.periodid "
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF

'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
'            cSqlStmt = "select * from tmpphealth where empid=" & cQuote & oRecordSet("empid") & cQuote
'            QueryTemp cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                cSqlStmt = " update tmpphealth set ps1 = ps1 + " & oRecordSet("ded_amt") & ", es1 = es1 + " & oRecordSet("ded_amt") & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote
                
 '           Else
 '               cSqlStmt = "insert into tmpphealth(signatory,posname,Month_INFO,tel_num,cmp_tin,cmp_ph,empid,phealthnum,firstname,lastname,mi,ps1,es1,br1,rem1,[tag],date_efc)values(" & _
                           cQuote & Label8.Caption & cQuote & "," & _
                           cQuote & cPosName & cQuote & "," & _
                           cQuote & MonthName(ListView1.SelectedItem.Text) & " " & Combo1.Text & cQuote & "," & _
                           cQuote & gTelNum & cQuote & "," & _
                           cQuote & gTINNum & cQuote & "," & _
                           cQuote & gPHealthNum & cQuote & "," & _
                           cQuote & oRecordSet("empid") & cQuote & "," & _
                           cQuote & oRecordSet("phealthnum") & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(oRecordSet("firstname"))) & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(oRecordSet("lastname"))) & cQuote & "," & _
                           cQuote & left(oRecordSet("mname"), 1) & cQuote & "," & _
                           oRecordSet("ded_amt") & "," & oRecordSet("ded_amt") & "," & _
                           "0," & cQuote & cQuote & ",0," & _
                           cQuote & Format(oRecordSet("date_hire"), "mm/dd/yyyy") & cQuote & ")"
  '          End If
            
'            MsgBox cSqlStmt
 '           QueryTemp cSqlStmt, objdbRs, True
 '           oRecordSet.MoveNext
 '       Wend
'  End If

    
'---> Revision for ERP 20120
If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
'            If oRecordSet("EMPID") = "191378" Then
'                MsgBox "Stop!!!"
'            End If
            
        OpenQueryDNS "SELECT COSTCENTERID, DESCRIPTION, COMPCODE FROM pa37722 where costcenterid = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherinfo2(0) = objdbRs("DESCRIPTION")
                aOtherinfo2(1) = objdbRs("COMPCODE")
            Else
                aOtherinfo2(0) = ""
                aOtherinfo2(1) = ""
            End If


        OpenQueryDNS "SELECT WORKCENTERID, DESCRIPTION, COMPCODE FROM pa97722 where workcenterid = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherinfo2(2) = objdbRs("DESCRIPTION")
                aOtherinfo2(3) = objdbRs("COMPCODE")
            Else
                aOtherinfo2(2) = ""
                aOtherinfo2(3) = ""
            End If
            
            cSqlStmt = "select * from tmpphealth where empid=" & cQuote & oRecordSet("empid") & cQuote
        QueryTemp cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                cSqlStmt = " update tmpphealth set ps1 = ps1 + " & oRecordSet("ded_amt") & ", es1 = es1 + " & oRecordSet("ded_amt") & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote
                
            Else
              cSqlStmt = "insert into tmpphealth(signatory,posname,COSTID, CDESC, COMPCODE,WORKID, WDESC, Month_INFO,tel_num,cmp_tin,cmp_ph,empid,phealthnum,firstname,lastname,mi,ps1,es1,br1,rem1,[tag],date_efc)values(" & _
                           cQuote & Label8.Caption & cQuote & "," & _
                           cQuote & cPosName & cQuote & "," & _
                           cQuote & oRecordSet("COSTCENTERID") & cQuote & "," & _
                           cQuote & aOtherinfo2(0) & cQuote & "," & _
                           cQuote & aOtherinfo2(1) & cQuote & "," & _
                           cQuote & oRecordSet("WORKCENTERID") & cQuote & "," & _
                           cQuote & aOtherinfo2(2) & cQuote & "," & _
                           cQuote & MonthName(ListView1.SelectedItem.Text) & " " & Combo1.Text & cQuote & "," & _
                           cQuote & gTelNum & cQuote & "," & _
                           cQuote & gTINNum & cQuote & "," & _
                           cQuote & gPHealthNum & cQuote & "," & _
                           cQuote & oRecordSet("empid") & cQuote & "," & _
                           cQuote & oRecordSet("phealthnum") & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(oRecordSet("firstname"))) & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(oRecordSet("lastname"))) & cQuote & "," & _
                           cQuote & left(oRecordSet("mname"), 1) & cQuote & "," & _
                           oRecordSet("ded_amt") & "," & oRecordSet("ded_amt") & "," & _
                           "0," & cQuote & cQuote & ",0," & _
                           cQuote & Format(oRecordSet("date_hire"), "mm/dd/yyyy") & cQuote & ")"

            End If
            
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            oRecordSet.MoveNext
        Wend
  End If
        
    If Not lClose Then
        lClose = True
        GoTo loopd2
    End If

    QueryTemp "select empid,ps1,es1 from tmpphealth", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100

            cParam = ""

            oTempADO.Requery adAsyncFetch
            oTempADO.Find "EMPID='" & oRecordSet("EMPID") & "'"
            If Not oTempADO.EOF Then
                Select Case oTempADO("active")
                    Case 0
                        cParam = "tag=1," & _
                                 "rem1 =" & cQuote & "NH" & cQuote & "," & _
                                 "date_efc=" & cQuote & Format(oTempADO("date_hire"), "mm/dd/yyyy") & cQuote & ","
                    Case 1, 2
                        cParam = "tag=1," & _
                                 "rem1 =" & cQuote & "S" & cQuote & "," & _
                                 "date_efc=" & cQuote & Format(IIf(oTempADO("active") = 1, oTempADO("date_res"), oTempADO("date_fin")), "mm/dd/yyyy") & cQuote & ","
                End Select
            End If

            If Trim(cParam) = "" Then
                If oRecordSet("ps1") + oRecordSet("es1") = 0 Then
                    cParam = IIf(Trim(cParam) = "", "tag=2,", cParam) & _
                             "rem1=" & cQuote & "NE" & cQuote & ","
                End If
            End If

            cSqlStmt = "update tmpphealth set " & _
                       IIf(Trim(cParam) = "", "", cParam) & _
                       " br1=" & getBracket(oRecordSet("ps1") + oRecordSet("es1")) & _
                       " where empid=" & cQuote & oRecordSet("empid") & cQuote
            QueryTemp cSqlStmt, objdbRs, True
            oRecordSet.MoveNext
        Wend

            ShowProgress 3
        
        If vbYes = MsgBox("Generate Ascii File", vbCritical + vbYesNo, "Generate Ascii File") Then

            If Text4.Text <> "" Then

                GenMedSavetoFIle Text4.Text, Combo1.Text, ListView1.SelectedItem.Text
            Else
                MsgBox "Please Specify location for the ASCII FIle", vbInformation, "System Advisory!!!"
            End If

        Else
        

            ShowProgress 3
            If Check3.Value = vbUnchecked Then
            
                GenerateReport "Philhealth Employer Monthly Remittance Report", "rpt74n.rpt"
            Else
                GenerateReport "Philhealth Employer Monthly Remittance Report", "rpt74n_cost.rpt"
            End If
            

        End If
    Else
    
        ShowProgress 3
        
        MsgBox "No Report to generate!"
    End If

    ShowProgress 4

    Set oRecordSet = Nothing
End Sub

Sub GenMedSavetoFIle(ByVal cPath As String, ByVal cYear As String, ByVal cMOnth As String)
    Dim oTextFile As New FileSystemObject, _
        oTxtStream As TextStream, _
        FileSys As FileSystemObject, _
        oFile As File, _
        cMedFile As String, _
        cSqlStmt As String, _
        cString As String, _
        aContribution As Variant, _
        nPhilTotal As Variant, _
        aUserInfo As Variant, _
        cWant As String
        
    
    Set FileSys = New FileSystemObject
    
    'cMedFile = CheckPath(cPath) & "ph" & cYear & cQuarter & ".ASCII"
    cMedFile = CheckPath(cPath) & "ph" & cYear & cMOnth & ".ASCII"
    
    If FileSys.FileExists(cMedFile) = True Then
        cWant = MsgBox(cMedFile & " Already Exsist... do you want to overwrite the file ?", vbYesNo + vbCritical, App.Title)
        If cWant = vbYes Then
            FileSys.DeleteFile cMedFile
        Else
            Exit Sub
        End If
    End If

'    aContribution(0) = PHEALTHNUM & LASTNAME & FIRSTNAME & MI & "00000000"
'    aContribution(1) = PS1
'    aContribution(2) = ES1
'    aContribution(3) = REM1
    aContribution = Array("", "", "", "")
    
'    aUserinfo(0) = Signatory Fullname
'    aUserinfo(1) = Position
    aUserInfo = Array("", "")
    
'    nPhilTotal(0) = SP1
'    nPhilTotal(1) = EP1
'    nPhilTotal(2) = manpower 1
    nPhilTotal = Array(0#, 0#, 0#)
    
    
    ShowProgress 0
    
    If Dir(cMedFile) = "" Then
        Set oTxtStream = oTextFile.CreateTextFile(cMedFile, True)
    Else
        Set oFile = oTextFile.GetFile(cMedFile)
        Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
    End If
   
    oTxtStream.WriteLine "REMITTANCE REPORT"
    OpenQueryDNS "select * from di2660 where cmpid = " & gCompanyID, objdbRs, False
    oTxtStream.WriteLine IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    oTxtStream.WriteLine IIf(objdbRs.RecordCount > 0, objdbRs("cmpaddress1"), "")
    oTxtStream.WriteLine checkchar(gPHealthNum, "-") & " " & cMOnth & cYear & left(Combo1.Text, 1)
    
    oTxtStream.WriteLine "MEMBERS"
    
    QueryTemp "select * from TMPPHealth order by lastname, firstname, mi", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        
        aUserInfo(0) = objdbRs("signatory")
        aUserInfo(1) = objdbRs("posname")
        
        While Not objdbRs.EOF
        
            ShowProgress 2, (objdbRs.AbsolutePosition / objdbRs.RecordCount) * 100
                        
            aContribution(0) = objdbRs("PHEALTHNUM") & " " & objdbRs("LASTNAME") & " " & objdbRs("FIRSTNAME") & " " & objdbRs("MI") & "00000000"
            aContribution(1) = PadStr(checkchar(Format(objdbRs("PS1"), "###0.00"), "."), "0", objdbRs("PS1").DefinedSize - 2)
            aContribution(2) = PadStr(checkchar(Format(objdbRs("ES1"), "###0.00"), "."), "0", objdbRs("ES1").DefinedSize - 2)
                    
            If Trim(objdbRs("REM1")) <> "" Then
                aContribution(3) = Trim(objdbRs("REM1")) & Format(objdbRs("DATE_EFC"), "mmddyyyy")
            Else
                aContribution(3) = ""
            End If

            nPhilTotal(0) = nPhilTotal(0) + objdbRs("PS1")
            nPhilTotal(1) = nPhilTotal(1) + objdbRs("ES1")
            
            If (objdbRs("PS1") <> 0) Or (objdbRs("ES1") <> 0) Then nPhilTotal(2) = nPhilTotal(2) + 1
            
            cString = aContribution(0) & aContribution(1) & aContribution(2) & aContribution(3)
            oTxtStream.WriteLine cString
            
            objdbRs.MoveNext
        Wend
    End If
    
    oTxtStream.WriteLine "M5-SUMMARY"
    cString = ""
    cString = Val(nPhilTotal(0)) + Val(nPhilTotal(1))
    cString = "1" & PadStr(checkchar(Format(cString, "###0.00"), "."), "0", 8) & "      " & "XXXXXXXX" & "      " & "XXXXXXXX" & "      " & nPhilTotal(2)
    oTxtStream.WriteLine cString
    
    cString = Val(nPhilTotal(0)) + Val(nPhilTotal(1))
    oTxtStream.WriteLine "GRAND TOTAL " & PadStr(checkchar(Format(cString, "###0.00"), "."), "0", 10)
    
    oTxtStream.WriteLine Trim(aUserInfo(0)) & " " & Trim(aUserInfo(1))
    
    oTxtStream.Close

    ShowProgress 4

    Set oTxtStream = Nothing
    Set oTextFile = Nothing
    Set oFile = Nothing
    
End Sub


Private Function checkchar(ByVal cString As String, ByVal cChar As String) As String
    Dim nCtr As Integer
    
    ReDim aLabelID(99)
    While InStr(1, cString, cChar) > 0
        aLabelID(nCtr) = left(cString, InStr(1, cString, cChar) - 1)
        cString = Mid(cString, InStr(1, cString, cChar) + 1, Len(cString) - InStr(1, cString, cChar))
        nCtr = nCtr + 1
    Wend
    
    If Trim(cString) <> "" Then aLabelID(nCtr) = cString
    
    checkchar = aLabelID(0) & aLabelID(1) & aLabelID(2) & aLabelID(3) & aLabelID(4) & aLabelID(5)
    
End Function

' + -->
' |     Procedure Name  :   GenLeaveRep(cParam As String, nmode As Integer)
' |     Description     :   Generate Leave Report
' |     Date Created    :   27 july 2006
' + -->
Sub Createleaverep()
        On Error GoTo ErrCreate
    Dim cSqlStmt As String

    cSqlStmt = " CREATE TABLE Tmp367583( " & _
               " [leave_no] char(10),           [date_leave] date, " & _
               " [date_start] date,             [date_end] date, " & _
               " [duration] char(30), " & _
               " [sl_avail_cnt] double,         [vl_avail_cnt] double, " & _
               " [el_avail_cnt] double,         [ml_avail_cnt] double, " & _
               " [pl_avail_cnt] double,         [fl_avail_cnt] double, " & _
               " [ul_avail_cnt] double,         [Leavename] char(100), " & _
               " [tag] double,                  [paytag] double, " & _
               " [EMPID] char(6),               [fullname] char(100), " & _
               " [LineName] char(100),          [POSNAME] char(100), " & _
               " [rate_amt] double,             [CMPName] char(100), " & _
               " [prep_by] char(6),             [chk_by] char(6), " & _
               " [noted_by] char(6),            [appr_by] char(6), " & _
               " [prep_name] char(100),         [chk_name] char(100),  " & _
               " [noted_name] char(100),        [appr_name] char(100), " & _
               " [prep_pos] char(100),          [chk_pos] char(100),  " & _
               " [noted_pos] char(100),         [appr_pos] char(100))"
               

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM Tmp367583"
    QueryTemp cSqlStmt, oTempADO, True

End Sub

Sub Genleaverep(cParam As String, nMode As Integer)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cString As String, _
        cParam2 As String, _
        cParam3 As String, _
        cParam4 As String, _
        aOtherDate As Variant, _
        aUserInfo As Variant, _
        aLeave As Variant, _
        aCounter As Variant, _
        dDate_end As Date, _
        dDate_Start As Date, _
        nCtr As Integer
    
    aCounter = Array(0#, 0#, 0#)
    aOtherDate = Array("", "", "")
    aLeave = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aUserInfo = Array("", "", "", "", "")
    
    If Not ChkPersonnel(Text6) Then Exit Sub
    If Not ChkPersonnel(Text5) Then Exit Sub
    If Not ChkPersonnel(Text7) Then Exit Sub
    If Not ChkPersonnel(Text8) Then Exit Sub
    
    Createleaverep

    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")

        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aUserInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
        
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aUserInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")

        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text8.Text & "'"
        aUserInfo(3) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If
    
    OpenQueryDNS "select * from PA7730 where periodid = " & cQuote & cParam & cQuote, oRecordSet, False
    aOtherDate(0) = IIf(oRecordSet.RecordCount > 0, Format(oRecordSet("date_start"), "yyyy-mm-dd"), "")
    aOtherDate(1) = IIf(oRecordSet.RecordCount > 0, Format(oRecordSet("date_end"), "yyyy-mm-dd"), "")
    aOtherDate(2) = IIf(oRecordSet.RecordCount > 0, oRecordSet("duration"), "")
    
    cParam2 = "and ((date_start between " & cQuote & aOtherDate(0) & cQuote & " and " & cQuote & aOtherDate(1) & cQuote & ") or (date_end between " & cQuote & aOtherDate(0) & cQuote & " and " & cQuote & aOtherDate(1) & cQuote & "))  and a.paytag <> 1 order by a.empid"
    Select Case nMode
        Case 0
            cParam2 = " where a.Tag < 2 " & cParam2
        Case 1
            cParam2 = " where a.Tag = 2 " & cParam2
        Case 2
            cParam2 = " where a.Tag = 3 " & cParam2
        Case 3
            cParam2 = " where a.Tag = 4 " & cParam2
        Case 4
            cParam2 = " where a.Tag = 5 " & cParam2
        Case 5
            cParam2 = " where a.Tag = 6 " & cParam2
        Case Else
            cParam2 = "where a.tag < 6 " & cParam2
        
    End Select
    
    cSqlStmt = " SELECT a.leave_no, a.date_leave, a.date_start,a.date_end, a.empid,concat(b.lastname,', ', b.firstname) as fullname, " & _
               " c.linename,d.posname, ifnull(a.leave_cnt,0) as leave_cnt, ifnull(a.Tag,0) as tag, ifnull(a.paytag,0) as paytag, a.ul_avail,ifnull(rate_amt,0) as rate_amt " & _
               " FROM pa367583 a left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid left join di7670 d on b.posid=d.posid "
               
'    MsgBox cSqlStmt & " " & cParam2
    OpenQueryDNS cSqlStmt & cParam2, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        ShowProgress 0
        cString = ""
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("FULLNAME")) & "..."
            
            dDate_Start = IIf(Format(oRecordSet("date_start"), "yyyy-mm-dd") > aOtherDate(0), oRecordSet("date_start"), aOtherDate(0))
            
            dDate_end = IIf(Format(oRecordSet("date_end"), "yyyy-mm-dd") > aOtherDate(1), aOtherDate(1), oRecordSet("date_end"))
            
            aCounter = Array(0, 0#, 0#)
            
            nCtr = 0
            
            
            
            For nCtr = Day(dDate_Start) To Day(dDate_end)
                
                 aCounter(0) = aCounter(0) + 1
                
                 If Weekday(dDate_Start + aCounter(0) - 1) = vbSunday Then
                    aCounter(1) = aCounter(1) + 1
                 End If
            Next nCtr
            
            'holiday.... if the system detect holiday, the counter will not increment by 1
        
            OpenQueryDNS " select * from PA4329 Where Date between " & cQuote & Format(dDate_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(dDate_end, "yyyy-mm-dd") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aCounter(2) = objdbRs.RecordCount
            End If
            
            If nMode <> 0 Then
                Select Case nMode
                    Case 1
                        cParam3 = "el_avail_cnt"
                        cParam4 = aCounter(0) - aCounter(1)
                    Case 2
                        cParam3 = "ml_avail_cnt"
                        cParam4 = aCounter(0) - aCounter(1)
                    Case 3
                        cParam3 = "pl_avail_cnt"
                        cParam4 = aCounter(0) - aCounter(1)
                    Case 4
                        cParam3 = "fl_avail_cnt"
                        cParam4 = aCounter(0) - aCounter(1)
                    Case 5
                        cParam3 = "ul_avail_cnt"
                        cParam4 = aCounter(0) - aCounter(1)
                        
                    Case Else
                        aLeave = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
                        cParam4 = aCounter(0) - aCounter(1)
                        aLeave(oRecordSet("tag")) = cParam4
                                
                        cParam3 = "sl_avail_cnt,vl_avail_cnt,el_avail_cnt,ml_avail_cnt,pl_avail_cnt,fl_avail_cnt,ul_avail_cnt"
                        cParam4 = aLeave(0) & "," & aLeave(1) & "," & aLeave(2) & "," & aLeave(3) & "," & aLeave(4) & "," & aLeave(5) & "," & aLeave(6)

                End Select
                
                cSqlStmt = " INSERT INTO Tmp367583(leave_no,date_leave,date_start,date_end,duration," & cParam3 & ",tag,paytag,EMPID,fullname,LineName,POSNAME, " & _
                           " prep_by,chk_by,noted_by,appr_by,prep_name,chk_name,noted_name,appr_name,prep_pos,chk_pos,noted_pos,appr_pos,Leavename,rate_amt)VALUES(" & _
                           cQuote & oRecordSet("leave_no") & cQuote & "," & cQuote & Format(oRecordSet("date_leave"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oRecordSet("date_start"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("date_end"), "yyyy-mm-dd") & cQuote & "," & cQuote & aOtherDate(2) & cQuote & "," & _
                           cParam4 & "," & oRecordSet("tag") & "," & oRecordSet("paytag") & "," & _
                           cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("fullname") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & _
                           cQuote & oRecordSet("posname") & cQuote & "," & _
                           cQuote & Text6.Text & cQuote & "," & cQuote & Text5.Text & cQuote & "," & cQuote & Text7.Text & cQuote & "," & _
                           cQuote & Text8.Text & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label6.Caption)) & cQuote & "," & _
                           cQuote & EncodeStr2(DecodeStr(Label15.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label16.Caption)) & cQuote & "," & _
                           cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & "," & _
                           cQuote & aUserInfo(3) & cQuote & "," & _
                           cQuote & IIf(Combo1.ListIndex = 6, "Leave", Combo1.Text) & cQuote & "," & oRecordSet("rate_amt") & ")"
            
            Else
                
                cParam4 = aCounter(0) - aCounter(1) - aCounter(2)
                If oRecordSet("tag") = 1 Then
                    If Trim(cString) = oRecordSet("empid") Then
                        cSqlStmt = " update Tmp367583 set vl_avail_cnt =" & cParam4 & " where empid = " & cQuote & oRecordSet("empid") & cQuote
                        cString = ""
                    Else
                        cSqlStmt = " INSERT INTO Tmp367583(leave_no,date_leave,date_start,date_end,duration,vl_avail_cnt,tag,paytag,EMPID,fullname,LineName,POSNAME, " & _
                                   " prep_by,chk_by,noted_by,appr_by,prep_name,chk_name,noted_name,appr_name,prep_pos,chk_pos,noted_pos,appr_pos,Leavename,rate_amt)VALUES(" & _
                                   cQuote & oRecordSet("leave_no") & cQuote & "," & cQuote & Format(oRecordSet("date_leave"), "yyyy-mm-dd") & cQuote & "," & _
                                   cQuote & Format(oRecordSet("date_start"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("date_end"), "yyyy-mm-dd") & cQuote & "," & cQuote & aOtherDate(2) & cQuote & "," & _
                                   cParam4 & "," & oRecordSet("tag") & "," & oRecordSet("paytag") & "," & _
                                   cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("fullname") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & _
                                   cQuote & oRecordSet("posname") & cQuote & "," & _
                                   cQuote & Text6.Text & cQuote & "," & cQuote & Text5.Text & cQuote & "," & cQuote & Text7.Text & cQuote & "," & _
                                   cQuote & Text8.Text & cQuote & "," & _
                                   cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label6.Caption)) & cQuote & "," & _
                                   cQuote & EncodeStr2(DecodeStr(Label15.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label16.Caption)) & cQuote & "," & _
                                   cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & "," & _
                                   cQuote & aUserInfo(3) & cQuote & "," & _
                                   cQuote & Combo1.Text & cQuote & "," & oRecordSet("rate_amt") & ")"

                    End If
                Else
                    cSqlStmt = " INSERT INTO Tmp367583(leave_no,date_leave,date_start,date_end,duration,sl_avail_cnt,vl_avail_cnt,tag,paytag,EMPID,fullname,LineName,POSNAME, " & _
                               " prep_by,chk_by,noted_by,appr_by,prep_name,chk_name,noted_name,appr_name,prep_pos,chk_pos,noted_pos,appr_pos,Leavename,rate_amt)VALUES(" & _
                               cQuote & oRecordSet("leave_no") & cQuote & "," & cQuote & Format(oRecordSet("date_leave"), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & Format(oRecordSet("date_start"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("date_end"), "yyyy-mm-dd") & cQuote & "," & cQuote & aOtherDate(2) & cQuote & "," & _
                               cParam4 & "," & 0 & "," & oRecordSet("tag") & "," & oRecordSet("paytag") & "," & _
                               cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("fullname") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & _
                               cQuote & oRecordSet("posname") & cQuote & "," & _
                               cQuote & Text6.Text & cQuote & "," & cQuote & Text5.Text & cQuote & "," & cQuote & Text7.Text & cQuote & "," & _
                               cQuote & Text8.Text & cQuote & "," & _
                               cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label6.Caption)) & cQuote & "," & _
                               cQuote & EncodeStr2(DecodeStr(Label15.Caption)) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label16.Caption)) & cQuote & "," & _
                               cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & "," & _
                               cQuote & aUserInfo(3) & cQuote & "," & _
                               cQuote & Combo1.Text & cQuote & "," & oRecordSet("rate_amt") & ")"
                    cString = oRecordSet("empid")
                End If
            End If
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        Wend
        ShowProgress 3
        cParam2 = ""
        Select Case nMode
            Case 0
                cParam2 = "PRV367583"
            Case 1
                cParam2 = "PRV367583_El"
            Case 2
                cParam2 = "PRV367583_Ml"
            Case 3
                cParam2 = "PRV367583_Pl"
            Case 4
                cParam2 = "PRV367583_Fl"
            Case 5
                cParam2 = "PRV367583_Ul"
            Case Else
                cParam2 = "PRV367583_All"
        End Select
                         
        GenerateReport "Leave report", cParam2 & ".RPT", , True
        
        ShowProgress 4
    Else
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
    End If
    
End Sub

' + -->
' |     Procedure Name  :   GenWithRep(cParam As String)
' |     Description     :   Generate Withholding Tax Report
' |     Date Created    :   27 july 2006
' + -->
Sub CreateWithTax()
        On Error GoTo ErrCreate
    Dim cSqlStmt As String

    cSqlStmt = " CREATE TABLE TmpWithTax( [PERIODID]  char(5), " & _
               " [EMPID]  char(6),       [TINNUM]  char(15), " & _
               " [FIRSTNAME]  char(25),  [MNAME]  char(25), " & _
               " [LASTNAME]  char(25),   [DED_AMT]  double, " & _
               " [WTMONTH] char(30),     [WTYEAR] char(30))"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TmpWithTax"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Sub GenWithRep(cParam As String)
    Dim oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        cSqlStmt As String, _
        cPeriodID As String, _
        cMOnth As String, _
        cYear As String
    
    CreateWithTax
    
    cSqlStmt = " select * from PA7730 " & _
               " where (13month=0) and ((month(date_start)=" & cParam & ") and (year(date_start) = " & cQuote & Combo1.Text & cQuote & " )) and " & _
               " ((Month(date_end) = " & cParam & ") And (Year(date_end) = " & cQuote & Combo1.Text & cQuote & "))"
    
    OpenQueryDNS cSqlStmt, oRSet, False
    If oRSet.RecordCount > 0 Then
        cMOnth = MonthName(Month(oRSet("date_end")))
        cYear = Year(oRSet("date_end"))
        While Not oRSet.EOF
            If oRSet("pclose") <> 1 Then
                cSqlStmt = " select a.periodid,a.empid,ifnull(b.lastname,'') as lastname, " & _
                           " ifnull(b.firstname,'') as firstname,ifnull(b.mname,'') as mname,ifnull(b.tinnum,'') as tinnum,a.ded_amt from pa87263 a " & _
                           " left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
                           " Where a.dedid = " & cQuote & "006" & cQuote & " And a.periodid = " & cQuote & oRSet("periodid") & cQuote & " order by a.empid "
                
            Else
                cSqlStmt = " select a.periodid,a.empid,ifnull(b.lastname,'') as lastname, " & _
                           " ifnull(b.firstname,'') as firstname,ifnull(b.mname,'') as mname,ifnull(b.tinnum,'') as tinnum,a.ded_amt from pah87263 a " & _
                           " left join pah87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
                           " Where a.dedid = " & cQuote & "006" & cQuote & " And a.periodid = " & cQuote & oRSet("periodid") & cQuote & " order by a.empid "
            End If
            
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, oRecordSet, False
            If oRecordSet.RecordCount > 0 Then
                ShowProgress 0
                While Not oRecordSet.EOF
                
                ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
                
                If oRecordSet("ded_amt") <> 0 Then
                    QueryTemp " select periodid,empid from TmpWithTax where empid = " & cQuote & oRecordSet("empid") & cQuote, oRSet2, False
                    If oRSet2.RecordCount > 0 Then
                        'dito yung update
                        cSqlStmt = " update TmpWithTax set ded_amt = ded_amt + " & oRecordSet("ded_amt") & _
                                   " where empid = " & cQuote & oRecordSet("empid") & cQuote
                    Else
                        'insert naman d2
                        cSqlStmt = " INSERT INTO TmpWithTax(PERIODID,EMPID,TINNUM,FIRSTNAME,MNAME,LASTNAME,DED_AMT,WTMONTH,WTYEAR)VALUES(" & _
                                   cQuote & oRecordSet("periodid") & cQuote & "," & _
                                   cQuote & oRecordSet("empid") & cQuote & "," & _
                                   cQuote & oRecordSet("tinnum") & cQuote & "," & _
                                   cQuote & EncodeStr2(DecodeStr(oRecordSet("firstname"))) & cQuote & "," & _
                                   cQuote & EncodeStr2(DecodeStr(oRecordSet("mname"))) & cQuote & "," & _
                                   cQuote & EncodeStr2(DecodeStr(oRecordSet("lastname"))) & cQuote & "," & _
                                   oRecordSet("ded_amt") & "," & cQuote & cMOnth & cQuote & "," & cQuote & cYear & cQuote & ")"
                    End If
                    
'                    MsgBox cSqlStmt
                    QueryTemp cSqlStmt, objdbRs, True
                    End If
                    oRecordSet.MoveNext
                Wend
            End If
            
            oRSet.MoveNext
        Wend
        ShowProgress 3
        
        GenerateReport "Individual Withheld Tax Report", "PRVWithTax.RPT", , True
        
        ShowProgress 4
    End If
End Sub


' + -->
' |     Procedure Name  :   GenPagIbig(ByVal cParam As String)
' |     Description     :   Generate Pag-Ibig Remittance Report (Premium & Loan)
' |     Date Created    :   28 july 2006
' + -->
Sub CreateTmpHDMF()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String

'    cSqlStmt = "CREATE TABLE TmpHDMF(   [CMP_PAGIBIG] char(20), " & _
'               " [MONTH_INFO] char(50), [TEL_NUM] char(50)," & _
'               " [SIGNATORY] char(50),  [POSNAME] char(50), " & _
'               " [SIGNATORY2] char(50), [POSNAME2] char(50), " & _
'               " [EMPID]  char(6),      [PAGIBIGNUM]  char(15), " & _
'               " [FIRSTNAME]  char(25), [MNAME]  char(25), " & _
'               " [LASTNAME]  char(25),  [DED_AMT]  double, " & _
'               " [DED_AMT2]  double,    [TOT_AMT]  double, " & _
'               " [date_grant] date,     [ref_no] char(15))"
               
               
    cSqlStmt = "CREATE TABLE TmpHDMF(       [CMP_PAGIBIG] char(20), " & _
               " [MONTH_INFO] char(50),     [TEL_NUM] char(50)," & _
               " [SIGNATORY] char(50),      [POSNAME] char(50), " & _
               " [SIGNATORY2] char(50),     [POSNAME2] char(50), " & _
               " [EMPID]  char(6),          [PAGIBIGNUM]  char(15), " & _
               " [FIRSTNAME]  char(25),     [MNAME]  char(25), " & _
               " [LASTNAME]  char(25),      [DED_AMT]  double, " & _
               " [DED_AMT2]  double,        [TOT_AMT]  double, " & _
               " [date_grant] date,         [ref_no] char(15), " & _
               " [COSTCENTERID] char(10),   [DESCRIPTION] char(100), " & _
               " [COMPCODE] char(4))"
               


    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TmpHDMF"
    QueryTemp cSqlStmt, oTempADO, True
End Sub
Sub GenPagIbig(ByVal cParam As String)


    Dim cSqlStmt, _
        cPosName, _
        cPeriod As String, _
        aOtherInfo As Variant, _
        aOtherinfo2 As Variant, _
        nDefAmt As Double
    
    aOtherInfo = Array("", "")
    aOtherinfo2 = Array("", "")
    
    If Not ChkPersonnel(Text6, "Please specify signatory first!!!") Then
        SSTab1.Tab = 3
        Text6.SetFocus
        Exit Sub
    End If
    
    If Check4.Value <> vbChecked Then
        If Not ChkPersonnel(Text6, "Please specify signatory first!!!") Then
            SSTab1.Tab = 3
            Text6.SetFocus
            Exit Sub
        End If
    End If
    
    If Check3.Value <> vbChecked Then
        If Not ChkPersonnel(Text6, "Please specify signatory first!!!") Then
            SSTab1.Tab = 3
            Text6.SetFocus
            Exit Sub
        End If
    End If
       
    CreateTmpHDMF
    
    OpenQueryDNS "select position from pa2360 where userid=" & cQuote & Text6.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then aOtherInfo(0) = objdbRs("position")
    
    OpenQueryDNS "select position from pa2360 where userid=" & cQuote & Text5.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then aOtherInfo(1) = objdbRs("position")
    
    If Check4.Value = vbChecked Then
        cSqlStmt = "select def_amt from pa3330 where dedid='003'"
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then nDefAmt = objdbRs("def_amt")
    End If
    
    cSqlStmt = " select periodid from PA7730 " & _
               " where (13month=0) and ((month(date_start)=" & cParam & ") and (year(date_start) = " & cQuote & Combo1.Text & cQuote & " )) and " & _
               " ((Month(date_end) = " & cParam & ") And (Year(date_end) = " & cQuote & Combo1.Text & cQuote & "))"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            cPeriod = cPeriod & cQuote & objdbRs("periodid") & cQuote & ","
            objdbRs.MoveNext
        Wend
        
    End If
        
'    If Trim(cPeriod) <> "" Then cPeriod = "(" & left(cPeriod, Len(cPeriod) - 1) & ")"

'    If Trim(cPeriod) = "" Then cPeriod = "(" & cQuote & ("") & cQuote & ")"


'     If Trim(cPeriod) <> "" Then
'        cPeriod = "(" & left(cPeriod, Len(cPeriod) - 1) & ")"
'     Else
'        cPeriod = "(" & cQuote & ("") & cQuote & ")"
'     End If
     
        
'----> remarked 201406-03
    If Trim(cPeriod) <> "" Then
        cPeriod = "(" & left(cPeriod, Len(cPeriod) - 1) & ")"
    Else
         MsgBox "No data to process!", vbInformation, "System Advisory!!!"

    Exit Sub

    End If
    
     
' --> remarked 20070705
'    cSqlStmt = "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, b.firstname, b.mname, b.lastname, b.pagibigno " & _
'               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
'               "from pah87263 a left join pah87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
'               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
'               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", "004") & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
'               "Union All " & _
'               "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, b.firstname, b.mname, b.lastname, b.pagibigno " & _
'               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
'               "from pa87263 a left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
'               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
'               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", "004") & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
'               "order by periodid,firstname "

    cSqlStmt = ""
    
'    cSqlStmt = "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, if(b.firstname is null,0,a.ded_amt2) as ded_amt2, b.firstname, b.mname, b.lastname, b.pagibigno,b.COSTCENTERID  " & _
'               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
'               "from pah87263 a left join pah87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
'               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
'               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", "004") & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
'               "Union All " & _
'               "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, if(b.firstname is null,0,a.ded_amt2) as ded_amt2, b.firstname, b.mname, b.lastname, b.pagibigno,b.COSTCENTERID  " & _
'               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
'               "from pa87263 a left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
'               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
'               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", "004") & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
'               "order by periodid,firstname "
              
'----> remarked 201406-03
    cSqlStmt = "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, if(b.firstname is null,0,a.ded_amt2) as ded_amt2, b.firstname, b.mname, b.lastname, b.pagibigno,b.COSTCENTERID  " & _
               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
               "from pah87263 a left join pah87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", IIf(Check7.Value = vbChecked, "030", "004")) & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
               "Union All " & _
               "select a.periodid, a.empid, if(b.firstname is null,0,a.ded_amt) as ded_amt, if(b.firstname is null,0,a.ded_amt2) as ded_amt2, b.firstname, b.mname, b.lastname, b.pagibigno,b.COSTCENTERID  " & _
               IIf(Check4.Value = vbChecked, ",'' as ref_no, now() as date_grant, 0 as cut_off_amt, '' as ctrl_no ", " ,c.ref_no, c.date_grant, c.cut_off_amt, c.ctrl_no ") & _
               "from pa87263 a left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
               IIf(Check4.Value = vbChecked, "", " left join di3673 c on a.empid=c.empid and a.dedid=c.dedid and a.ctrl_no=c.ctrl_no ") & _
               "where (a.dedid='" & IIf(Check4.Value = vbChecked, "003", IIf(Check7.Value = vbChecked, "030", "004")) & "') and (a.periodid in " & cPeriod & ")" & IIf(Check4.Value = vbChecked, "", " and (trim(a.ctrl_no)<>'') ") & _
               "order by periodid,firstname "

'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        ShowProgress 0
        
        While Not oTempADO.EOF
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
                        
            OpenQueryDNS "SELECT COSTCENTERID, DESCRIPTION, COMPCODE FROM pa37722 where costcenterid = " & cQuote & oTempADO("COSTCENTERID") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aOtherinfo2(0) = objdbRs("DESCRIPTION")
                aOtherinfo2(1) = objdbRs("COMPCODE")
            Else
                aOtherinfo2(0) = ""
                aOtherinfo2(1) = ""
            End If
            
            If oTempADO("ded_amt") > 0 Then
                cSqlStmt = "select * from TmpHDMF where empid=" & cQuote & oTempADO("empid") & cQuote
                QueryTemp cSqlStmt, objdbRs, False
                If objdbRs.RecordCount = 0 Then
                    cSqlStmt = "insert into tmpHDMF(month_info,tel_num,cmp_pagibig,signatory,posname,signatory2,posname2,COSTCENTERID,DESCRIPTION,COMPCODE, " & _
                               "empid,pagibignum,firstname,mname,lastname,ded_amt,ded_amt2,tot_amt" & IIf(Check4.Value <> vbChecked, ",ref_no,date_grant", "") & ")values(" & _
                               cQuote & MonthName(cParam) & " " & Combo1.Text & cQuote & "," & _
                               cQuote & gTelNum & cQuote & "," & _
                               cQuote & gSSSNum & cQuote & "," & _
                               cQuote & Label8.Caption & cQuote & "," & _
                               cQuote & aOtherInfo(0) & cQuote & "," & _
                               cQuote & Label6.Caption & cQuote & "," & _
                               cQuote & aOtherInfo(1) & cQuote & "," & _
                               cQuote & oTempADO("COSTCENTERID") & cQuote & "," & _
                               cQuote & aOtherinfo2(0) & cQuote & "," & _
                               cQuote & aOtherinfo2(1) & cQuote & "," & _
                               cQuote & oTempADO("empid") & cQuote & "," & _
                               cQuote & oTempADO("pagibigno") & cQuote & "," & _
                               cQuote & oTempADO("firstname") & cQuote & "," & _
                               cQuote & oTempADO("mname") & cQuote & "," & _
                               cQuote & oTempADO("lastname") & cQuote & "," & _
                               oTempADO("ded_amt") & "," & _
                               IIf(Check4.Value = vbChecked, oTempADO("ded_amt2"), oTempADO("cut_off_amt")) & "," & _
                               IIf(Check4.Value = vbChecked, oTempADO("ded_amt") + oTempADO("ded_amt2"), 0) & _
                               IIf(Check4.Value <> vbChecked, "," & cQuote & oTempADO("ref_no") & cQuote & "," & cQuote & Format(oTempADO("date_grant"), "mm/dd/yyyy") & cQuote, "") & ")"
                Else
                    cSqlStmt = "update tmpHDMF set ded_amt = ded_amt + " & oTempADO("ded_amt") & "," & _
                               "ded_amt2 = ded_amt2 + " & oTempADO("ded_amt2") & ", tot_amt = tot_amt + " & oTempADO("ded_amt") & "+" & oTempADO("ded_amt2") & _
                               " where empid=" & cQuote & oTempADO("empid") & cQuote
                End If
                QueryTemp cSqlStmt, objdbRs, True
            End If
            
            oTempADO.MoveNext
        Wend
    
        ShowProgress 3
        
        
     
        
    If Check3.Value = vbUnchecked Then
    
'        GenerateReport IIf(Check4.Value = vbChecked, "MEMBERSHIP REGISTRATION/REMITTANCE FORM", ""), IIf(Check4.Value = vbChecked, "rpt4364.rpt", "rpt4364l.rpt")
        GenerateReport IIf(Check4.Value = vbChecked, "MEMBERSHIP REGISTRATION/REMITTANCE FORM", ""), IIf(Check4.Value = vbChecked, "rpt4364.rpt", IIf(Check7.Value = vbChecked, "rpt4364l_cal.rpt", "rpt4364l.rpt")) '----> remarked 201406-03
        
        
    Else
    
        GenerateReport IIf(Check4.Value = vbChecked, "MEMBERSHIP REGISTRATION/REMITTANCE FORM", ""), IIf(Check4.Value = vbChecked, "rpt4364_cost.rpt", "rpt4364l_cost.rpt")
    
    End If
    
        
        ShowProgress 4
        
    Else
        MsgBox "No data to process!", vbInformation, "System Advisory!!!"
    End If
      
    Exit Sub
    
End Sub

Function DetectTempSSS(cDBFPath As String) As Boolean
    On Error GoTo ErrDetect
    Dim oCatalog As ADOX.Catalog, _
        oTextFile As New FileSystemObject
    
    DoEvents
    
'    ' --> delete temporary file if it's existing...
'    If Dir(cTempPath & cDBFPath, vbNormal) <> "" Then
'        oTextFile.DeleteFile cDBFPath
'        Set oTextFile = Nothing
'    End If

'    Set oCatalog = New ADOX.Catalog
'    oCatalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                    "Data Source=" & cDBFPath & ";"
'    Set oCatalog = Nothing

    With oSSSConn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cDBFPath
'        .ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath(cDBFPath) & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5"
    '    oTempConn.ConnectionString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" & cTempPath
        .Open
    End With
    
'    Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder
'\myAccess2007file.accdb;Persist Security Info=False;
'g_Admentor_strConnect = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/admentor/ad2000.mdb") & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5"


    
    DetectTempSSS = True
    
    Exit Function
    
ErrDetect:
    MsgBox "Error retrieving SSS temporary file", vbCritical
    DetectTempSSS = False
End Function

Function DetectDBF(cDBFPath As String) As Boolean
    On Error GoTo ErrDetect
    Dim cString As String
    
    DoEvents

    If oDBFConn.State = adStateOpen Then oDBFConn.Close
    With oDBFConn
        .CursorLocation = adUseClient
        cString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & cDBFPath & ";" & _
                   "Extended Properties=" & cQuote & "DBASE IV;" & cQuote & ";"
'        MsgBox cString
        .ConnectionString = cString
        .Open
    End With

    DetectDBF = True

    Exit Function

ErrDetect:
    MsgBox "Error retrieving DBF file", vbCritical
    DetectDBF = False
End Function

Sub QueryDBF(ByVal cSqlStmt As String, oADORSet As ADODB.Recordset, ByVal lState As Boolean)
    On Error GoTo ErrQuery
    
    DoEvents
    If Not lState Then
        Set oADORSet = oDBFConn.Execute(cSqlStmt)
    Else
        oDBFConn.Execute (cSqlStmt)
        While oDBFConn.State = adStateExecuting
            DoEvents
        Wend
    End If
    Exit Sub
    
ErrQuery:
    ErrorMsg Err.Number, Err.Description, "Open DBF Query", "uSRM"
End Sub

Sub CreateTMPSSS()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
'    cSqlStmt = "CREATE TABLE TMPSSS( " & _
'               " [Month_INFO] char(50), [TEL_NUM] char(50)," & _
'               " [CMP_TIN] char(50),    [CMP_PH] char(50), " & _
'               " [SIGNATORY] char(50),  [POSNAME] char(50), " & _
'               " [EMPID] char(6),       [PHEALTHNUM] char(50), " & _
'               " [FIRSTNAME] CHAR(30),  [LASTNAME] char(30),   [MI] char(1), " & _
'               " [BR1] integer,          [PS1] double, " & _
'               " [ES1] double,           [REM1] char(2)," & _
'               " [DATE_EFC] date,       [TAG] integer," & _
'               " [COSTID] char(10),   [CDESC] char(100), " & _
'               " [WORKID] char(10),   [WDESC] char(100), " & _
'               " [COMPCODE] char(4))"

    cSqlStmt = "CREATE TABLE TMPSSS( " & _
               " [EMPID] char(6), " & _
               " [FIRSTNAME] CHAR(20),[MINAME] CHAR(20),[LASTNAME] CHAR(20),  " & _
               " [SSSNO] char(10), [SSSAMT] double," & _
               " [EC_AMT] double, [REM] char(1)," & _
               " [DATE_REM] integer," & _
               " [COSTID] char(10),   [CDESC] char(100), " & _
               " [WORKID] char(10),   [WDESC] char(100), " & _
               " [ER_AMT] double, [EE_AMT] double)"
               
               
'cSqlStmt = "INSERT INTO employer(ID,ernum,ername,apmo,apyr,r3file)VALUES(" & _

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TMPSSS"
    QueryTemp cSqlStmt, oTempADO, True
End Sub
' + -->
' |     Procedure Name  :   SSSPData(cParam As String)
' |     Description     :   Generate SSS Premium Remittance Report
' |     Date Created    :   8 aug 2006
' + -->
Sub SSSPData(cParam As String)
    On Error GoTo ErrLoad
    
    Dim cSqlStmt As String, oFile As File, oTextFile As New FileSystemObject, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        cParam2 As String, cCompanyName As String, cR3file As String, _
        nPClose As Integer, nCtr As Integer, _
        aEmpInfo As Variant, aERPInfo As Variant
        
    Set oSSSConn = Nothing
    
    CreateTMPSSS

    Check4.Value = vbChecked
    
    nCtr = 1
    
    aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
    
    aERPInfo = Array("", "")
        
    CommonDialog1.CancelError = False

    CommonDialog1.InitDir = CheckPath(cUploadPath) & "R3DISKETTE\Data\"
    CommonDialog1.Filter = IIf(Check4.Value = vbChecked, "SSS Premium File | r3td.mdb", "SSS Loan File | lmstrndu.dbf")
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    
    Set oFile = oTextFile.GetFile(CommonDialog1.FileName)
    If DetectTempSSS(oFile.Path) Then
    
        cSqlStmt = "delete from employer"
        QuerySSS cSqlStmt, objdbRs, True
        
        cSqlStmt = "delete from " & IIf(Check4.Value = vbChecked, "employee", "lmstrndu")
        QuerySSS cSqlStmt, objdbRs, True
    
    End If
    
    cSqlStmt = " SELECT periodid,date_start,date_end FROM pa7730 " & _
               " where (13month=0) and (month(date_start) = " & cParam & " and  month(date_end) = " & cParam & ") and " & _
               " (year(date_start) = " & cQuote & Combo1.Text & cQuote & " and year(date_end) = " & cQuote & Combo1.Text & cQuote & ")"
    OpenQueryDNS cSqlStmt, oRSet, False
    
    cParam2 = ""
    If oRSet.RecordCount > 0 Then
        While Not oRSet.EOF
            cParam2 = IIf(cParam2 = "", cQuote & oRSet("periodid"), cParam2 & cQuote & "," & cQuote & oRSet("periodid") & cQuote)
            oRSet.MoveNext
        Wend
    End If
    
    If cParam2 = "" Then GoTo ErrLoad:
    
    If Check4.Value = vbChecked Then
    
        OpenQueryDNS "select * from di2660 where cmpid =" & cQuote & gCompanyID & cQuote, objdbRs, False
        cCompanyName = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
        
        cR3file = "R3" & ListView1.SelectedItem.Text & Combo1.Text & "." & IIf(Month(Now) >= 10, Month(Now), "0" & Month(Now)) & Day(Now) & right(Year(Now), 2) & "1"

        cSqlStmt = "INSERT INTO employer(ID,ernum,ername,apmo,apyr,r3file)VALUES(" & _
                   cQuote & "01" & cQuote & "," & _
                   cQuote & gSSSNum & cQuote & "," & _
                   cQuote & UCase(left(cCompanyName, 30)) & cQuote & "," & _
                   cQuote & ListView1.SelectedItem.Text & cQuote & "," & _
                   cQuote & Combo1.Text & cQuote & "," & _
                   cQuote & cR3file & cQuote & ")"
'        MsgBox left(cCompanyName, 30)

        QuerySSS cSqlStmt, objdbRs, True
               
    End If
    
    cSqlStmt = "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215  " & _
               "FROM pah87263 a left join di3670 b on a.empid=b.empid " & _
               "  left join pah87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & IIf(Check4.Value = vbChecked, "001", "002") & cQuote & ") and ((a.ded_amt + a.ded_amt2) <> 0) " & _
               "group by a.empid " & _
               "Union All " & _
               "SELECT a.periodid,a.empid as empid, ifnull(b.firstname,c.firstname) as firstname, " & _
               "  left(ifnull(b.mname,c.mname),1) as mname, " & _
               "  ifnull(b.lastname,c.lastname) as lastname, ifnull(ssnum,'') as ssnum, " & _
               "  round(Sum(a.ded_amt + a.ded_amt2), 2) As ded_amt,c.date_hire,b.costcenterid,b.workcenterid, c.sser, c.sser1215, c.ssprem, c.ssprem1215 " & _
               "FROM pa87263 a left join di3670 b on a.empid=b.empid " & _
               "  left join pa87260 c on a.periodid=c.periodid and a.empid=c.empid " & _
               "where (a.periodid in ( " & cParam2 & " )) and (a.dedid=" & cQuote & IIf(Check4.Value = vbChecked, "001", "002") & cQuote & ") and ((a.ded_amt + a.ded_amt2) <> 0) " & _
               "group by a.empid " & _
               "order by empid, periodid "
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        
        ShowProgress 0
        
        While Not oRecordSet.EOF
        
            aEmpInfo = Array(0#, "", "", "", "", 0#, 0#, 0#)
            
            aERPInfo = Array("", "")
            
'            If oRecordSet("empid") = "JM4804" Then
'                MsgBox "stop"
'            End If

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."

            cSqlStmt = "select empid from TMPSSS where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
            QueryTemp cSqlStmt, objdbRs, False
            If Not objdbRs.RecordCount > 0 Then
'                MsgBox "insert"
                
                If Check3.Value = vbChecked Then
    
                    OpenQueryDNS "Select * FROM pa37722 where costcenterid = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                            aERPInfo(0) = objdbRs("DESCRIPTION")
                        Else
                            aERPInfo(0) = ""
                        End If
    
    
                    OpenQueryDNS "SELECT * FROM pa97722 where workcenterid = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                                aERPInfo(1) = objdbRs("DESCRIPTION")
                        Else
                                aERPInfo(1) = ""
                        End If
                Else
                    aERPInfo(0) = ""
                    aERPInfo(1) = ""
                End If

                aEmpInfo(0) = oRecordSet("empid")                                                           'ID
                aEmpInfo(1) = left(oRecordSet("ssnum"), 10)                                                 'ssnum
                aEmpInfo(2) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("lastname"), 15))))                'esurn
                aEmpInfo(3) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("firstname"), 15))))               'ename
                aEmpInfo(4) = UCase(EncodeStr2(DecodeStr(left(oRecordSet("mname"), 1))))                    'emidinit
                aEmpInfo(5) = oRecordSet("ded_amt")                                                         'ssamt
                aEmpInfo(6) = oRecordSet("sser") + oRecordSet("sser1215")
                aEmpInfo(7) = oRecordSet("ssprem") + oRecordSet("ssprem1215")
                
                cSqlStmt = "INSERT INTO employee(Id,ssnum,esurn,ename,emidinit,ssamt,ecamt,hrdate,remarks)VALUES(" & _
                           nCtr & "," & _
                           cQuote & aEmpInfo(1) & cQuote & "," & _
                           cQuote & aEmpInfo(2) & cQuote & "," & _
                           cQuote & aEmpInfo(3) & cQuote & "," & _
                           cQuote & aEmpInfo(4) & cQuote & "," & _
                           aEmpInfo(5) & "," & _
                           IIf(aEmpInfo(5) >= 1650, "30.00", IIf(aEmpInfo(5) > 0, "10.00", "0.00")) & ",0," & cQuote & "N" & cQuote & ")"
                           
                QuerySSS cSqlStmt, objdbRs, True
                
                cSqlStmt = "INSERT INTO TMPSSS(EMPID,SSSNO,LASTNAME,FIRSTNAME,MINAME,SSSAMT,EC_AMT,DATE_REM,REM,ER_AMT,EE_AMT,COSTID,CDESC,WORKID,WDESC)VALUES(" & _
                       cQuote & oRecordSet("empid") & cQuote & "," & _
                       cQuote & aEmpInfo(1) & cQuote & "," & _
                       cQuote & aEmpInfo(2) & cQuote & "," & _
                       cQuote & aEmpInfo(3) & cQuote & "," & _
                       cQuote & aEmpInfo(4) & cQuote & "," & _
                       aEmpInfo(5) & "," & _
                       IIf(aEmpInfo(5) >= 1650, "30.00", IIf(aEmpInfo(5) > 0, "10.00", "0.00")) & ",0," & cQuote & "N" & cQuote & "," & _
                       aEmpInfo(6) & "," & _
                       aEmpInfo(7) & "," & _
                       cQuote & oRecordSet("costcenterid") & cQuote & "," & _
                       cQuote & aERPInfo(0) & cQuote & "," & _
                       cQuote & oRecordSet("workcenterid") & cQuote & "," & _
                       cQuote & aERPInfo(1) & cQuote & ")"
                
'                    MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, False
                nCtr = nCtr + 1
            
            Else
                
                cSqlStmt = "update employee set " & _
                           " ssamt = ssamt + " & oRecordSet("ded_amt") & _
                           " where ssnum = " & cQuote & oRecordSet("ssnum") & cQuote
                QuerySSS cSqlStmt, objdbRs, True
                
                cSqlStmt = "update employee set " & _
                           " ecamt = iif (ssamt >= 1650, 30.00, iif(ssamt > 0, 10,0) ) " & _
                           " where ssnum = " & cQuote & oRecordSet("ssnum") & cQuote
'                MsgBox cSqlStmt
                
                QuerySSS cSqlStmt, objdbRs, True
                
                cSqlStmt = "update TMPSSS set " & _
                           " SSSAMT = SSSAMT + " & oRecordSet("ded_amt") & "," & _
                           " ER_AMT = " & oRecordSet("sser") + oRecordSet("sser1215") & "," & _
                           " EE_AMT = " & oRecordSet("ssprem") + oRecordSet("ssprem1215") & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote

'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
                
                '               ---> update 201506-03

                cSqlStmt = "update TMPSSS set " & _
                           " EC_AMT = iif (SSSAMT >= 1650 , 30.00, iif(SSSAMT > 0, 10, 0) ) " & _
                           " where empid = " & cQuote & oRecordSet("empid") & cQuote & " and SSSNO = " & cQuote & oRecordSet("ssnum") & cQuote
                           
'                MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
                
            End If
            
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 4
                
        If Check3.Value = vbChecked Then
                
            MsgBox "Saving Done... Press [OK] to continue...", vbInformation, App.Title
        
            GenerateReport "SSS R3 Data File Report", "rptSSSR3.rpt"
        
        Else
        
            MsgBox "Saving Done... Press [OK] to continue...", vbInformation, App.Title
        
        End If
    End If
    
    Set oDBFConn = Nothing
    Set oSSSConn = Nothing
    
    Exit Sub
    
ErrLoad:
    MsgBox "Error uploading SSS..."
End Sub


' + -->
' |     Procedure Name  :   GenBackupPay
' |     Description     :   Generate Backup of Processed Payroll Transaction...
' |     Date Created    :   15 aug 2006
' + -->
' + -->
' |     Procedure Name  :   GenBackupPay
' |     Description     :   Generate Backup of Processed Payroll Transaction...
' |     Date Created    :   15 aug 2006
' + -->
'Sub createBackupPay()
Sub createBackupK1Pay(nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cFieldName As String, _
        nCtr As Integer

    If nMode = 0 Then
        OpenQueryDNS "select dedid, short_desc from pa3330", objdbRs, False
        If objdbRs.RecordCount > 0 Then
            cFieldName = ""
            nCtr = 0
            
            While Not objdbRs.EOF
                If InStr("'001','006'", objdbRs("dedid")) = 0 Then
                    nCtr = nCtr + 1
                    cFieldName = cFieldName & "[" & objdbRs("short_desc") & "] double,"
                End If
                
                objdbRs.MoveNext
            Wend
        End If
        
        cSqlStmt = "CREATE TABLE K1PAY (" & _
                   " [EMPID] char(6),                [BACCNTNO] char(16),            [FULLNAME] char(60),            [POSITION] char(60), " & _
                   " [DEPID] char(3),                [DEPARTMENT] char(60),          " & _
                   " [EMP_STAT] char(1),             [WAP] char(10),                  [DATE_HIRE] char(10),           [STATUS] char(2)," & _
                   " [date_res] char(10),            [TAXCODE] char(5),              [RATE_AMT] NUMERIC(18,4),       [POS_ALLOW] NUMERIC(18,4)," & _
                   " [COLA_AMT] double,              [REG_DAY] double,               [REG_PAY] double,               [REG_OT_HR] NUMERIC(18,4)," & _
                   " [REG_OT_PAY] NUMERIC(18,4),     [ND_DAY] NUMERIC(18,4),         [ND_PAY] NUMERIC(18,4),         [ND_OT_HR] NUMERIC(18,4)," & _
                   " [ND_OT_PAY] NUMERIC(18,4),      [HOLIDAY] NUMERIC(18,4),        [HOL_PAY] NUMERIC(18,4),        [SA_REG_OT] NUMERIC(18,4)," & _
                   " [SA_REG_PAY] NUMERIC(18,4),     [SA_ND_OT] NUMERIC(18,4),       [SA_ND_PAY] NUMERIC(18,4),      " & _
                   " [SUN_HR] NUMERIC(18,4),         [SUN_PAY] NUMERIC(18,4),        [SUN_OT] NUMERIC(18,4),         [SUN_OT_PAY] NUMERIC(18,4),     [SUN_COLA] NUMERIC(18,4)," & _
                   " [SUN_ND] NUMERIC(18,4),         [SUN_ND_PAY] NUMERIC(18,4),     [SUN_ND_OT] NUMERIC(18,4),      [SUNNDOTPAY] NUMERIC(18,4),     " & _
                   " [ADJ_PAY] NUMERIC(18,4),        [SA_ADJ_PAY] NUMERIC(18,4),     [OTHER_PAY] NUMERIC(18,4),      [LEAVE_PAY] NUMERIC(18,4)," & _
                   " [M13PAY] NUMERIC(18,4),         [WTAX] NUMERIC(18,4),           [TAXABLE] NUMERIC(18,4),        [SSPREM] NUMERIC(18,4)," & _
                   cFieldName & " [DED_AMT] NUMERIC(18,4)," & _
                   " [GROSS_PAY] NUMERIC(18,4),      [NET_PAY] NUMERIC(18,4),        [SA_NET_PAY] NUMERIC(18,4),     [GROSS16231] NUMERIC(18,4)," & _
                   " [SSER] NUMERIC(18,4),           [SSS01] NUMERIC(18,4),          [EC001] NUMERIC(18,4),          [SSER1215] NUMERIC(18,4)," & _
                   " [SSPREM1215] NUMERIC(18,4),     [MEDICARE2] NUMERIC(18,4),      [MED01] NUMERIC(18,4),          [PS1215] NUMERIC(18,4)," & _
                   " [ES1215] NUMERIC(18,4),         [SSSNUM] char(15),              [TINNUM] char(15),              [PAYSTATUS] NUMERIC(1,0)," & _
                   " [PAGIBIGNO] char(15),           [FIRSTNAME] char(60),           [MNAME] char(60),               [LASTNAME] char(60)," & _
                   " [PHEALTHNUM] char(15),          [BASIC1215] NUMERIC(18,4),      [BASICPAY] NUMERIC(18,4),       [SEQ_NO] NUMERIC(1,0), [INC_HR] NUMERIC(18,4),         [INC_PAY] NUMERIC(18,4), [BIRTHDAY] char(60), [COSTCENTERID] char(10)  ) "
                
    Else
        cSqlStmt = " CREATE TABLE MASTERK1 (" & _
                   " [EMPID] char(6),        [BACCNTNO] char(16),        [LASTNAME] char(50),    [MNAME] char(50),       [FIRSTNAME] char(50), " & _
                   " [BIRTHDAY] char(60),        [SEX] char(1),              [PAGIBIGNO] char(20),   [DEPID] char(3), " & _
                   " [LineName] char(50),    [POSID] char(3),            [POSNAME] char(50),     [POS_ALLOW] double, " & _
                   " [RATE_AMT] double,      [MTD_BASIC] double,         [MTD_GROSS] double,     [YTD_BASIC] double, " & _
                   " [YTD_GROSS] double,     [YTD_WTAX] double,          [EMP_STAT] char(1),     [PAYSTATUS] char(1), " & _
                   " [SSNUM] char(15),       [ISUNION] char(1),          [TIN] char(15),         [TAXID] char(3), " & _
                   " [TAXCODE] char(15),     [TAXNAME] char(100),        [FULLNAME] char(100),   [DATE_HIRE] date, " & _
                   " [SL_AVAIL] double,      [VL_AVAIL] double,          [SL_USE] double,        [VL_USE] double, " & _
                   " [SSPREM1215] double,    [SSER1215] double,          [PS1215] double,        [ES1215] double, " & _
                   " [DATE_RES] date,        [STATUS] integer,           [ACTIVE] char(2),       [WAP] char(10), " & _
                   " [CMPID] char(4),        [CMPName] char(50), [COSTCENTERID] char(10))"
    
    End If
    
    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    If nMode = 0 Then
        cSqlStmt = "DELETE FROM K1PAY"
    Else
        cSqlStmt = "DELETE FROM MASTERK1"
    End If
    QueryDBF cSqlStmt, oTempADO, True
End Sub

Sub GenbackupK1pay(cPath As String, ByVal cPeriod As String)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cParam As String, _
        cString, _
        cFieldName As String, _
        cBAccntNo As String, _
        cBDay As String, _
        nCtr As Integer, _
        cPclose As Integer

    DetectDBF cPath
    
    createBackupK1Pay 0
    createBackupK1Pay 1
    
    ShowProgress 0
    
    OpenQueryDNS "select pclose from pa7730 where periodid = " & cQuote & cPeriod & cQuote, objdbRs, False
    cPclose = IIf(objdbRs.RecordCount > 0, objdbRs("pclose"), 0)
    
    OpenQueryDNS "select dedid, short_desc from pa3330", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cParam = ""
        cFieldName = ""
        nCtr = 0
        
        While Not objdbRs.EOF
            If InStr("'001','006'", objdbRs("dedid")) = 0 Then
                nCtr = nCtr + 1
                cString = Chr$(100 + nCtr)
                cParam = cParam & _
                         " left join " & IIf(cPclose = 1, " pah87263 ", " pa87263 ") & cString & " on a.periodid=" & cString & ".periodid and a.empid=" & cString & ".empid and " & cString & ".dedid=" & cQuote & objdbRs("dedid") & cQuote & vbCrLf
                cFieldName = cFieldName & "ifnull(" & cString & ".ded_amt,0) as " & objdbRs("short_desc") & ","
            End If
            
            objdbRs.MoveNext
        Wend
    End If

    cSqlStmt = " select a.EMPID, a.FULLNAME,  ifnull(d.posname,'') as position, a.depid, ifnull(b.linename,'') as department," & _
           " a.paystatus, if(a.EMP_STAT=0,'W',if(a.EMP_STAT=1,'C','R')) as emp_stat, if(a.WAP=1,'WAP-C','') as WAP, " & _
           " a.DATE_HIRE, if(a.ACTIVE>0,if(a.ACTIVE=1,'R',if(a.ACTIVE=3,'T','FC')),'') as status, if(a.ACTIVE>0,a.date_res,'') as date_res, " & _
           " ifnull(c.taxcode,'') as taxcode, a.RATE_AMT, a.POS_ALLOW, a.COLA, a.SUN_COLA, " & _
           " a.REG_DAY, a.REG_PAY, a.REG_OT_HR, a.REG_OT_PAY, a.NDIFF_DAY, a.NDIFF_PAY, a.NDIFF_OT_HR, a.NDIFF_OT_PAY, a.HOLIDAY, a.HOL_PAY, a.SA_REG_OT, a.SA_REG_PAY, a.SA_NDIFF_OT, a.SA_NDIFF_PAY, " & _
           " a.SUN_HR, a.SUN_PAY, a.SUN_OT, a.SUN_OT_PAY, a.SUN_ND, a.SUN_ND_PAY, a.SUN_ND_OT, a.SUN_ND_OT_PAY, " & _
           " a.ADJ_PAY, a.SA_ADJ_PAY, a.OTHER_PAY, a.LEAVE_PAY, a.M13PAY, a.WTAX, a.TAXABLE, a.SSPREM, " & _
           cFieldName & _
           " a.DED_AMT, a.GROSS_PAY, a.NET_PAY, a.SA_NET_PAY, a.GROSS16231, a.SSER, a.SSS01, a.EC001, a.SSER1215, a.SSPREM1215, a.MEDICARE2, a.MED01, a.PS1215, a.ES1215, " & _
           " a.SSSNUM , a.TINNUM, a.PAGIBIGNO, a.FIRSTNAME, a.MNAME, a.LASTNAME, a.PHEALTHNUM, a.BASIC1215, " & _
           " a.basicpay,a.seq_no,a.INC_HR, a.INC_PAY, a.COSTCENTERID " & _
           " from " & IIf(cPclose = 1, " pah87260 a ", " pa87260 a ") & _
           cParam & _
           " left join di5463 b on a.depid=b.lineid left join pa8290 c on a.taxid=c.taxid left join di7670 d on a.posid=d.posid " & _
           " where a.periodid= " & cQuote & cPeriod & cQuote & _
           " order by status, a.depid, a.emp_stat desc, a.lastname, a.firstname "
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF

            OpenQueryDNS "select dedid, short_desc from pa3330", objdbRs, False
            If objdbRs.RecordCount > 0 Then
                cParam = ""
                cFieldName = ""
                nCtr = 0
                
                While Not objdbRs.EOF
                    If InStr("'001','006'", objdbRs("dedid")) = 0 Then
                        nCtr = nCtr + 1
                        cString = Chr$(100 + nCtr)
                        cParam = cParam & "[" & objdbRs("short_desc") & "],"
                        cFieldName = cFieldName & oRecordSet(Trim(objdbRs("short_desc"))) & ","
                    End If
                    
                    objdbRs.MoveNext
                Wend
            End If
                
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            
            '2009-08-12 ATM additional for backup k1pay and master K1
            OpenQueryDNS " select BACCNTNO,birthday from di3670 where empid = " & cQuote & oRecordSet("empid") & cQuote, objdbRs, False
            cBAccntNo = IIf(objdbRs.RecordCount > 0, objdbRs("BACCNTNO"), "")
            cBDay = IIf(objdbRs.RecordCount > 0, objdbRs("birthday"), "")
            

            cSqlStmt = " INSERT INTO K1pay (EMPID,BACCNTNO,FULLNAME,[POSITION],DEPID,DEPARTMENT,EMP_STAT,WAP,DATE_HIRE,[STATUS],date_res,TAXCODE,RATE_AMT,POS_ALLOW,COLA_AMT,SUN_COLA,REG_DAY,REG_PAY,REG_OT_HR,REG_OT_PAY,ND_DAY,ND_PAY,ND_OT_HR," & _
                       " ND_OT_PAY,[HOLIDAY],HOL_PAY,SA_REG_OT,SA_REG_PAY,SA_ND_OT,SA_ND_PAY,SUN_HR,SUN_PAY,SUN_OT,SUN_OT_PAY,SUN_ND,SUN_ND_PAY,SUN_ND_OT,SUNNDOTPAY," & _
                       " ADJ_PAY,SA_ADJ_PAY,OTHER_PAY,LEAVE_PAY,M13PAY,WTAX,TAXABLE,SSPREM," & _
                       cParam & _
                       " DED_AMT,GROSS_PAY,NET_PAY,SA_NET_PAY,GROSS16231,SSER,SSS01,EC001,SSER1215,SSPREM1215,MEDICARE2,MED01,PS1215,ES1215,SSSNUM,TINNUM,PAYSTATUS,PAGIBIGNO,FIRSTNAME,MNAME,LASTNAME,PHEALTHNUM,BASIC1215,BASICPAY,SEQ_NO,INC_HR,INC_PAY,BIRTHDAY,COSTCENTERID)VALUES(" & _
                       cQuote & left(oRecordSet("EMPID"), 6) & cQuote & "," & cQuote & Trim(cBAccntNo) & cQuote & "," & cQuote & EncodeStr2(DecodeStr(left(oRecordSet("FULLNAME"), 100))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(left(oRecordSet("POSITION"), 100))) & cQuote & "," & _
                       cQuote & oRecordSet("DEPID") & cQuote & "," & cQuote & EncodeStr2(DecodeStr(left(oRecordSet("DEPARTMENT"), 100))) & cQuote & "," & _
                       cQuote & left(oRecordSet("EMP_STAT"), 1) & cQuote & "," & cQuote & left(oRecordSet("WAP"), 5) & cQuote & "," & _
                       cQuote & left(oRecordSet("DATE_HIRE"), 10) & cQuote & "," & _
                       cQuote & left(oRecordSet("STATUS"), 2) & cQuote & "," & _
                       cQuote & left(oRecordSet("date_res"), 10) & cQuote & "," & _
                       cQuote & left(oRecordSet("TAXCODE"), 5) & cQuote & "," & _
                       oRecordSet("RATE_AMT") & "," & oRecordSet("POS_ALLOW") & "," & oRecordSet("COLA") & "," & oRecordSet("SUN_COLA") & "," & oRecordSet("REG_DAY") & "," & oRecordSet("REG_PAY") & "," & oRecordSet("REG_OT_HR") & "," & _
                       oRecordSet("REG_OT_PAY") & "," & oRecordSet("NDIFF_DAY") & "," & oRecordSet("NDIFF_PAY") & "," & oRecordSet("NDIFF_OT_HR") & "," & oRecordSet("NDIFF_OT_PAY") & "," & oRecordSet("HOLIDAY") & "," & oRecordSet("HOL_PAY") & "," & _
                       oRecordSet("SA_REG_OT") & "," & oRecordSet("SA_REG_PAY") & "," & oRecordSet("SA_NDIFF_OT") & "," & oRecordSet("SA_NDIFF_PAY") & "," & oRecordSet("SUN_HR") & "," & oRecordSet("SUN_PAY") & "," & oRecordSet("SUN_OT") & "," & oRecordSet("SUN_OT_PAY") & "," & _
                       oRecordSet("SUN_ND") & "," & oRecordSet("SUN_ND_PAY") & "," & oRecordSet("SUN_ND_OT") & "," & oRecordSet("SUN_ND_OT_PAY") & "," & _
                       oRecordSet("ADJ_PAY") & "," & oRecordSet("SA_ADJ_PAY") & "," & oRecordSet("OTHER_PAY") & "," & oRecordSet("LEAVE_PAY") & "," & oRecordSet("M13PAY") & "," & oRecordSet("WTAX") & "," & oRecordSet("TAXABLE") & "," & oRecordSet("SSPREM") & "," & _
                       cFieldName & oRecordSet("DED_AMT") & "," & oRecordSet("GROSS_PAY") & "," & oRecordSet("NET_PAY") & "," & oRecordSet("SA_NET_PAY") & "," & _
                       oRecordSet("GROSS16231") & "," & oRecordSet("SSER") & "," & oRecordSet("SSS01") & "," & oRecordSet("EC001") & "," & oRecordSet("SSER1215") & "," & oRecordSet("SSPREM1215") & "," & oRecordSet("MEDICARE2") & "," & oRecordSet("MED01") & "," & oRecordSet("PS1215") & "," & oRecordSet("ES1215") & "," & _
                       cQuote & left(oRecordSet("SSSNUM"), 15) & cQuote & "," & cQuote & left(oRecordSet("TINNUM"), 15) & cQuote & "," & oRecordSet("PAYSTATUS") & "," & cQuote & left(oRecordSet("PAGIBIGNO"), 15) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(left(oRecordSet("FIRSTNAME"), 100))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(left(oRecordSet("MNAME"), 100))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(left(oRecordSet("LASTNAME"), 100))) & cQuote & "," & cQuote & left(oRecordSet("PHEALTHNUM"), 15) & cQuote & "," & oRecordSet("BASIC1215") & "," & _
                       oRecordSet("BASICPAY") & "," & oRecordSet("seq_no") & "," & oRecordSet("INC_HR") & "," & oRecordSet("INC_PAY") & "," & cQuote & left(cBDay, 10) & cQuote & "," & cQuote & oRecordSet("COSTCENTERID") & cQuote & ")"
'            MsgBox cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            oRecordSet.MoveNext

        Wend

        ShowProgress 4
    Else
        ShowProgress 4
        MsgBox "Data not found...!!!", vbInformation, App.Title
        Exit Sub
    End If

    ShowProgress 0
    
    cSqlStmt = " SELECT a.date_hire,if(a.active>0,if(a.active=1,a.date_res,a.date_fin),'') as date_res,a.status,a.rate_amt,a.empid,a.BACCNTNO,a.firstname,a.mname,a.lastname," & _
               " a.cmpid,b.cmpname,a.depid,c.linename,ifnull(a.ssnum,'') as ssnum,a.pagibigno,a.tin,a.taxcode,d.taxname,a.birthday,if(a.emp_stat=0,'W',if(a.emp_stat=1,'C','R')) as emp_stat, " & _
               " if(a.isunion=0,'Y','N') as isUnion,if(a.sex=0,'M','F') as sex,a.posid,e.posname,a.ytd_gross,a.ytd_basic,a.ytd_wtax,a.sl_avail,a.vl_avail,a.sl_use,a.vl_use," & _
               " a.pos_allow,a.date_res,a.ytd_cola,a.ytd_gross_sa,a.mtd_basic,a.mtd_gross,a.es1215,a.ps1215,a.ssprem1215,a.sser1215,if(a.paystatus=0,'D','M') as paystatus," & _
               " if(a.active>0,if(a.active=1,'R','FC'),'') as active,if(a.wap=1,'WAP-C','') as wap,a.taxid,concat(a.lastname,', ', a.firstname,' ', if(trim(a.mname)='','',concat(left(a.mname,1),'.'))) as fullname, a.COSTCENTERID FROM di3670 a " & _
               " left join di2660 b on a.cmpid=b.cmpid left join di5463 c on a.depid=c.lineid left join pa8290 d on a.taxcode=d.taxcode left join DI7670 e on a.posid=e.posid "
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            
            cSqlStmt = " INSERT INTO MASTERK1 (EMPID,BACCNTNO,LASTNAME,MNAME,FIRSTNAME,BIRTHDAY,SEX,PAGIBIGNO,DEPID,LineName,POSID,POSNAME,POS_ALLOW,RATE_AMT,MTD_BASIC,MTD_GROSS,YTD_BASIC,YTD_GROSS,YTD_WTAX,EMP_STAT,PAYSTATUS," & _
                       " SSNUM,ISUNION,TIN,TAXID,TAXCODE,TAXNAME,FULLNAME,DATE_HIRE,SL_AVAIL,VL_AVAIL,SL_USE,VL_USE,SSPREM1215,SSER1215,PS1215,ES1215,DATE_RES,STATUS,ACTIVE,WAP,CMPID,CMPName,COSTCENTERID)VALUES(" & _
                       cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("BACCNTNO") & cQuote & "," & cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("birthday") & cQuote & "," & cQuote & oRecordSet("sex") & cQuote & "," & _
                       cQuote & oRecordSet("pagibigno") & cQuote & "," & cQuote & oRecordSet("depid") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & cQuote & oRecordSet("posid") & cQuote & "," & cQuote & oRecordSet("posname") & cQuote & "," & oRecordSet("pos_allow") & "," & _
                       oRecordSet("rate_amt") & "," & oRecordSet("mtd_basic") & "," & oRecordSet("mtd_gross") & "," & oRecordSet("ytd_basic") & "," & oRecordSet("ytd_gross") & "," & oRecordSet("ytd_wtax") & "," & cQuote & oRecordSet("emp_stat") & cQuote & "," & cQuote & oRecordSet("paystatus") & cQuote & "," & _
                       cQuote & oRecordSet("ssnum") & cQuote & "," & cQuote & oRecordSet("isunion") & cQuote & "," & cQuote & oRecordSet("tin") & cQuote & "," & cQuote & oRecordSet("taxid") & cQuote & "," & cQuote & oRecordSet("taxcode") & cQuote & "," & cQuote & oRecordSet("taxname") & cQuote & "," & cQuote & EncodeStr2(DecodeStr(oRecordSet("fullname"))) & cQuote & "," & _
                       cQuote & oRecordSet("DATE_HIRE") & cQuote & "," & oRecordSet("sl_avail") & "," & oRecordSet("vl_avail") & "," & oRecordSet("sl_use") & "," & oRecordSet("vl_use") & "," & oRecordSet("ssprem1215") & "," & oRecordSet("sser1215") & "," & oRecordSet("ps1215") & "," & _
                       oRecordSet("es1215") & "," & cQuote & oRecordSet("date_res") & cQuote & "," & cQuote & oRecordSet("status") & cQuote & "," & cQuote & oRecordSet("active") & cQuote & "," & cQuote & oRecordSet("wap") & cQuote & "," & cQuote & oRecordSet("cmpid") & cQuote & "," & _
                       cQuote & oRecordSet("cmpname") & cQuote & "," & cQuote & oRecordSet("COSTCENTERID") & cQuote & ")"
                                   
'            MsgBox cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            oRecordSet.MoveNext

        Wend

        ShowProgress 4

        MsgBox "Backup File Done... Press [OK] to continue...", vbInformation, App.Title
    Else
        MsgBox "Data not found...!!!", vbInformation, App.Title
    End If
    
    Set oRecordSet = Nothing
    
End Sub

Sub checkK1Path(cPathcheck As String, cDuration As String)
    Dim cWant As String, _
        cHistoryPath As String, _
        FileSys As FileSystemObject, _
        nCtr As Integer
        
    Set FileSys = New FileSystemObject
    
    cHistoryPath = ""
                    
    cHistoryPath = CheckPath(Text4.Text) & "history"
    
    cHistoryPath = CheckPath(cHistoryPath)
    
    If Dir(cHistoryPath, vbDirectory) = "" Then MkDir cHistoryPath

    cHistoryPath = cHistoryPath & cDuration & "\"
    
    If Dir(cHistoryPath, vbDirectory) = "" Then MkDir cHistoryPath
    
    cHistoryPath = CheckPath(cHistoryPath)

Loop1:
    If Dir(cHistoryPath) = "" Then
        FileSys.CopyFile cPathcheck & "K1pay.dbf", cHistoryPath & "K1pay.dbf"
        FileSys.CopyFile cPathcheck & "MasterK1.dbf", cHistoryPath & "MasterK1.dbf"
    Else
        If (FileSys.FileExists(cHistoryPath & "K1pay" & IIf(nCtr <> 0, nCtr, "") & ".dbf") = True) Or (FileSys.FileExists(cHistoryPath & "MasterK1pay" & IIf(nCtr <> 0, nCtr, "") & ".dbf") = True) Then
            nCtr = nCtr + 1
            GoTo Loop1
        Else
            FileSys.CopyFile cPathcheck & "K1pay.dbf", cHistoryPath & "K1pay" & nCtr & ".dbf"
            FileSys.CopyFile cPathcheck & "MasterK1.dbf", cHistoryPath & "MasterK1" & nCtr & ".dbf"
        End If
        
    End If
End Sub

Sub GenBackupPay(cParam As String)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cWant As String, _
        cPayPath As String, _
        cDuration As String, _
        cDurationPath As String, _
        cHistoryPath As String, _
        FileSys As FileSystemObject, _
        nCtr As Integer
        
    Set FileSys = New FileSystemObject
    
    cSqlStmt = " SELECT duration FROM pa7730 " & _
               " Where periodid = " & cQuote & cParam & cQuote
        
    OpenQueryDNS cSqlStmt, oRecordSet, False
    cDuration = IIf(oRecordSet.RecordCount > 0, oRecordSet("duration"), "")
    
    cDurationPath = CheckPath(Text4.Text) & "Backup"
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath
    
    cDurationPath = cDurationPath & "\" & cDuration
    cDurationPath = CheckPath(cDurationPath)
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath
    
    
    If (FileSys.FileExists(cDurationPath & "K1Pay.dbf") = True) Or (FileSys.FileExists(cDurationPath & "MasterK1.dbf") = True) Then
        cWant = MsgBox("K1Pay.dbf / MasterK1.dbf Already Exsist... do you want to backup previously created file ?", vbYesNoCancel + vbCritical, App.Title)
        If cWant = vbYes Then
            checkK1Path cDurationPath, cDuration
        End If
        
        FileSys.DeleteFile cDurationPath & "K1Pay.dbf"
        FileSys.DeleteFile cDurationPath & "MasterK1.dbf"
        
    End If
    
    GenbackupK1pay cDurationPath, cParam
    
End Sub



' + -->
' |     Procedure Name  :   GenSalDiv(ByVal cPeriod As String)
' |     Description     :   Generate Salary Division Report
' |     Date Created    :   22 aug 2006
' + -->
Sub CreateTmpSalDiv(ByVal nPeriod As Integer)
    On Error GoTo ErrCreate
    Dim cParam, _
        cSqlStmt As String, _
        nFieldCnt As Integer

    nFieldCnt = 0
    cParam = "fldname0 char(50), fldvalue0 double, "
    
    cSqlStmt = "select dedid, dedname from pa3330 where period" & nPeriod & "=1 order by dedid"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        nFieldCnt = objdbRs.RecordCount
        While Not objdbRs.EOF
            cParam = cParam & _
                     "fldname" & objdbRs.AbsolutePosition & " char(50)," & _
                     "fldvalue" & objdbRs.AbsolutePosition & " double,"
            objdbRs.MoveNext
        Wend
    End If
    
    cParam = cParam & _
             "fldname" & (nFieldCnt + 1) & " char(50)," & _
             "fldvalue" & (nFieldCnt + 1) & " double," & _
             "fldname" & (nFieldCnt + 2) & " char(50)," & _
             "fldvalue" & (nFieldCnt + 2) & " double," & _
             "fldname" & (nFieldCnt + 3) & " char(50)," & _
             "fldvalue" & (nFieldCnt + 3) & " double," & _
             "fldname" & (nFieldCnt + 4) & " char(50)," & _
             "fldvalue" & (nFieldCnt + 4) & " double," & _
             "fldname" & (nFieldCnt + 5) & " char(50)," & _
             "fldvalue" & (nFieldCnt + 5) & " double,"

    cSqlStmt = "CREATE TABLE TmpSalDiv( " & _
               " [periodname] char(100),    [tag] integer, " & _
               " [days_work] double,        [holiday] double," & _
               " [depid] char(3),           [depname] char(50), " & _
               cParam & " [mp_reg] double,  [mp_wap] double)"
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TmpSalDiv"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Private Sub GenSalDiv(ByVal cPeriod As String, ByVal ntag2 As Integer)
    Dim oRecordSet As New ADODB.Recordset, _
        nFieldCnt, ntag, nCtr, nPeriod As Integer, _
        cPeriodName, cParam, cParam2, cSqlStmt, cString As String, _
        aFieldDesc As Variant
       
    aFieldDesc = Array("", "", "", "")
       
    cSqlStmt = "select date_start, date_end, status from pa7730 where periodid=" & cQuote & cPeriod & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        nPeriod = IIf(objdbRs.RecordCount > 0, objdbRs("status"), 0)
        cPeriodName = "For the Payroll period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
    End If
    
    cParam = ""
    cParam2 = ""
    
    cSqlStmt = "select dedid, dedname from pa3330 where period" & nPeriod & "=1 order by dedid"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        nFieldCnt = objdbRs.RecordCount
    
        ReDim aFieldDesc(nFieldCnt + 6) '+ 5)
        
        aFieldDesc(0) = "GROSS PAY"
    
        While Not objdbRs.EOF
            cString = Chr$(98 + objdbRs.AbsolutePosition)
            cParam = cParam & " sum(ifnull(" & cString & ".ded_amt,0)) as dedamt" & objdbRs.AbsolutePosition & ","
            cParam2 = cParam2 & " left join pa87263 " & cString & _
                      " on a.periodid=" & cString & ".periodid" & _
                      " and a.empid=" & cString & ".empid " & _
                      " and " & cString & ".dedid=" & cQuote & objdbRs("dedid") & cQuote
            aFieldDesc(objdbRs.AbsolutePosition) = objdbRs("dedname")
            objdbRs.MoveNext
        Wend
    End If
    
    aFieldDesc(nFieldCnt + 1) = "Deductions"
    aFieldDesc(nFieldCnt + 2) = "NET PAY"
    aFieldDesc(nFieldCnt + 3) = "SA"
    aFieldDesc(nFieldCnt + 4) = "WAP"
    aFieldDesc(nFieldCnt + 5) = "WAP-SA"
       
    CreateTmpSalDiv nPeriod
    
    ShowProgress 0
    
'    cSqlStmt = "select a.p_day, a.p_holiday, a.depid, ifnull(b.linename,'') as department, b.production, " & _
'               "  sum(if(a.emp_stat>0,1,0)) as mp_reg, " & _
'               "  sum(if(a.emp_stat=0,1,0)) as mp_wap, " & _
'               "  sum(if(a.emp_stat>0,a.gross_pay,0)) as gross_pay, " & _
'               "  sum(a.ded_amt) as deduction, " & _
'               cParam & _
'               "  sum(if(a.emp_stat>0,a.net_pay,0)) as net_reg, " & _
'               "  sum(if(a.emp_stat=0,a.net_pay,0)) as net_wap, " & _
'               "  sum(if(a.emp_stat>0,a.sa_net_pay,0)) as sa_net_reg, " & _
'               "  sum(if(a.emp_stat=0,a.sa_net_pay,0)) as sa_net_wap " & _
'               "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
'               cParam2 & _
'               " where (a.periodid=" & cQuote & cPeriod & cQuote & ")" & _
'               " and (a.paystatus" & IIf(Combo1.ListIndex = 0, "<>", "=") & "2)" & _
'               " and (a.active=0)" & _
'               " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
'               " group by a.depid"
               
    cSqlStmt = "select a.p_day, a.p_holiday, a.depid, ifnull(b.linename,'') as department, b.production, " & _
               "  sum(if(a.emp_stat>0,1,0)) as mp_reg, " & _
               "  sum(if(a.emp_stat=0,1,0)) as mp_wap, " & _
               "  sum(if(a.emp_stat>0,a.gross_pay,0)) as gross_pay, " & _
               "  sum(a.ded_amt) as deduction, " & _
               cParam & _
               "  sum(if(a.emp_stat>0,a.net_pay,0)) as net_reg, " & _
               "  sum(if(a.emp_stat=0,a.net_pay,0)) as net_wap, " & _
               "  sum(if(a.emp_stat>0,a.sa_net_pay,0)) as sa_net_reg, " & _
               "  sum(if(a.emp_stat=0,a.sa_net_pay,0)) as sa_net_wap " & _
               "from pa87260 a left join di5463 b on a.depid=b.lineid " & _
               cParam2 & _
               " where (a.periodid=" & cQuote & cPeriod & cQuote & ")" & _
               " and (a.paystatus" & IIf(Combo1.ListIndex = 0, "<>", "=") & "2)" & _
               " and (a.active" & IIf(ntag2 = 46, "<>", "=") & "0)" & _
               IIf(Combo1.ListIndex <> 0, "", " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")") & _
               " group by a.depid"
               
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            If oRecordSet("depid") = gAdmin Then
                ntag = 0
            ElseIf oRecordSet("depid") = gStaff Then
                ntag = 1
            ElseIf oRecordSet("production") = 1 Then
                ntag = 2
            Else
                ntag = 3
            End If
            
            cParam = "fldname0, fldvalue0, "
            cParam2 = cQuote & aFieldDesc(0) & cQuote & "," & oRecordSet("gross_pay") & ","
            
            For nCtr = 1 To nFieldCnt
                cParam = cParam & _
                         "fldname" & nCtr & ", " & _
                         "fldvalue" & nCtr & ", "
                cParam2 = cParam2 & _
                          cQuote & aFieldDesc(nCtr) & cQuote & ", " & _
                          oRecordSet("dedamt" & nCtr) & ", "
            Next nCtr
            
            cParam = cParam & _
                     "fldname" & (nFieldCnt + 1) & ", " & _
                     "fldvalue" & (nFieldCnt + 1) & ", " & _
                     "fldname" & (nFieldCnt + 2) & ", " & _
                     "fldvalue" & (nFieldCnt + 2) & ", " & _
                     "fldname" & (nFieldCnt + 3) & ", " & _
                     "fldvalue" & (nFieldCnt + 3) & ", " & _
                     "fldname" & (nFieldCnt + 4) & ", " & _
                     "fldvalue" & (nFieldCnt + 4) & ", " & _
                     "fldname" & (nFieldCnt + 5) & ", " & _
                     "fldvalue" & (nFieldCnt + 5) & ", "
            cParam2 = cParam2 & _
                      cQuote & aFieldDesc(nFieldCnt + 1) & cQuote & ", " & _
                      oRecordSet("deduction") & ", " & _
                      cQuote & aFieldDesc(nFieldCnt + 2) & cQuote & ", " & _
                      oRecordSet("net_reg") & ", " & _
                      cQuote & aFieldDesc(nFieldCnt + 3) & cQuote & ", " & _
                      oRecordSet("sa_net_reg") & ", " & _
                      cQuote & aFieldDesc(nFieldCnt + 4) & cQuote & ", " & _
                      oRecordSet("net_wap") & ", " & _
                      cQuote & aFieldDesc(nFieldCnt + 5) & cQuote & ", " & _
                      oRecordSet("sa_net_wap") & ", "
            
            cSqlStmt = "insert into tmpsaldiv(periodname, [tag], [days_work], [holiday], depid, depname, " & _
                       cParam & " mp_reg, mp_wap)values(" & _
                       cQuote & cPeriodName & cQuote & "," & _
                       ntag & "," & _
                       oRecordSet("p_day") & "," & oRecordSet("p_holiday") & "," & _
                       cQuote & oRecordSet("depid") & cQuote & "," & cQuote & EncodeStr2(oRecordSet("department")) & cQuote & "," & _
                       cParam2 & _
                       oRecordSet("mp_reg") & "," & oRecordSet("mp_wap") & ")"
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
        
        Wend
        
        ShowProgress 3

        GenerateReport IIf(Combo1.ListIndex = 0, IIf(Check6.Value = vbChecked, "NO ATM ", ""), "") & IIf(ntag2 = 46, "Resigend/FC ", "") & "Salary Division", IIf(Check4.Value = vbChecked, "rpt725348.rpt", "rpt725348s.rpt")
        
    Else
        ShowProgress 3
        MsgBox "No Report to Generate!!!", vbCritical, "System Advisory"
    End If
    
    ShowProgress 4
    
    Set oRecordSet = Nothing
End Sub


' + -->
' |     Procedure Name  :   Gen13mobackup
' |     Description     :   Generate 13 Month Pay backup file
' |     Date Created    :   xx dec 2006
' + -->
Sub CreateTmpM13(ByVal nMode As Integer)
    Dim nCtr As Integer, _
        cSqlStmt As String, _
        aField As Variant
        
    aField = Array("", "", "")
        
    For nCtr = 1 To 24
        aField(0) = aField(0) & "[bas" & Format(nCtr, "00") & "] double,"
        aField(1) = aField(1) & "[gr" & Format(nCtr, "00") & "] double,"
        aField(2) = aField(2) & "[sa" & Format(nCtr, "00") & "] double,"
    Next nCtr
    
    
'    ' --> for MICO to be removed by 2008...
'    If gCompanyID = "0003" Then
'        aField(1) = ""
'        For nCtr = 1 To 8
'            aField(1) = aField(1) & "[bas" & Format(nCtr, "00") & "] double,"
'        Next nCtr
'        For nCtr = 9 To 24
'            aField(1) = aField(1) & "[gr" & Format(nCtr, "00") & "] double,"
'        Next nCtr
'    End If
'    ' --> end of MICO...

    If nMode = 0 Then
        cSqlStmt = "create table C13MOPAY (" & _
                   "[BACCNTNO] char(16), [empid] char(6), [date_hire] date,[date_fin] date,[date_res] date, [fullname] char(100), [emp_stat] char(1), " & _
                   aField(0) & _
                   "[totbasic] double, [leave_pay] double, [m13pay] double, " & _
                   "[CA] double, [MCA] double,  [gross13] double, [cash_adv] double, [net13] double, " & _
                   "[COSTCENTERID] char(10), [WORKCENTERID] char(10) )"
                   
                   
    Else
        cSqlStmt = "create table R13MOPAY (" & _
                   "[BACCNTNO] char(16), [empid] char(6), [date_hire] date,[date_fin] date,[date_res] date,[fullname] char(100), [emp_stat] char(1), [tin] char(20), [taxcode] char(10), " & _
                   aField(1) & _
                   aField(2) & _
                   "[totgross] double, [totsa] double, [leave_pay] double, [m13pay] double, " & _
                   "[CA] double, [MCA] double, [gross13] double, [cash_adv] double, [net13] double, " & _
                   "[COSTCENTERID] char(10), [WORKCENTERID] char(10) )"
    End If

    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
    
'ErrCreate:
'    ' in case table is already existing, let's clear it...
'    If nMode = 0 Then
'        cSqlStmt = "DELETE FROM C13MOPAY"
'    Else
'        cSqlStmt = "DELETE FROM R13MOPAY"
'    End If
'    QueryDBF cSqlStmt, oTempADO, True
End Sub

Sub Gen13mobackup()
    Dim cSqlStmt As String, _
        cBackupPath As String, _
        oFileSys As New FileSystemObject, _
        aField As Variant, _
        aFieldValue As Variant, _
        nCtr As Integer
        
    aField = Array("", "", "")
        
    cBackupPath = CheckPath(Text4.Text)
    
    If oFileSys.FileExists(CheckPath(cBackupPath) & "C13MOPAY.DBF") Then
        oFileSys.DeleteFile CheckPath(cBackupPath) & "C13MOPAY.DBF"
    End If
    
    If oFileSys.FileExists(CheckPath(cBackupPath) & "R13MOPAY.DBF") Then
        oFileSys.DeleteFile CheckPath(cBackupPath) & "R13MOPAY.DBF"
    End If
    
    DetectDBF cBackupPath
    
    CreateTmpM13 0      ' --> contractual
    CreateTmpM13 1      ' --> regular
    
    
    For nCtr = 1 To 24
        aField(0) = aField(0) & "a.bas" & Format(nCtr, "00") & ","
        aField(1) = aField(1) & "a.gr" & Format(nCtr, "00") & ","
        aField(2) = aField(2) & "a.sa" & Format(nCtr, "00") & ","
    Next nCtr
    
'    cSqlStmt = "select a.empid, a.emp_stat, a.fullname, a.totgross, a.totbasic, a.totsa, b.emp_stat, a.tin, a.taxcode, " & _
'               aField(0) & aField(1) & aField(2) & _
'               " b.leave_pay, b.m13pay, b.ded_amt, b.net_pay, b.gross_pay " & _
'               "from pa13667 a left join pa87260 b on a.empid=b.empid and b.periodid=" & cQuote & Text2.Text & cQuote & _
'               " where a.year = (select year(date_start) from pa7730 where periodid=" & cQuote & Text2.Text & cQuote & ")"
    
    If gCompanyID = "0002" Then
        cSqlStmt = "select a.BACCNTNO,a.empid, a.date_hire,a.date_fin,a.date_res, a.emp_stat, a.fullname, a.totgross, a.totbasic, a.totsa, b.emp_stat, a.tin, a.taxcode, " & _
                   aField(0) & aField(1) & aField(2) & _
                   " a.leave_pay, b.m13pay, ifnull(c.ded_amt,0) as CA,ifnull(d.ded_amt,0) as MCA, b.ded_amt, b.net_pay, b.gross_pay,b.COSTCENTERID,b.WORKCENTERID " & _
                   " from pa13667 a left join pa87260 b on a.empid=b.empid and b.periodid=" & cQuote & Text2.Text & cQuote & _
                   " left join pa87263 c on a.empid=c.empid and c.periodid=b.periodid and c.dedid=" & cQuote & "009" & cQuote & _
                   " left join pa87263 d on a.empid=d.empid and d.periodid=b.periodid and d.dedid=" & cQuote & "017" & cQuote & _
                   " where a.year = (select year(date_start) from pa7730 where periodid=" & cQuote & Text2.Text & cQuote & ")"
    Else
        cSqlStmt = "select a.BACCNTNO,a.empid, a.date_hire,a.date_fin,a.date_res, a.emp_stat, a.fullname, a.totgross, a.totbasic, a.totsa, b.emp_stat, a.tin, a.taxcode, " & _
                   aField(0) & aField(1) & aField(2) & _
                   " a.leave_pay, b.m13pay, b.ded_amt, b.net_pay, b.gross_pay,b.COSTCENTERID,b.WORKCENTERID " & _
                   "from pa13667 a left join pa87260 b on a.empid=b.empid and b.periodid=" & cQuote & Text2.Text & cQuote & _
                   " where a.year = (select year(date_start) from pa7730 where periodid=" & cQuote & Text2.Text & cQuote & ")"
        
    End If
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
    
        ShowProgress 0
    
        While Not oTempADO.EOF
        
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
            aField = Array("", "", "")
            aFieldValue = Array("", "", "")
            For nCtr = 1 To 24
                aField(0) = aField(0) & "bas" & Format(nCtr, "00") & ","
                aField(1) = aField(1) & "gr" & Format(nCtr, "00") & ","
                aField(2) = aField(2) & "sa" & Format(nCtr, "00") & ","
            
                aFieldValue(0) = aFieldValue(0) & oTempADO("bas" & Format(nCtr, "00")) & ","
                aFieldValue(1) = aFieldValue(1) & oTempADO("gr" & Format(nCtr, "00")) & ","
                aFieldValue(2) = aFieldValue(2) & oTempADO("sa" & Format(nCtr, "00")) & ","
            Next nCtr
            If gCompanyID = "0002" Then
            
                cSqlStmt = "insert into " & IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), "C13MOPAY", "R13MOPAY") & "(BACCNTNO,empid,fullname,emp_stat," & _
                           IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), aField(0) & "totbasic,", "[tin],taxcode," & aField(1) & aField(2) & "totgross, totsa,") & _
                           "leave_pay,m13pay,gross13,cash_adv,net13,CA,MCA,date_hire,date_fin,date_res,COSTCENTERID,WORKCENTERID)values(" & _
                           cQuote & oTempADO("BACCNTNO") & cQuote & "," & _
                           cQuote & oTempADO("empid") & cQuote & "," & _
                           cQuote & oTempADO("fullname") & cQuote & "," & _
                           cQuote & IIf(oTempADO("emp_stat") = 1, "C", "R") & cQuote & "," & _
                           IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), aFieldValue(0) & oTempADO("totbasic") & ",", cQuote & oTempADO("tin") & cQuote & "," & cQuote & oTempADO("taxcode") & cQuote & "," & aFieldValue(1) & aFieldValue(2) & oTempADO("totgross") & "," & oTempADO("totsa") & ",") & _
                           oTempADO("leave_pay") & "," & _
                           oTempADO("m13pay") & "," & _
                           oTempADO("gross_pay") & "," & _
                           oTempADO("ded_amt") & "," & _
                           oTempADO("net_pay") & "," & _
                           oTempADO("CA") & "," & _
                           oTempADO("MCA") & "," & _
                           cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_res"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & oTempADO("COSTCENTERID") & cQuote & "," & _
                           cQuote & oTempADO("WORKCENTERID") & cQuote & ")"
            Else
                cSqlStmt = "insert into " & IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), "C13MOPAY", "R13MOPAY") & "(BACCNTNO,empid,fullname,emp_stat," & _
                           IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), aField(0) & "totbasic,", "[tin],taxcode," & aField(1) & aField(2) & "totgross, totsa,") & _
                           "leave_pay,m13pay,gross13,cash_adv,net13,date_hire,date_fin,date_res,COSTCENTERID,WORKCENTERID)values(" & _
                           cQuote & oTempADO("BACCNTNO") & cQuote & "," & _
                           cQuote & oTempADO("empid") & cQuote & "," & _
                           cQuote & oTempADO("fullname") & cQuote & "," & _
                           cQuote & IIf(oTempADO("emp_stat") = 1, "C", "R") & cQuote & "," & _
                           IIf((oTempADO("emp_stat") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)), aFieldValue(0) & oTempADO("totbasic") & ",", cQuote & oTempADO("tin") & cQuote & "," & cQuote & oTempADO("taxcode") & cQuote & "," & aFieldValue(1) & aFieldValue(2) & oTempADO("totgross") & "," & oTempADO("totsa") & ",") & _
                           oTempADO("leave_pay") & "," & _
                           oTempADO("m13pay") & "," & _
                           oTempADO("gross_pay") & "," & _
                           oTempADO("ded_amt") & "," & _
                           oTempADO("net_pay") & "," & _
                           cQuote & Format(oTempADO("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(oTempADO("date_res"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & oTempADO("COSTCENTERID") & cQuote & "," & _
                           cQuote & oTempADO("WORKCENTERID") & cQuote & ")"
                           
            End If
            Script2File cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            oTempADO.MoveNext
            
        Wend
        
        ShowProgress 4
        
        MsgBox "Done!"
    End If
    
    Set oFileSys = Nothing
End Sub


' + -->
' |     Procedure Name  :   GenAlphaList(ByVal cPeriodID As String)
' |     Description     :   Generate Alphalist (Annual Withholding Tax)
' |     Date Created    :   6 jan 2007
' + -->

Sub CreateAlphaLst()
    Dim cSqlStmt As String

'***sched 7.1***
    '-->7.1
    
    cSqlStmt = "create table alpha7_1 (" & _
               " [sched1] integer,          [sched2] char(20)," & _
               " [sched3a] char(100),       [sched3b] char(100)," & _
               " [sched3c] char(100),       [sched3d] char(100), " & _
               " [sched3e] char(100),       [sched4a] double," & _
               " [sched4b] double,          [sched4c] double," & _
               " [sched4d] double,          [sched4e] double," & _
               " [sched4f] double,          [sched4g] double," & _
               " [sched4h] double,          [sched4i] double," & _
               " [sched4j] double,          [sched5a] char(3)," & _
               " [sched5b] double,          [sched6] double," & _
               " [sched7] double,           [sched8] double," & _
               " [sched9] double,           [sched10a] double," & _
               " [sched10b] double,         [sched11] double," & _
               " [sched12] char(3) )"

    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
    
'***sched 7.3***
    '-->7.3
    
    cSqlStmt = "create table alpha7_3 (" & _
               " [sched1] integer,          [sched2] char(20)," & _
               " [sched3a] char(100),       [sched3b] char(100)," & _
               " [sched3c] char(100),       [sched4a] double," & _
               " [sched4b] double,          [sched4c] double," & _
               " [sched4d] double,          [sched4e] double," & _
               " [sched4f] double,          [sched4g] double," & _
               " [sched4h] double,          [sched4i] double," & _
               " [sched4j] double,          [sched5a] char(3)," & _
               " [sched5b] double,          [sched6] double," & _
               " [sched7] double,           [sched8] double," & _
               " [sched9] double,           [sched10a] double," & _
               " [sched10b] double,         [sched11] double," & _
               " [sched12] char(3) )"

    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
    
'***sched 7.5***
    '-->7.5
    
    cSqlStmt = "create table alpha7_5 (" & _
               " [sched1] integer,          [sched2] char(20)," & _
               " [sched3a] char(100),       [sched3b] char(100)," & _
               " [sched3c] char(100),       [sched4] char(100)," & _
               " [sched5a] double,          [sched5b] double,           [sched5c] double,          [sched5d] double," & _
               " [sched5e] double,          [sched5f] double,           [sched5g] double,          [sched5h] double," & _
               " [sched5i] double,          [sched5j] double,           [sched5k] double,          [sched5l] double," & _
               " [sched5m] double,          [sched5n] double,           [sched5o] char(100),       [sched5p] char(100)," & _
               " [sched5q] double,          [sched5r] double,           [sched5s] double,          [sched5t] double," & _
               " [sched5u] double,          [sched5v] double,           [sched5w] double,          [sched5x] double," & _
               " [sched5y] double,          [sched5z] double,           [sched5aa] double,         [sched5ab] double," & _
               " [sched5ac] double,         [sched5ad] double,          [sched5ae] double,         [sched5af] double," & _
               " [sched5ag] double,         [sched6] char(3)," & _
               " [sched6b] double,          [sched7] double," & _
               " [sched8] double,           [sched9] double," & _
               " [sched10a] double,         [sched10b] double," & _
               " [sched11a] double,         [sched11b] double," & _
               " [sched12] double,          [sched13] double )"


    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
    
End Sub

Sub GenAlphaList(ByVal cPeriodID As String)
    Dim cDedID As String, _
        cSqlStmt As String, _
        cParam As String, _
        cParam2 As String, _
        cParam3 As String, _
        cParam4 As String, _
        cParam5 As String, _
        cString As String, _
        oFileSys As New FileSystemObject, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        oTaxSet As New ADODB.Recordset, _
        nCtr As Integer, _
        nYear As Integer, _
        n13m_SLVL As Double, _
        l13month As String, _
        nTaxAmt As Double, _
        nTaxAmt2 As Double, _
        nWTax As Double, _
        aTaxable7_3 As Variant, _
        aTaxable7_5 As Variant, _
        aTaxable As Variant, _
        lIba As Boolean
    
    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_1.DBF") Then
        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_1.DBF"
    End If
    DetectDBF CheckPath(Text4.Text)
    
    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_3.DBF") Then
        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_3.DBF"
    End If
    DetectDBF CheckPath(Text4.Text)
    
    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_5.DBF") Then
        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_5.DBF"
    End If
    DetectDBF CheckPath(Text4.Text)

    CreateAlphaLst

    For nCtr = 0 To UBound(aTaxExempt)
        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
        cDedID = cDedID & aTaxExempt(nCtr) & ","
    Next nCtr
    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)


    cSqlStmt = "select periodid, pclose " & _
               "From PA7730 " & _
               "Where (13month=0) and ((Year(date_start) = " & Combo1.Text & ") Or (Year(date_end) = " & Combo1.Text & "))" & _
               " order by date_start "
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
'        nFldCnt = objdbRs.RecordCount
'        cField = ""
'        cFldValue = ""
        cParam = ""
        cParam2 = ""
        While Not objdbRs.EOF
            cString = Chr$(97 + objdbRs.AbsolutePosition)
            'k1 computation 20081206
'            cField = cField & _
'                     "gr" & Format(objdbRs.AbsolutePosition, "00") & ","

'20120118 revise as per advice of accounting
'            cParam = cParam & _
'                     " ifnull(" & cString & ".gross_pay,0 ) + "

            If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
                cParam = cParam & _
                     " ifnull(" & cString & ".gross_pay,0 ) + "
            Else
                cParam = cParam & _
                         " ifnull(" & cString & ".gross_pay,0 ) + ifnull(" & cString & ".sa_net_pay,0 ) + "
            End If
            
'            cParam = cParam & _
'                     " ifnull(" & cString & ".gross_pay,0 ) + ifnull(" & cString & ".sa_net_pay,0 ) + "
                     
                     
            cParam3 = cParam3 & _
                     " ifnull(" & cString & ".basicpay,0 ) + "
            cParam4 = cParam4 & _
                     " ifnull(" & cString & ".wtax,0 ) + "
                     
                     
'            MsgBox cParam
            cParam2 = cParam2 & " left join " & IIf(objdbRs("pclose") = 1, "pah87260 ", "pa87260 ") & cString & _
                      " on a.empid=" & cString & ".empid and " & cString & ".periodid=" & cQuote & objdbRs("periodid") & cQuote & vbCrLf
            
            objdbRs.MoveNext
        Wend
    End If

    cSqlStmt = "select periodid, pclose " & _
               "From PA7730 " & _
               "Where (13month=0) and ((Year(date_start) = " & Combo1.Text & ") Or (Year(date_end) = " & Combo1.Text & "))" & _
               " and ((month(date_start)<>12) or (month(date_end)<>12)) order by date_start "
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
'        nFldCnt = objdbRs.RecordCount
'        cField = ""
'        cFldValue = ""
        cParam4 = ""
        cParam5 = ""
        While Not objdbRs.EOF
            cString = Chr$(97 + objdbRs.AbsolutePosition)
            cParam4 = cParam4 & _
                     " ifnull(" & cString & ".wtax,0 ) + "
                     
'            MsgBox cParam
            cParam5 = cParam5 & " left join " & IIf(objdbRs("pclose") = 1, "pah87260 ", "pa87260 ") & cString & _
                      " on a.empid=" & cString & ".empid and " & cString & ".periodid=" & cQuote & objdbRs("periodid") & cQuote & vbCrLf
            
            objdbRs.MoveNext
        Wend
    End If



'    aTaxable7_3(0) = gross_pay
'   non -Taxable
'    aTaxable7_3(1) = 13month
'    aTaxable7_3(2) = non_tax
'    aTaxable7_3(3) = total of non tax
'   Taxable
'    aTaxable7_3(4) = basic
'    aTaxable7_3(5) = 13month
'   aTaxable7_3(13) = Salaries and other forms
'    aTaxable7_3(6) = total of taxable
'    aTaxable7_3(7) = Tax Amount
'    aTaxable7_3(8) = Net Taxable Compensation Income
'    aTaxable7_3(9) = Tax Due
'    aTaxable7_3(10) = Amount withheld and paid for in december
'    aTaxable7_3(11) = Over withheld tax refunded to employee
'    aTaxable7_3(12) = Amount of tax withheld as asjusted

'aTaxable7_3(13) = aTaxable7_3(0) - aTaxable7_3(3) - aTaxable7_3(4) - aTaxable7_3(5)

'aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))

'
    ShowProgress 0
'7.3 regular

    aTaxable7_3 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)

    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            l13month = IIf(objdbRs("pclose") = 0, False, True)
            objdbRs.MoveNext
        Wend
    End If

 '    --> active employee
    nCtr = 1
    cSqlStmt = "select a.empid, ifnull(c.lastname,a.lastname) as lastname,ifnull(c.firstname,a.firstname) as firstname,ifnull(c.mname,a.mname) as mname, c.tin, ifnull(e.taxcode,'S') as taxcode , " & _
               "  round(c.ytd_gross + a.gross_pay,2) as gross_pay, " & _
               "  round(c.ytd_gross_sa + a.SA_NET_PAY,2) as gross_pay_sa, " & _
               "  round((a.SUN_PAY + a.SUN_OT_PAY + a.SUN_COLA + a.SUN_ND_PAY + a.SUN_ND_OT_PAY),2) as suntot, " & _
               "  round(c.ytd_basic,2) as basic, " & _
               "  round(b.ded_amt3,2) as non_tax, " & _
               " (c.sl_avail)+ (c.vl_avail) as leave_unuse, " & _
               "  c.rate_amt, " & _
               "  round(c.ytd_wtax,2) as tax_wheld, " & _
               "  round(b.ded_amt2, 2) As adj_tax " & _
               "from pah87260 a left join pah87263 b on a.periodid=b.periodid and a.empid=b.empid and b.dedid='006' " & _
               "  left join di3670 c on a.empid=c.empid " & _
               "  left join pa8290 e on c.taxid=e.taxid " & _
               "where c.active=0 and c.emp_stat=2 " & _
               " and c.rate_amt <>" & cQuote & gBasicRate & cQuote & _
               " and a.periodid=" & cQuote & cPeriodID & cQuote & _
               " order by a.fullname"
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")


'            If oRecordSet("empid") = "000683" Then MsgBox "stop"


            cSqlStmt = "select a.empid, " & _
                        cParam & "0 as totgrosspay, " & _
                        cParam3 & "0 as totbasicpay " & _
                       " from di3670 a " & _
                       cParam2 & _
                       "where a.empid =" & cQuote & oRecordSet("empid") & cQuote
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

'2010-01-07
'            If gCompanyID <> "0002" Then
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay"), 2)
'            Else
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay") + (oRecordSet("gross_pay_sa") - (oRecordSet("suntot") + objdbRs("suntot"))), 2)
'            End If

            aTaxable7_3(0) = Round(objdbRs("totgrosspay"), 2)
            aTaxable7_3(4) = Round(objdbRs("totbasicpay"), 2)

            aTaxable7_3(2) = oRecordSet("non_tax")

            cSqlStmt = "select a.empid, " & _
                        cParam4 & "0 as totwtax " & _
                       " from di3670 a " & _
                       cParam5 & _
                       "where a.empid =" & cQuote & oRecordSet("empid") & cQuote

            OpenQueryDNS cSqlStmt, objdbRs, False


            nWTax = Round(objdbRs("totwtax"), 2)
'            If oRecordSet("empid") = "235627" Then MsgBox "stop"

            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
                       " From pah87263 " & _
                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
                       " (dedid in (" & cDedID & ")) " & _
                       " and (periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ")) " & _
                       " group by empid "
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

            aTaxable7_3(2) = objdbRs("non_tax")

            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
                       " From pa87263 " & _
                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
                       " (dedid in (" & cDedID & ")) " & _
                       " and periodid=" & cQuote & cPeriodID & cQuote & _
                       " group by empid "

            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aTaxable7_3(2) = aTaxable7_3(2) + objdbRs("non_tax")
            Else
                aTaxable7_3(2) = aTaxable7_3(2) + 0
            End If



            cSqlStmt = "select gross_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
                       " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
                       "  and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"

'            Script2File cSqlStmt
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
            
                n13m_SLVL = Round((objdbRs("gross_pay") + (oRecordSet("leave_unuse") * oRecordSet("rate_amt"))), 2)

                If n13m_SLVL > 30000 Then
                    aTaxable7_3(1) = 30000
                    aTaxable7_3(5) = n13m_SLVL - 30000
                Else
                    aTaxable7_3(1) = n13m_SLVL
                    aTaxable7_3(5) = 0
                End If
            End If

'            aTaxable7_3(0) = aTaxable7_3(0) + n13m_SLVL

            'total of non tax
            aTaxable7_3(3) = Round(aTaxable7_3(1) + aTaxable7_3(2), 2)



            'aTaxable7_3(6) = total of taxable

            aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))

            'taxable - salaries and other forms
            aTaxable7_3(13) = aTaxable7_3(0) - aTaxable7_3(3) - aTaxable7_3(4) - aTaxable7_3(5)

            aTaxable7_3(6) = Round(aTaxable7_3(4) + aTaxable7_3(5) + aTaxable7_3(13), 2)



'            'old taxcode n amount
'            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
'            OpenQueryDNS cSqlStmt, oTaxSet, False
'
'            If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                nTaxAmt = oTaxSet("S_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                nTaxAmt = oTaxSet("H_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                nTaxAmt = oTaxSet("M_AMT")
'            End If
'
'            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'            End If


            'new taxcode amount
            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
            OpenQueryDNS cSqlStmt, oTaxSet, False

            If InStr(oRecordSet("taxcode"), "S") = 1 Then
                nTaxAmt2 = oTaxSet("S_AMT")
            End If

            If InStr(oRecordSet("taxcode"), "H") = 1 Then
                nTaxAmt2 = oTaxSet("H_AMT")
            End If

            If InStr(oRecordSet("taxcode"), "M") = 1 Then
                nTaxAmt2 = oTaxSet("M_AMT")
            End If

            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
                nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
            End If

            'aTaxable7_3(7) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
            aTaxable7_3(7) = Round(nTaxAmt2, 2)

            aTaxable7_3(8) = Round(aTaxable7_3(6) - aTaxable7_3(7), 2)

            'tax due
            If gCompanyID = "0002" Then
                aTaxable7_3(9) = nWTax + oRecordSet("adj_tax")
            Else
                cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
                           " Where range1 <= " & aTaxable7_3(8) & " And range2 >=" & aTaxable7_3(8)
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    aTaxable7_3(9) = Round(((aTaxable7_3(8) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
                Else
                    aTaxable7_3(9) = 0
                End If
            End If

'nWTax

            If gCompanyID = "0002" Then
                aTaxable7_3(10) = Round(nWTax - aTaxable7_3(9), 2)
                aTaxable7_3(11) = Round(aTaxable7_3(9) - nWTax, 2)
            Else
            aTaxable7_3(10) = Round(aTaxable7_3(9) - nWTax, 2)

            aTaxable7_3(11) = Round(nWTax - aTaxable7_3(9), 2)
            End If
            aTaxable7_3(12) = Round(nWTax + aTaxable7_3(10), 2)

            cSqlStmt = "insert into alpha7_3(sched1,sched2,sched3a,sched3b,sched3c,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g, " & _
                       "sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12)values(" & _
                       nCtr & "," & _
                       cQuote & oRecordSet("tin") & cQuote & "," & _
                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & _
                       aTaxable7_3(0) & "," & aTaxable7_3(1) & ",0," & aTaxable7_3(2) & ",0," & aTaxable7_3(3) & "," & aTaxable7_3(4) & "," & aTaxable7_3(5) & "," & aTaxable7_3(13) & "," & aTaxable7_3(6) & "," & _
                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
                       aTaxable7_3(7) & ",0," & _
                       aTaxable7_3(8) & "," & _
                       aTaxable7_3(9) & "," & _
                       nWTax & "," & _
                       aTaxable7_3(10) & "," & aTaxable7_3(11) & "," & aTaxable7_3(12) & "," & _
                       cQuote & "N" & cQuote & ")"

'                MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True

            nCtr = nCtr + 1
            oRecordSet.MoveNext
        Wend
    End If

    ShowProgress 4
'
'
' ----------------------------------------------2011-----------------------------


'7.1 finished/resigned



'    aTaxable7_3(0) = gross_pay
'   non -Taxable
'    aTaxable7_3(1) = 13month
'    aTaxable7_3(2) = non_tax
'    aTaxable7_3(3) = total of non tax
'   Taxable
'    aTaxable7_3(4) = basic
'    aTaxable7_3(5) = 13month
'   aTaxable7_3(13) = Salaries and other forms
'    aTaxable7_3(6) = total of taxable
'    aTaxable7_3(7) = Tax Amount
'    aTaxable7_3(8) = Net Taxable Compensation Income
'    aTaxable7_3(9) = Tax Due
'    aTaxable7_3(10) = Amount withheld and paid for in december
'    aTaxable7_3(11) = Over withheld tax refunded to employee
'    aTaxable7_3(12) = Amount of tkax withheld as asjusted

'aTaxable7_3(13) = aTaxable7_3(0) - aTaxable7_3(3) - aTaxable7_3(4) - aTaxable7_3(5)

'aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))

'aTaxable7_3(3) = Round(aTaxable7_3(1) + aTaxable7_3(2), 2)


    ShowProgress 0

    aTaxable7_3 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)

    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            l13month = IIf(objdbRs("pclose") = 0, False, True)
            objdbRs.MoveNext
        Wend
    End If

 '    --> active employee
    nCtr = 1

    cSqlStmt = " select a.empid,a.date_hire,a.date_res, " & _
               " a.lastname,a.firstname,a.mname, " & _
               " a.tin,ifnull(b.taxcode,a.taxcode) as taxcode, " & _
               " round(a.ytd_gross,2) as gross_pay, " & _
               " round(a.ytd_gross_sa,2) as gross_pay_sa, " & _
               " round(a.ytd_basic,2) as basic, " & _
               " (a.sl_avail)+ (a.vl_avail) as leave_unuse, " & _
               " a.rate_amt, " & _
               " a.ytd_wtax as tax_wheld, " & _
               " sum(c.ded_amt) As non_tax " & _
               "from di3670 a left join pa8290 b on a.taxid=b.taxid " & _
               "  left join pah87263 c on a.empid=c.empid and (c.periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ")) " & _
               "Where (a.emp_stat = 2) " & _
               "  and (((a.active=1) and (year(a.date_res)=" & Combo1.Text & "))" & _
               "        or ((a.active=3) and (year(a.date_res)=" & Combo1.Text & "))" & _
               "        or ((a.active=2) and (year(a.date_fin)=" & Combo1.Text & ")))" & _
               "  and (c.dedid in (" & cDedID & ")) " & _
               " and a.rate_amt >" & cQuote & gBasicRate & cQuote & _
               "group by c.empid " & _
               " order by a.lastname,a.firstname,a.mname "

'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")


            cSqlStmt = "select a.empid, " & _
                        cParam & "0 as totgrosspay, " & _
                        cParam3 & "0 as totbasicpay " & _
                       " from di3670 a " & _
                       cParam2 & _
                       "where a.empid =" & cQuote & oRecordSet("empid") & cQuote
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False

'2010-01-07
'            If gCompanyID <> "0002" Then
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay"), 2)
'            Else
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay") + (oRecordSet("gross_pay_sa") - (oRecordSet("suntot") + objdbRs("suntot"))), 2)
'            End If

            aTaxable7_3(0) = Round(objdbRs("totgrosspay"), 2)
            aTaxable7_3(4) = Round(objdbRs("totbasicpay"), 2)

            aTaxable7_3(2) = oRecordSet("non_tax")

            cSqlStmt = "select a.empid, " & _
                        cParam4 & "0 as totwtax " & _
                       " from di3670 a " & _
                       cParam5 & _
                       "where a.empid =" & cQuote & oRecordSet("empid") & cQuote
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False


            nWTax = Round(objdbRs("totwtax"), 2)
'            If oRecordSet("empid") = "235602" Then MsgBox "stop"

'            If gCompanyID <> "0002" Then
                If oRecordSet("non_tax") = 0 Then

                    cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
                               " From pah87263 " & _
                               " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
                               " (dedid in (" & cDedID & ")) " & _
                               " and (periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ")) " & _
                               " group by empid "

                    OpenQueryDNS cSqlStmt, objdbRs, False

                    aTaxable7_3(2) = objdbRs("non_tax")

                    cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
                               " From pa87263 " & _
                               " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
                               " (dedid in (" & cDedID & ")) " & _
                               " and periodid=" & cQuote & cPeriodID & cQuote & _
                               " group by empid "

                    OpenQueryDNS cSqlStmt, objdbRs, False
                    If objdbRs.RecordCount > 0 Then
                        aTaxable7_3(2) = aTaxable7_3(2) + objdbRs("non_tax")
                    Else
                        aTaxable7_3(2) = aTaxable7_3(2) + 0
                    End If

                End If

'            Else
'                aTaxable7_3(2) = oRecordSet("non_tax")
'            End If

'            cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
'                       " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       "  and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"

            cSqlStmt = "select sum(m13pay) as gross_pay  from " & IIf(l13month = "False", "pa87260", "pah87260") & _
                       " where (periodid in (select periodid from pa7730 where year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
                       "  and (empid=" & cQuote & oRecordSet("empid") & cQuote & ") group by empid"

'            Script2File cSqlStmt
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then

                n13m_SLVL = Round(objdbRs("gross_pay"), 2)

                If n13m_SLVL > 30000 Then
                    aTaxable7_3(1) = 30000
                    aTaxable7_3(5) = n13m_SLVL - 30000
                Else
                    aTaxable7_3(1) = n13m_SLVL
                    aTaxable7_3(5) = 0
                End If
            End If

            aTaxable7_3(0) = aTaxable7_3(0) + n13m_SLVL

            'total of non tax
            aTaxable7_3(3) = Round(aTaxable7_3(1) + aTaxable7_3(2), 2)



            'aTaxable7_3(6) = total of taxable
'            aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))

            'taxable - salaries and other forms
            'aTaxable7_3(13) = aTaxable7_3(0) - aTaxable7_3(3) - aTaxable7_3(4) - aTaxable7_3(5)
            '20130122
            aTaxable7_3(13) = aTaxable7_3(0) - aTaxable7_3(3) - aTaxable7_3(4)
            

            aTaxable7_3(6) = Round(aTaxable7_3(4) + aTaxable7_3(5) + aTaxable7_3(13), 2)



'            'old taxcode n amount
'            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
'            OpenQueryDNS cSqlStmt, oTaxSet, False
'
'            If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                nTaxAmt = oTaxSet("S_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                nTaxAmt = oTaxSet("H_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                nTaxAmt = oTaxSet("M_AMT")
'            End If
'
'            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'            End If


            'new taxcode amount
            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
            OpenQueryDNS cSqlStmt, oTaxSet, False

            If InStr(oRecordSet("taxcode"), "S") = 1 Then
                nTaxAmt2 = oTaxSet("S_AMT")
            End If

            If InStr(oRecordSet("taxcode"), "H") = 1 Then
                nTaxAmt2 = oTaxSet("H_AMT")
            End If

            If InStr(oRecordSet("taxcode"), "M") = 1 Then
                nTaxAmt2 = oTaxSet("M_AMT")
            End If

            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") And (right(oRecordSet("taxcode"), 1) <> "") Then
                nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
            End If

            'aTaxable7_3(7) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
            aTaxable7_3(7) = Round(nTaxAmt2, 2)

            aTaxable7_3(8) = Round(aTaxable7_3(6) - aTaxable7_3(7), 2)

            'tax due
            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
                       " Where range1 <= " & aTaxable7_3(8) & " And range2 >=" & aTaxable7_3(8)
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aTaxable7_3(9) = Round(((aTaxable7_3(8) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
            Else
                aTaxable7_3(9) = 0
            End If

            aTaxable7_3(10) = Round(aTaxable7_3(9) - nWTax, 2)

            aTaxable7_3(11) = Round(nWTax - aTaxable7_3(9), 2)

            aTaxable7_3(12) = Round(nWTax + aTaxable7_3(10), 2)

'            cSqlStmt = "insert into alpha7_1(sched1,sched2,sched3a,sched3b,sched3c,sched3d,sched3e,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g, " & _
'                       "sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12)values(" & _
'                       nCtr & "," & _
'                       cQuote & orecordset("tin") & cQuote & "," & _
'                       cQuote & orecordset("lastname") & cQuote & "," & cQuote & orecordset("firstname") & cQuote & "," & cQuote & orecordset("mname") & cQuote & "," & _
'                       cQuote & Format(orecordset("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
'                       cQuote & Format(orecordset("date_res"), "yyyy-mm-dd") & cQuote & "," & _
'                       aTaxable7_3(0) & "," & aTaxable7_3(1) & ",0," & orecordset("ded_amt") & ",0," & aTaxable7_3(3) & "," & aTaxable7_3(4) & "," & aTaxable7_3(5) & ",0," & aTaxable7_3(6) & "," & _
'                       cQuote & Replace(orecordset("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_3(7) & ",0," & _
'                       aTaxable7_3(8) & "," & _
'                       aTaxable7_3(9) & "," & _
'                       orecordset("tax_wheld") & "," & _
'                       aTaxable7_3(10) & "," & aTaxable7_3(11) & "," & aTaxable7_3(12) & "," & _
'                       cQuote & "N" & cQuote & ")"


            cSqlStmt = "insert into alpha7_1(sched1,sched2,sched3a,sched3b,sched3c,sched3d,sched3e,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g, " & _
                       "sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12)values(" & _
                       nCtr & "," & _
                       cQuote & oRecordSet("tin") & cQuote & "," & _
                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                       cQuote & Format(oRecordSet("date_res"), "yyyy-mm-dd") & cQuote & "," & _
                       aTaxable7_3(0) & "," & aTaxable7_3(1) & ",0," & aTaxable7_3(2) & ",0," & aTaxable7_3(3) & "," & aTaxable7_3(4) & "," & aTaxable7_3(5) & "," & aTaxable7_3(13) & "," & aTaxable7_3(6) & "," & _
                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
                       aTaxable7_3(7) & ",0," & _
                       aTaxable7_3(8) & "," & _
                       aTaxable7_3(9) & "," & _
                       nWTax & "," & _
                       aTaxable7_3(10) & "," & aTaxable7_3(11) & "," & aTaxable7_3(12) & "," & _
                       cQuote & "N" & cQuote & ")"

'                MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True


            nCtr = nCtr + 1
            oRecordSet.MoveNext
        Wend
    End If

    ShowProgress 4



'---------------------------------------------------------------------------------------------------------------

'-----------------------------------------2011---------------------------------
'-----7.5
'-----Regular
    ShowProgress 0
'(aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(8)
'    aTaxable7_5(0) = grosspay
'    aTaxable7_5(1) = basic per day
'    aTaxable7_5(2) = basic per month
'    aTaxable7_5(3) = basic per year
'    aTaxable7_5(4) = holiday
'    aTaxable7_5(5) = ot
'    aTaxable7_5(6) = night shift differential
'non-taxable
'    aTaxable7_5(7) = 13 month pay
'    aTaxable7_5(8) = nontax
'non-taxable
'    aTaxable7_5(9) = 13 month pay

'    aTaxable7_5(10) = total compensation
'    aTaxable7_5(11) = Tax Amount
'    aTaxable7_5(12) = Net taxable Compensation Income
'    aTaxable7_5(13) = Tax Due

'    aTaxable7_5(14) = Amount withheld and paid for in december

'    aTaxable7_5(15) = Over withheld tax refunded to employee
'    aTaxable7_5(16) = Amount of tax withheld as asjusted
'    aTaxable7_5(17) = Salaries and Other forms of compensation
'    aTaxable7_5(19) = basicpay + colapay + nd pay

'aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(8))

    aTaxable7_5 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    aTaxable = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)

    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        While Not objdbRs.EOF
            l13month = IIf(objdbRs("pclose") = 0, False, True)
            objdbRs.MoveNext
        Wend
    End If

    ' --> active employee minimun wage earner
    nCtr = 1

    cSqlStmt = " select b.cola_amt,b.rate_amt,a.paystatus,b.emp_stat,b.active,a.periodid,ifnull(b.rate_amt,0) as rate_amt,a.empid,ifnull(b.lastname,a.lastname) as lastname,ifnull(b.firstname,a.firstname) as firstname,ifnull(b.mname,a.mname) as mname, ifnull(b.tin,'') as tin,ifnull(c.taxcode,'S') as taxcode,ifnull(b.date_hire,'') as date_hire, ifnull(b.date_fin,'') as date_fin, " & _
               " round(sum(ifnull(a.gross_pay,'')),2) as gross_pay, " & _
               " round(sum(a.SA_NET_PAY),2) as gross_pay_sa, " & _
               " round((sum(a.SUN_PAY) + sum(a.SUN_OT_PAY) + sum(a.SUN_COLA) + sum(a.SUN_ND_PAY) + sum(a.SUN_ND_OT_PAY)),2) as suntot, " & _
               " round(sum(a.reg_pay) + sum(a.ndiff_pay),2) as basic, " & _
               " round(sum(a.hol_pay),2) as hol_pay, " & _
               " round(sum(a.ndiff_pay)+sum(a.ndiff_ot_pay)+sum(a.sa_ndiff_pay),2) as ndiff, " & _
               " round(sum(a.reg_ot_pay)+sum(a.sa_reg_pay),2) as ot, " & _
               " round(sum(a.SUN_OT_PAY),2) as sa_sun_ot_pay, " & _
               " round(ifnull(b.ytd_wtax,0), 2) As tax_wheld, " & _
               " (b.sl_avail)+ (b.vl_avail) as leave_unuse " & _
               " from pah87260 a left join di3670 b on a.empid=b.empid " & _
               "  left join pa8290 c on b.taxid=c.taxid or a.taxid=c.taxid " & _
               " where a.periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & "  and 13month <> 1 ) and b.paystatus<>1 " & _
               " and b.rate_amt <=" & cQuote & gBasicRate & cQuote & _
               " group by b.emp_stat,b.empid order by b.emp_stat,b.active,a.fullname "

               '" and a.empid in ('025970','028346','028460','021029','021175')" & _

'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF

            If oRecordSet("empid") = "080652" Then MsgBox "stop"

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
'            If orecordset("empid") = "0716" Then MsgBox "stop"
            cSqlStmt = " select cola_amt,empid,rate_amt,round(sum(gross_pay),2) as gross_pay, " & _
                       " round(SA_NET_PAY,2) as gross_pay_sa, " & _
                       " round(SUN_PAY + SUN_OT_PAY + SUN_COLA + SUN_ND_PAY + SUN_ND_OT_PAY,2) as suntot, " & _
                       " round(hol_pay,2) as hol_pay, " & _
                       " round(reg_pay + ndiff_pay, 2) As basic, " & _
                       " round(ndiff_pay + ndiff_ot_pay+sa_ndiff_pay,2) as ndiff, " & _
                       " round(reg_ot_pay+sa_reg_pay,2) as ot, " & _
                       " round(SUN_OT_PAY,2) as sa_sun_ot_pay, " & _
                       " round(wtax,2) as tax_wheld " & _
                       " From pa87260 " & _
                       " where periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & " and 13month=0) " & _
                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
                       " group by empid "
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, oRSet, False

            If oRSet.RecordCount > 0 Then
                aTaxable7_5(1) = Round(oRSet("rate_amt") + oRSet("cola_amt"), 2)
'                aTaxable7_5(2) = Round(orecordset("basic") + orecordset("cola_amt"), 2)
'                aTaxable7_5(3) = Round(oRSet("basic") + oRSet("cola_amt") + orecordset("cola_amt") + orecordset("basic"), 2)
                aTaxable7_5(2) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 26, 2)
                aTaxable7_5(3) = Round((oRSet("rate_amt") + oRSet("cola_amt") + oRecordSet("cola_amt") + oRecordSet("rate_amt")) * 312, 2)

                aTaxable7_5(4) = Round(oRSet("hol_pay") + oRecordSet("hol_pay"), 2)

'2010-01-19
'                If gCompanyID <> "0002" Then
'                    If gCompanyID = "0003" Then
'                        aTaxable7_5(5) = Round((oRSet("ot") + oRSet("sa_sun_ot_pay")) + (orecordset("ot") + orecordset("sa_sun_ot_pay")), 2)
'                    Else
'                        aTaxable7_5(5) = Round(oRSet("ot") + orecordset("ot"), 2)
'                    End If
'                Else
'                    aTaxable7_5(5) = Round((oRSet("gross_pay_sa") + orecordset("gross_pay_sa")) - oRSet("suntot") + orecordset("suntot"), 2)
'                End If

'                aTaxable7_5(5) = Round(oRSet("ot") + orecordset("ot"), 2)

                aTaxable7_5(6) = Round(oRSet("ndiff") + oRecordSet("ndiff"), 2)
'                aTaxable7_5(18) = (oRSet("gross_pay") + orecordset("gross_pay"))

' revise 20120120
'                If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
'
'                    aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
'                    aTaxable7_5(5) = Round(oRSet("ot") + oRecordSet("ot") + oRSet("sa_sun_ot_pay") + oRecordSet("sa_sun_ot_pay"), 2)
'
'                Else
'                    aTaxable7_5(18) = oRecordSet("gross_pay")
'                    aTaxable7_5(5) = Round(oRecordSet("ot"), 2)
'                End If

                aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
                aTaxable7_5(5) = Round(oRSet("ot") + oRecordSet("ot") + oRSet("sa_sun_ot_pay") + oRecordSet("sa_sun_ot_pay"), 2)


                aTaxable7_5(19) = Round(oRSet("basic") + oRecordSet("basic") + oRSet("cola_amt") + oRecordSet("cola_amt"), 2)
            Else

                aTaxable7_5(1) = Round(oRecordSet("rate_amt") + oRecordSet("cola_amt"), 2)
'                aTaxable7_5(2) = Round(orecordset("basic") + orecordset("cola_amt"), 2)
'                aTaxable7_5(3) = Round(orecordset("basic") + orecordset("cola_amt"), 2)
                aTaxable7_5(2) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 26, 2)
                aTaxable7_5(3) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 312, 2)

                aTaxable7_5(4) = Round(oRecordSet("hol_pay"), 2)



'                If gCompanyID <> "0002" Then
'                    If gCompanyID = "0003" Then
'                        aTaxable7_5(5) = Round(orecordset("ot") + orecordset("sa_sun_ot_pay"), 2)
'                    Else
'                        aTaxable7_5(5) = Round(orecordset("ot"), 2)
'                    End If
'                Else
'                    aTaxable7_5(5) = Round((orecordset("gross_pay_sa")) - orecordset("suntot"), 2)
'                End If

'                aTaxable7_5(5) = Round(orecordset("ot"), 2)

                aTaxable7_5(6) = Round(oRecordSet("ndiff"), 2)

'revise 20120120
'                If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
'                    aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
'                    aTaxable7_5(5) = Round(oRecordSet("ot") + oRecordSet("sa_sun_ot_pay"), 2)
'
'                Else
'                    aTaxable7_5(18) = oRecordSet("gross_pay")
'                    aTaxable7_5(5) = Round(oRecordSet("ot"), 2)
'                End If

                aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
                aTaxable7_5(5) = Round(oRecordSet("ot") + oRecordSet("sa_sun_ot_pay"), 2)




'                aTaxable7_5(18) = orecordset("gross_pay")
                aTaxable7_5(19) = Round(oRecordSet("basic") + oRecordSet("cola_amt"), 2)
            End If

'            If orecordset("emp_stat") = 2 Or orecordset("emp_stat") = 1 Then
'                If orecordset("paystatus") = 0 Then
'                    cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
'                               " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                               "      and (empid=" & cQuote & orecordset("empid") & cQuote & ")"
'        '            Script2File cSqlStmt
'        '            MsgBox cSqlStmt
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'
'                        n13m_SLVL = Round(objdbRs("net_pay"), 2)
'
'                    Else
'
'                        cSqlStmt = "select net_pay from pah87260" & _
'                                   " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                                   "      and (empid=" & cQuote & orecordset("empid") & cQuote & ")"
''                        MsgBox cSqlStmt
'                        OpenQueryDNS cSqlStmt, objdbRs, False
'                        If objdbRs.RecordCount > 0 Then
'
'                            n13m_SLVL = Round(objdbRs("net_pay"), 2)
'
'                        Else
'                            aTaxable7_5(7) = 0
'                            aTaxable7_5(9) = 0
'                        End If
'                    End If
'
'                    If n13m_SLVL <> 0 Then
'                       If n13m_SLVL > 30000 Then
'                            aTaxable7_5(7) = 30000
'                            aTaxable7_5(9) = n13m_SLVL - 30000
'                        Else
'                            aTaxable7_5(7) = n13m_SLVL
'                            aTaxable7_5(9) = 0
'                        End If
'                    Else
'                        aTaxable7_5(7) = 0
'                        aTaxable7_5(9) = 0
'                    End If
'                Else
'                    n13m_SLVL = 0
'                End If
'            Else
'                n13m_SLVL = 0
'            End If


            ' 20130125
'            cSqlStmt = "select empid,sum(m13pay) as  netpay  from pah87260" & _
'                       " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ") group by empid"

            cSqlStmt = "select empid,sum(m13pay) as  netpay  from pah87260" & _
                       " where periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & " and 13month=0) " & _
                       "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ") group by empid"


'            Script2File cSqlStmt
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then

                n13m_SLVL = Round(objdbRs("netpay"), 2)
                If n13m_SLVL > 30000 Then
                     aTaxable7_5(7) = 30000
                     aTaxable7_5(9) = n13m_SLVL - 30000
                Else
                     aTaxable7_5(7) = n13m_SLVL
                     aTaxable7_5(9) = 0
                End If
            Else
                aTaxable7_5(7) = 0
                aTaxable7_5(9) = 0
            End If

            aTaxable7_5(7) = aTaxable7_5(7) + (oRecordSet("leave_unuse") * oRecordSet("rate_amt"))


            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
                       " From pah87263 " & _
                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
                       " (dedid in (" & cDedID & ")) " & _
                       " and (periodid in (select periodid from pa7730 where year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
                       " group by empid "
'            Script2File cSqlStmt
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                aTaxable7_5(8) = objdbRs("non_tax")
            Else
                aTaxable7_5(8) = 0
            End If

            'aTaxable7_5(0) = aTaxable7_5(3) + aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(7)

'revise 20120120
'            If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
                aTaxable7_5(0) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa") + aTaxable7_5(7)
'            Else
'                aTaxable7_5(0) = oRecordSet("gross_pay")
'            End If

'            aTaxable7_5(0) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")


'            aTaxable7_5(0) = Holpay + (Round((oRecordSet("gross_pay_sa")) - oRecordSet("suntot"), 2)) + _
'                            ndiff + a13month

'            aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))

            'aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
            aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(8))

            '    aTaxable7_5(4) = holiday
'    aTaxable7_5(5) = ot
'    aTaxable7_5(6) = night shift differential
'non-taxable
'    aTaxable7_5(7) = 13 month pay
'    aTaxable7_5(8) = nontax

'                aTaxable7_5(18) = orecordset("gross_pay") - (excess13monthpay + deduction)

            aTaxable7_5(17) = 0

'            cSqlStmt = "select * from pa7730 where month(date_start) > 6  and periodid=" & cQuote & orecordset("periodid") & cQuote
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                ' non-taxable july - december
''                aTaxable7_5(18) = aTaxable7_5(0)
'                aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'
''                aTaxable7_5(18) = orecordset("gross_pay") - (excess13monthpay + deduction)
'
'
'                aTaxable7_5(17) = 0
'
'            Else
'                ' taxable January - June
'                '20100114
'                'aTaxable7_5(17) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'                aTaxable7_5(17) = 0
'                aTaxable7_5(18) = 0
''                aTaxable7_5(17) = aTaxable7_5(0)
'            End If


'            aTaxable7_5(10) = Round(aTaxable7_5(9) + aTaxable7_5(17) + aTaxable7_5(18), 2)
            aTaxable7_5(10) = Round(aTaxable7_5(0), 2)

            If Trim(oRecordSet("taxcode") <> "Z") Then
'                'old taxcode n amount
'                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
'                OpenQueryDNS cSqlStmt, oTaxSet, False
'
'    '            MsgBox right(oRecordSet("taxcode"), 1)
'                If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                    nTaxAmt = oTaxSet("S_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                    nTaxAmt = oTaxSet("H_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                    nTaxAmt = oTaxSet("M_AMT")
'                End If
'
'                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                    nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'                End If

                'new taxcode amount
                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
                OpenQueryDNS cSqlStmt, oTaxSet, False

                If InStr(oRecordSet("taxcode"), "S") = 1 Then
                    nTaxAmt2 = oTaxSet("S_AMT")
                End If

                If InStr(oRecordSet("taxcode"), "H") = 1 Then
                    nTaxAmt2 = oTaxSet("H_AMT")
                End If

                If InStr(oRecordSet("taxcode"), "M") = 1 Then
                    nTaxAmt2 = oTaxSet("M_AMT")
                End If

                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
                    nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
                End If
            Else
                nTaxAmt = 0
                nTaxAmt2 = 0
            End If


'            aTaxable7_5(11) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
            aTaxable7_5(11) = Round(nTaxAmt2, 2)


'            aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(10)) - aTaxable7_5(11), 2)
            aTaxable7_5(12) = 0


            'tax due
'            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
'                       " Where range1 <= " & aTaxable7_5(12) & " And range2 >=" & aTaxable7_5(12)
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(13) = Round(((aTaxable7_5(12) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
'            Else
'                aTaxable7_5(13) = 0
'            End If
            aTaxable7_5(13) = 0

            aTaxable7_5(14) = Round(aTaxable7_5(13) - oRecordSet("tax_wheld"), 2)
            aTaxable7_5(15) = Round(oRecordSet("tax_wheld") - aTaxable7_5(13), 2)

            aTaxable7_5(16) = Round(oRecordSet("tax_wheld") + aTaxable7_5(14), 2)

            cSqlStmt = "insert into alpha7_5(sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h,sched5i,sched5j,sched5k,sched5l, " & _
                       " sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v,sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad, " & _
                       " sched5ae,sched5af,sched5ag,sched6,sched6b,sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12,sched13)values(" & _
                       nCtr & "," & _
                       cQuote & oRecordSet("tin") & cQuote & "," & _
                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & "III" & cQuote & "," & _
                       "0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & Combo1.Text & "-12-31" & cQuote & "," & _
                       aTaxable7_5(0) & "," & _
                       aTaxable7_5(1) & "," & aTaxable7_5(2) & "," & aTaxable7_5(3) & ",312," & _
                       aTaxable7_5(4) & "," & aTaxable7_5(5) & "," & aTaxable7_5(6) & ",0," & _
                       aTaxable7_5(7) & ",0," & aTaxable7_5(8) & "," & aTaxable7_5(18) & ",0," & aTaxable7_5(17) & "," & aTaxable7_5(10) & "," & aTaxable7_5(10) & "," & _
                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
                       aTaxable7_5(11) & ",0," & _
                       aTaxable7_5(12) & "," & _
                       aTaxable7_5(13) & ",0," & _
                       oRecordSet("tax_wheld") & "," & _
                       aTaxable7_5(14) & "," & _
                       aTaxable7_5(15) & "," & _
                       aTaxable7_5(16) & "," & _
                       aTaxable7_5(19) & ")"
'                MsgBox cSqlStmt
            QueryDBF cSqlStmt, objdbRs, True

            nCtr = nCtr + 1
            oRecordSet.MoveNext
        Wend
    End If

    ShowProgress 4
    
    
    
'----------------------------------------------------------

''-----Contractual
'    ShowProgress 0
'
''    aTaxable7_5(0) = grosspay
''    aTaxable7_5(1) = basic per day
''    aTaxable7_5(2) = basic per month
''    aTaxable7_5(3) = basic per year
''    aTaxable7_5(4) = holiday
''    aTaxable7_5(5) = ot
''    aTaxable7_5(6) = night shift differential
''non-taxable
''    aTaxable7_5(7) = 13 month pay
''    aTaxable7_5(8) = nontax
''non-taxable
''    aTaxable7_5(9) = 13 month pay
'
''    aTaxable7_5(10) = total compensation
''    aTaxable7_5(11) = Tax Amount
''    aTaxable7_5(12) = Net taxable Compensation Income
''    aTaxable7_5(13) = Tax Due
'
''    aTaxable7_5(14) = Amount withheld and paid for in december
'
''    aTaxable7_5(15) = Over withheld tax refunded to employee
''    aTaxable7_5(16) = Amount of tax withheld as asjusted
''    aTaxable7_5(19) = basicpay + colapay + nd pay
'
'
'    aTaxable7_5 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    aTaxable = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'
'    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
'    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        While Not objdbRs.EOF
'            l13month = IIf(objdbRs("pclose") = 0, False, True)
'            objdbRs.MoveNext
'        Wend
'    End If
'
'    ' --> active employee minimun wage earner
'    nCtr = 1
'
'    cSqlStmt = " select a.cola_amt,a.paystatus,a.emp_stat,a.periodid,ifnull(b.rate_amt,0) as rate_amt,a.empid,ifnull(b.lastname,a.lastname) as lastname,ifnull(b.firstname,a.firstname) as firstname,ifnull(b.mname,a.mname) as mname, ifnull(b.tin,'') as tin,ifnull(c.taxcode,'S') as taxcode,ifnull(b.date_hire,'') as date_hire, ifnull(b.date_fin,'') as date_fin, " & _
'               " round(sum(ifnull(a.gross_pay,'')),2) as gross_pay, " & _
'               " round(sum(a.SA_NET_PAY),2) as gross_pay_sa, " & _
'               " round((sum(a.SUN_PAY) + sum(a.SUN_OT_PAY) + sum(a.SUN_COLA) + sum(a.SUN_ND_PAY) + sum(a.SUN_ND_OT_PAY)),2) as suntot, " & _
'               " round(sum(a.reg_pay) + sum(a.ndiff_pay),2) as basic, " & _
'               " round(sum(a.hol_pay),2) as hol_pay, " & _
'               " round(sum(a.ndiff_pay)+sum(a.ndiff_ot_pay)+sum(a.sa_ndiff_pay),2) as ndiff, " & _
'               " round(sum(a.reg_ot_pay),2) as ot, " & _
'               " round(sum(a.sa_reg_pay)+sum(a.SUN_OT_PAY),2) as sa_sun_ot_pay, " & _
'               " round(sum(ifnull(b.ytd_wtax,0)), 2) As tax_wheld " & _
'               " from pah87260 a left join di3670 b on a.empid=b.empid " & _
'               "  left join pa8290 c on b.taxid=c.taxid or a.taxid=c.taxid " & _
'               " where a.periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & " and 13month=0) and a.emp_stat<>2" & _
'               " group by a.empid order by a.emp_stat,a.active,a.fullname "
'
''               " and a.empid in ('025970','028346','028460','021029','021175')" & _
'
'
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
''            If oRecordSet("empid") = "0716" Then MsgBox "stop"
'            cSqlStmt = " select cola_amt,empid,rate_amt,round(sum(gross_pay),2) as gross_pay, " & _
'                       " round(SA_NET_PAY,2) as gross_pay_sa, " & _
'                       " round(SUN_PAY + SUN_OT_PAY + SUN_COLA + SUN_ND_PAY + SUN_ND_OT_PAY,2) as suntot, " & _
'                       " round(hol_pay,2) as hol_pay, " & _
'                       " round(reg_pay + ndiff_pay, 2) As basic, " & _
'                       " round(ndiff_pay + ndiff_ot_pay+sa_ndiff_pay,2) as ndiff, " & _
'                       " round(reg_ot_pay,2) as ot, " & _
'                       " round(sa_reg_pay+SUN_OT_PAY,2) as sa_sun_ot_pay, " & _
'                       " round(wtax,2) as tax_wheld " & _
'                       " From pa87260 " & _
'                       " where periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & " and 13month=0) " & _
'                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, oRSet, False
'
'            If oRSet.RecordCount > 0 Then
'                aTaxable7_5(1) = Round(oRSet("rate_amt") + oRSet("cola_amt"), 2)
'
'                aTaxable7_5(2) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 26, 2)
'                aTaxable7_5(3) = Round((oRSet("rate_amt") + oRecordSet("rate_amt") + oRSet("cola_amt") + oRecordSet("cola_amt")) * 312, 2)
'
''                aTaxable7_5(2) = Round(orecordset("basic"), 2)
''                aTaxable7_5(3) = Round(oRSet("basic") + orecordset("basic"), 2)
'
'                aTaxable7_5(4) = Round(oRSet("hol_pay") + oRecordSet("hol_pay"), 2)
'
''2010-01-19
''                If gCompanyID <> "0002" Then
''                    If gCompanyID = "0003" Then
''                        aTaxable7_5(5) = Round((oRSet("ot") + oRSet("sa_sun_ot_pay")) + (orecordset("ot") + orecordset("sa_sun_ot_pay")), 2)
''                    Else
''                        aTaxable7_5(5) = Round(oRSet("ot") + orecordset("ot"), 2)
''                    End If
''                Else
''                    aTaxable7_5(5) = Round((oRSet("gross_pay_sa") + orecordset("gross_pay_sa")) - oRSet("suntot") + orecordset("suntot"), 2)
''                End If
'
'                aTaxable7_5(6) = Round(oRSet("ndiff") + oRecordSet("ndiff"), 2)
'
'                If gCompanyID = "" Then
'                    aTaxable7_5(18) = (oRSet("gross_pay") + oRSet("gross_pay_sa") + oRecordSet("gross_pay") + oRecordSet("gross_pay_sa"))
'                    aTaxable7_5(5) = Round(oRSet("ot") + oRSet("sa_sun_ot_pay") + oRecordSet("ot") + oRecordSet("sa_sun_ot_pay"), 2)
'                Else
'                    aTaxable7_5(18) = (oRSet("gross_pay") + oRecordSet("gross_pay"))
'                    aTaxable7_5(5) = Round(oRSet("ot") + oRecordSet("ot"), 2)
'                End If
'                aTaxable7_5(19) = Round(oRSet("basic") + oRecordSet("basic") + oRSet("cola_amt") + oRecordSet("cola_amt"), 2)
'            Else
'
'                aTaxable7_5(1) = Round(oRecordSet("rate_amt") + oRecordSet("cola_amt"), 2)
'
'                aTaxable7_5(2) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 26, 2)
'                aTaxable7_5(3) = Round((oRecordSet("rate_amt") + oRecordSet("cola_amt")) * 312, 2)
'
''                aTaxable7_5(2) = Round(orecordset("basic"), 2)
''                aTaxable7_5(3) = Round(orecordset("basic"), 2)
'                aTaxable7_5(4) = Round(oRecordSet("hol_pay"), 2)
'
'
''                If gCompanyID <> "0002" Then
''                    If gCompanyID = "0003" Then
''                        aTaxable7_5(5) = Round(orecordset("ot") + orecordset("sa_sun_ot_pay"), 2)
''                    Else
''                        aTaxable7_5(5) = Round(orecordset("ot"), 2)
''                    End If
''                Else
''                    aTaxable7_5(5) = Round((orecordset("gross_pay_sa")) - orecordset("suntot"), 2)
''                End If
'
'                aTaxable7_5(5) = Round(oRecordSet("ot"), 2)
'
'                aTaxable7_5(6) = Round(oRecordSet("ndiff"), 2)
'
'                If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
'                    aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
'                    aTaxable7_5(5) = Round(oRecordSet("ot") + oRecordSet("sa_sun_ot_pay"), 2)
'
'                Else
'                    aTaxable7_5(18) = oRecordSet("gross_pay")
'                    aTaxable7_5(5) = Round(oRecordSet("ot"), 2)
'                End If
'            End If
'
'            aTaxable7_5(19) = Round(oRecordSet("basic") + oRecordSet("cola_amt"), 2)
'
''            If orecordset("emp_stat") = 2 Or orecordset("emp_stat") = 1 Then
''                If orecordset("paystatus") = 0 Then
''                    cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
''                               " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
''                               "      and (empid=" & cQuote & orecordset("empid") & cQuote & ")"
''        '            Script2File cSqlStmt
''        '            MsgBox cSqlStmt
''                    OpenQueryDNS cSqlStmt, objdbRs, False
''                    If objdbRs.RecordCount > 0 Then
''
''                        n13m_SLVL = Round(objdbRs("net_pay"), 2)
''
''                    Else
''
''                        cSqlStmt = "select net_pay from pah87260" & _
''                                   " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
''                                   "      and (empid=" & cQuote & orecordset("empid") & cQuote & ")"
'''                        MsgBox cSqlStmt
''                        OpenQueryDNS cSqlStmt, objdbRs, False
''                        If objdbRs.RecordCount > 0 Then
''
''                            n13m_SLVL = Round(objdbRs("net_pay"), 2)
''
''                        Else
''                            aTaxable7_5(7) = 0
''                            aTaxable7_5(9) = 0
''                        End If
''                    End If
''
''                    If n13m_SLVL <> 0 Then
''                       If n13m_SLVL > 30000 Then
''                            aTaxable7_5(7) = 30000
''                            aTaxable7_5(9) = n13m_SLVL - 30000
''                        Else
''                            aTaxable7_5(7) = n13m_SLVL
''                            aTaxable7_5(9) = 0
''                        End If
''                    Else
''                        aTaxable7_5(7) = 0
''                        aTaxable7_5(9) = 0
''                    End If
''                Else
''                    n13m_SLVL = 0
''                End If
''            Else
''                n13m_SLVL = 0
''            End If
'
'            cSqlStmt = "select empid,sum(m13pay) as  netpay  from pah87260" & _
'                       " where (periodid in (select periodid from pa7730 where  year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ") group by empid"
''            Script2File cSqlStmt
''            MsgBox cSqlStmt
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'
'                n13m_SLVL = Round(objdbRs("netpay"), 2)
'                If n13m_SLVL > 30000 Then
'                     aTaxable7_5(7) = 30000
'                     aTaxable7_5(9) = n13m_SLVL - 30000
'                Else
'                     aTaxable7_5(7) = n13m_SLVL
'                     aTaxable7_5(9) = 0
'                End If
'            Else
'                aTaxable7_5(7) = 0
'                aTaxable7_5(9) = 0
'            End If
'
'            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                       " From pah87263 " & _
'                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                       " (dedid in (" & cDedID & ")) " & _
'                       " and (periodid in (select periodid from pa7730 where year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       " group by empid "
''            Script2File cSqlStmt
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(8) = objdbRs("non_tax")
'            Else
'                aTaxable7_5(8) = 0
'            End If
'
'            'aTaxable7_5(0) = aTaxable7_5(3) + aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(7)
'            If (gCompanyID = "0003") Or (gCompanyID = "0007") Or (gCompanyID = "0002") Or (gCompanyID = "0004") Then
'                aTaxable7_5(0) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa") + aTaxable7_5(7)
'            Else
'                aTaxable7_5(0) = oRecordSet("gross_pay")
'            End If
'
''            aTaxable7_5(0) = bacis + Holpay + (Round((orecordset("gross_pay_sa")) - orecordset("suntot"), 2)) + _
''                            ndiff + a13month
'
''            aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'
'            'aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'            aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(8))
''                aTaxable7_5(18) = orecordset("gross_pay") - (excess13monthpay + deduction)
'
'            aTaxable7_5(17) = 0
'
''            cSqlStmt = "select * from pa7730 where month(date_start) > 6  and periodid=" & cQuote & orecordset("periodid") & cQuote
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                ' non-taxable july - december
'''                aTaxable7_5(18) = aTaxable7_5(0)
''                aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
''
'''                aTaxable7_5(18) = orecordset("gross_pay") - (excess13monthpay + deduction)
''
''
''                aTaxable7_5(17) = 0
''
''            Else
''                ' taxable January - June
''                '20100114
''                'aTaxable7_5(17) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
''                aTaxable7_5(17) = 0
''                aTaxable7_5(18) = 0
'''                aTaxable7_5(17) = aTaxable7_5(0)
''            End If
'
'
''            aTaxable7_5(10) = Round(aTaxable7_5(9) + aTaxable7_5(17) + aTaxable7_5(18), 2)
'            aTaxable7_5(10) = Round(aTaxable7_5(0), 2)
'
'            If Trim(oRecordSet("taxcode") <> "Z") Then
''                'old taxcode n amount
''                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
''                OpenQueryDNS cSqlStmt, oTaxSet, False
''
''    '            MsgBox right(oRecordSet("taxcode"), 1)
''                If InStr(oRecordSet("taxcode"), "S") = 1 Then
''                    nTaxAmt = oTaxSet("S_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "H") = 1 Then
''                    nTaxAmt = oTaxSet("H_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "M") = 1 Then
''                    nTaxAmt = oTaxSet("M_AMT")
''                End If
''
''                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
''                    nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
''                End If
'
'                'new taxcode amount
'                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
'                OpenQueryDNS cSqlStmt, oTaxSet, False
'
'                If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                    nTaxAmt2 = oTaxSet("S_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                    nTaxAmt2 = oTaxSet("H_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                    nTaxAmt2 = oTaxSet("M_AMT")
'                End If
'
'                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                    nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'                End If
'            Else
'                nTaxAmt = 0
'                nTaxAmt2 = 0
'            End If
'
'
''            aTaxable7_5(11) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
'            aTaxable7_5(11) = Round(nTaxAmt2, 2)
'
'
''            aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(10)) - aTaxable7_5(11), 2)
'            aTaxable7_5(12) = 0
'
'
'            'tax due
''            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
''                       " Where range1 <= " & aTaxable7_5(12) & " And range2 >=" & aTaxable7_5(12)
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                aTaxable7_5(13) = Round(((aTaxable7_5(12) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
''            Else
''                aTaxable7_5(13) = 0
''            End If
'            aTaxable7_5(13) = 0
'
'            aTaxable7_5(14) = Round(aTaxable7_5(13) - oRecordSet("tax_wheld"), 2)
'            aTaxable7_5(15) = Round(oRecordSet("tax_wheld") - aTaxable7_5(13), 2)
'
'            aTaxable7_5(16) = Round(oRecordSet("tax_wheld") + aTaxable7_5(14), 2)
'
'            cSqlStmt = "insert into alpha7_5(sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h,sched5i,sched5j,sched5k,sched5l, " & _
'                       " sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v,sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad, " & _
'                       " sched5ae,sched5af,sched5ag,sched6,sched6b,sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12,sched13)values(" & _
'                       nCtr & "," & _
'                       cQuote & oRecordSet("tin") & cQuote & "," & _
'                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & "III" & cQuote & "," & _
'                       "0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
'                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
'                       aTaxable7_5(0) & "," & _
'                       aTaxable7_5(1) & "," & aTaxable7_5(2) & "," & aTaxable7_5(3) & ",312," & _
'                       aTaxable7_5(4) & "," & aTaxable7_5(5) & "," & aTaxable7_5(6) & ",0," & _
'                       aTaxable7_5(7) & ",0," & aTaxable7_5(8) & "," & aTaxable7_5(18) & ",0," & aTaxable7_5(17) & "," & aTaxable7_5(10) & "," & aTaxable7_5(10) & "," & _
'                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_5(11) & ",0," & _
'                       aTaxable7_5(12) & "," & _
'                       aTaxable7_5(13) & ",0," & _
'                       oRecordSet("tax_wheld") & "," & _
'                       aTaxable7_5(14) & "," & _
'                       aTaxable7_5(15) & "," & _
'                       aTaxable7_5(16) & "," & _
'                       aTaxable7_5(19) & ")"
'
''                MsgBox cSqlStmt
'                QueryDBF cSqlStmt, objdbRs, True
'
'            nCtr = nCtr + 1
'            oRecordSet.MoveNext
'        Wend
'    End If

    ShowProgress 4

    MsgBox "done"

    Set oFileSys = Nothing
    Set oRecordSet = Nothing
    Set oRSet = Nothing
    Set oRSet2 = Nothing
    Set oTaxSet = Nothing
    Set oTaxSet = Nothing
   
End Sub

'Sub GenAlphaList(ByVal cPeriodID As String)
'    Dim cDedID As String, _
'        cSqlStmt As String, _
'        oFileSys As New FileSystemObject, _
'        oRecordSet As New ADODB.Recordset, _
'        oRSet As New ADODB.Recordset, _
'        oRSet2 As New ADODB.Recordset, _
'        oTaxSet As New ADODB.Recordset, _
'        nCtr As Integer, _
'        nYear As Integer, _
'        n13m_SLVL As Double, _
'        l13month As String, _
'        nTaxAmt As Double, _
'        nTaxAmt2 As Double, _
'        aTaxable7_3 As Variant, _
'        aTaxable7_5 As Variant, _
'        aTaxable As Variant, _
'        lIba As Boolean
'
'    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_1.DBF") Then
'        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_1.DBF"
'    End If
'    DetectDBF CheckPath(Text4.Text)
'
'    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_3.DBF") Then
'        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_3.DBF"
'    End If
'    DetectDBF CheckPath(Text4.Text)
'
'    If oFileSys.FileExists(CheckPath(Text4.Text) & "alpha7_5.DBF") Then
'        oFileSys.DeleteFile CheckPath(Text4.Text) & "alpha7_5.DBF"
'    End If
'    DetectDBF CheckPath(Text4.Text)
'
'    CreateAlphaLst
'
'    For nCtr = 0 To UBound(aTaxExempt)
'        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
'        cDedID = cDedID & aTaxExempt(nCtr) & ","
'    Next nCtr
'    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
'
'
''    aTaxable7_3(0) = gross_pay
''   non -Taxable
''    aTaxable7_3(1) = 13month
''    aTaxable7_3(2) = non_tax
''    aTaxable7_3(3) = total of non tax
''   Taxable
''    aTaxable7_3(4) = basic
''    aTaxable7_3(5) = 13month
''    aTaxable7_3(6) = total of taxable
''    aTaxable7_3(7) = Tax Amount
''    aTaxable7_3(8) = Net Taxable Compensation Income
''    aTaxable7_3(9) = Tax Due
''    aTaxable7_3(10) = Amount withheld and paid for in december
''    aTaxable7_3(11) = Over withheld tax refunded to employee
''    aTaxable7_3(12) = Amount of tax withheld as asjusted
'
'    ShowProgress 0
''7.1 regular
'
'    aTaxable7_3 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'
'    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
'    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        While Not objdbRs.EOF
'            l13month = IIf(objdbRs("pclose") = 0, False, True)
'            objdbRs.MoveNext
'        Wend
'    End If
'
' '    --> active employee
'    nCtr = 1
'    cSqlStmt = "select a.empid, ifnull(c.lastname,a.lastname) as lastname,ifnull(c.firstname,a.firstname) as firstname,ifnull(c.mname,a.mname) as mname, c.tin, ifnull(e.taxcode,'S') as taxcode , " & _
'               "  round(c.ytd_gross + a.gross_pay,2) as gross_pay, " & _
'               "  round(c.ytd_gross_sa + a.SA_NET_PAY,2) as gross_pay_sa, " & _
'               "  round((a.SUN_PAY + a.SUN_OT_PAY + a.SUN_COLA + a.SUN_ND_PAY + a.SUN_ND_OT_PAY),2) as suntot, " & _
'               "  round(c.ytd_basic,2) as basic, " & _
'               "  round(b.ded_amt3,2) as non_tax, " & _
'               "  (c.sl_avail - c.sl_use) + (c.vl_avail - c.vl_use) as leave_unuse, " & _
'               "  c.rate_amt, " & _
'               "  round(c.ytd_wtax,2) as tax_wheld, " & _
'               "  round(b.ded_amt2, 2) As adj_tax " & _
'               "from pa87260 a left join pa87263 b on a.periodid=b.periodid and a.empid=b.empid and b.dedid='006' " & _
'               "  left join di3670 c on a.empid=c.empid " & _
'               "  left join pa8290 e on c.taxid=e.taxid " & _
'               "where a.active=0 and a.emp_stat=2 and a.rate_amt <> 292 and a.periodid=" & cQuote & cPeriodID & cQuote & _
'               " order by a.fullname"
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
'
'            cSqlStmt = " select round(sum(SUN_PAY)+sum(SUN_OT_PAY)+sum(SUN_COLA)+ sum(SUN_ND_PAY)+sum(SUN_ND_OT_PAY),2) as suntot " & _
'                       " From pah87260 " & _
'                       " where periodid in ( " & _
'                       " select periodid from pa7730 " & _
'                       " where year(date_start)=2008 and 13month <> 1) and empid = " & cQuote & oRecordSet("empid") & cQuote
'            OpenQueryDNS cSqlStmt, objdbRs, False
'
'            If gCompanyID <> "0002" Then
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay"), 2)
'            Else
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay") + (oRecordSet("gross_pay_sa") - (oRecordSet("suntot") + objdbRs("suntot"))), 2)
'            End If
'
'            aTaxable7_3(2) = oRecordSet("non_tax")
'
''            If oRecordSet("empid") = "002758" Then MsgBox "stop"
'
'            If gCompanyID <> "0002" Then
'                If oRecordSet("non_tax") = 0 Then
'
'                    cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                               " From pah87263 " & _
'                               " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                               " (dedid in (" & cDedID & ")) " & _
'                               " and (periodid in (select periodid from pa7730 where year(date_end)=2008)) " & _
'                               " group by empid "
'
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'
'                    aTaxable7_3(2) = objdbRs("non_tax")
'
'                    cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                               " From pa87263 " & _
'                               " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                               " (dedid in (" & cDedID & ")) " & _
'                               " and periodid=" & cQuote & cPeriodID & cQuote & _
'                               " group by empid "
'
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'
'                    aTaxable7_3(2) = aTaxable7_3(2) + objdbRs("non_tax")
'                End If
'
'            Else
'                aTaxable7_3(2) = oRecordSet("non_tax")
'            End If
'
'            cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
'                       " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
''            Script2File cSqlStmt
''            MsgBox cSqlStmt
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'
'                n13m_SLVL = Round((objdbRs("net_pay") + (oRecordSet("leave_unuse") * oRecordSet("rate_amt"))), 2)
'
'                If n13m_SLVL > 30000 Then
'                    aTaxable7_3(1) = 30000
'                    aTaxable7_3(5) = n13m_SLVL - 30000
'                Else
'                    aTaxable7_3(1) = n13m_SLVL
'                    aTaxable7_3(5) = 0
'                End If
'            End If
'
'            aTaxable7_3(0) = aTaxable7_3(0) + n13m_SLVL
'
'            'total of non tax
'            aTaxable7_3(3) = Round(aTaxable7_3(1) + aTaxable7_3(2), 2)
'
'            'aTaxable7_3(6) = total of taxable
'            aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))
'            aTaxable7_3(6) = Round(aTaxable7_3(4) + aTaxable7_3(5), 2)
'
''            'old taxcode n amount
''            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
''            OpenQueryDNS cSqlStmt, oTaxSet, False
''
''            If InStr(oRecordSet("taxcode"), "S") = 1 Then
''                nTaxAmt = oTaxSet("S_AMT")
''            End If
''
''            If InStr(oRecordSet("taxcode"), "H") = 1 Then
''                nTaxAmt = oTaxSet("H_AMT")
''            End If
''
''            If InStr(oRecordSet("taxcode"), "M") = 1 Then
''                nTaxAmt = oTaxSet("M_AMT")
''            End If
''
''            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
''                nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
''            End If
'
'
'            'new taxcode amount
'            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
'            OpenQueryDNS cSqlStmt, oTaxSet, False
'
'            If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                nTaxAmt2 = oTaxSet("S_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                nTaxAmt2 = oTaxSet("H_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                nTaxAmt2 = oTaxSet("M_AMT")
'            End If
'
'            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'            End If
'
'            'aTaxable7_3(7) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
'            aTaxable7_3(7) = Round(nTaxAmt2 / 2, 2)
'
'            aTaxable7_3(8) = Round(aTaxable7_3(6) - aTaxable7_3(7), 2)
'
'            'tax due
'            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
'                       " Where range1 <= " & aTaxable7_3(8) & " And range2 >=" & aTaxable7_3(8)
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_3(9) = Round(((aTaxable7_3(8) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
'            Else
'                aTaxable7_3(9) = 0
'            End If
'
'            aTaxable7_3(10) = Round(aTaxable7_3(9) - oRecordSet("tax_wheld"), 2)
'            aTaxable7_3(11) = Round(oRecordSet("tax_wheld") - aTaxable7_3(9), 2)
'
'            aTaxable7_3(12) = Round(oRecordSet("tax_wheld") + aTaxable7_3(10), 2)
'
'            cSqlStmt = "insert into alpha7_3(sched1,sched2,sched3a,sched3b,sched3c,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g, " & _
'                       "sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12)values(" & _
'                       nCtr & "," & _
'                       cQuote & oRecordSet("tin") & cQuote & "," & _
'                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & _
'                       aTaxable7_3(0) & "," & aTaxable7_3(1) & ",0," & aTaxable7_3(2) & ",0," & aTaxable7_3(3) & "," & aTaxable7_3(4) & "," & aTaxable7_3(5) & ",0," & aTaxable7_3(6) & "," & _
'                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_3(7) & ",0," & _
'                       aTaxable7_3(8) & "," & _
'                       aTaxable7_3(9) & "," & _
'                       oRecordSet("tax_wheld") & "," & _
'                       aTaxable7_3(10) & "," & aTaxable7_3(11) & "," & aTaxable7_3(12) & "," & _
'                       cQuote & "N" & cQuote & ")"
'
''                MsgBox cSqlStmt
'                QueryDBF cSqlStmt, objdbRs, True
'
'            nCtr = nCtr + 1
'            oRecordSet.MoveNext
'        Wend
'    End If
'
'    ShowProgress 4
'
'
'    ShowProgress 0
'    '7.1 regular
'      ' --> resigned employee
'
'    nCtr = 1
'
'    cSqlStmt = " select a.empid,a.date_hire,a.date_res, " & _
'               " a.lastname,a.firstname,a.mname, " & _
'               " a.tin,ifnull(b.taxcode,a.taxcode) as taxcode, " & _
'               " round(a.ytd_gross,2) as gross_pay, " & _
'               " round(a.ytd_gross_sa,2) as gross_pay_sa, " & _
'               " round(a.ytd_basic,2) as basic, " & _
'               " (a.sl_avail - a.sl_use) + (a.vl_avail - a.vl_use) as leave_unuse, " & _
'               " a.rate_amt, " & _
'               " a.ytd_wtax as tax_wheld, " & _
'               " sum(c.ded_amt) As ded_amt " & _
'               "from di3670 a left join pa8290 b on a.taxid=b.taxid " & _
'               "  left join pah87263 c on a.empid=c.empid and (c.periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ")) " & _
'               "Where (a.emp_stat = 2) " & _
'               "  and (((a.active=1) and (year(a.date_res)=" & Combo1.Text & "))" & _
'               "       or ((a.active=2) and (year(a.date_fin)=" & Combo1.Text & ")))" & _
'               "  and (c.dedid in (" & cDedID & ")) " & _
'               "group by c.empid " & _
'               " order by a.lastname,a.firstname,a.mname "
''    Script2File cSqlStmt
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
'
'            cSqlStmt = " SELECT round(sum(SUN_PAY) + sum(SUN_OT_PAY) + sum(SUN_COLA) + sum(SUN_ND_PAY) + sum(SUN_ND_OT_PAY),2) as suntot " & _
'                       " From pah87260 " & _
'                       " where periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ") " & _
'                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, objdbRs, False
'
'            If gCompanyID <> "0002" Then
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay"), 2)
'
'            Else
'                aTaxable7_3(0) = Round(oRecordSet("gross_pay") + (oRecordSet("gross_pay_sa") - objdbRs("suntot")), 2)
'            End If
'
'            aTaxable7_3(2) = oRecordSet("ded_amt")
'
'            cSqlStmt = "select sum(leave_pay) as leave_pay, sum(m13pay) as 13mopay " & _
'                       "From pah87260 " & _
'                       "where (periodid in (select periodid from pa7730 where year(date_end)=" & Combo1.Text & ")) and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'
'                n13m_SLVL = Round((objdbRs("leave_pay") + objdbRs("13mopay")), 2)
'
'                If n13m_SLVL > 30000 Then
'                    aTaxable7_3(1) = 30000
'                    aTaxable7_3(5) = n13m_SLVL - 30000
'                Else
'                    aTaxable7_3(1) = n13m_SLVL
'                    aTaxable7_3(5) = 0
'                End If
'            End If
'
'            aTaxable7_3(0) = aTaxable7_3(0) + n13m_SLVL
'
'            'total of non tax
'            aTaxable7_3(3) = Round(aTaxable7_3(1) + aTaxable7_3(2), 2)
'
'            'aTaxable7_3(6) = total of taxable
'            aTaxable7_3(4) = aTaxable7_3(0) - (aTaxable7_3(1) + aTaxable7_3(2))
'            aTaxable7_3(6) = Round(aTaxable7_3(4) + aTaxable7_3(5), 2)
'
'            If Trim(oRecordSet("taxcode") <> "") Then
''                'old taxcode n amount
''                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
''                OpenQueryDNS cSqlStmt, oTaxSet, False
''
''    '            MsgBox right(oRecordSet("taxcode"), 1)
''                If InStr(oRecordSet("taxcode"), "S") = 1 Then
''                    nTaxAmt = oTaxSet("S_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "H") = 1 Then
''                    nTaxAmt = oTaxSet("H_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "M") = 1 Then
''                    nTaxAmt = oTaxSet("M_AMT")
''                End If
''
''                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
''                    nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
''                End If
'
'                'new taxcode amount
'                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
'                OpenQueryDNS cSqlStmt, oTaxSet, False
'
'                If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                    nTaxAmt2 = oTaxSet("S_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                    nTaxAmt2 = oTaxSet("H_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                    nTaxAmt2 = oTaxSet("M_AMT")
'                End If
'
'                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                    nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'                End If
'            Else
'                nTaxAmt = 0
'                nTaxAmt2 = 0
'            End If
''            If Year(oRecordSet("date_fin")) = Combo1.Text Then
''                If Month(oRecordSet("date_fin")) < 7 Then
''                    aTaxable7_3(7) = Round(nTaxAmt / 2, 2)
''                Else
''                    aTaxable7_3(7) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
''                End If
''            Else
''                aTaxable7_3(7) = Round(nTaxAmt / 2, 2)
''            End If
'
'            'aTaxable7_3(7) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
'            aTaxable7_3(7) = Round(nTaxAmt2 / 2, 2)
'
'            aTaxable7_3(8) = Round(aTaxable7_3(6) - aTaxable7_3(7), 2)
'
'            'tax due
'            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
'                       " Where range1 <= " & aTaxable7_3(8) & " And range2 >=" & aTaxable7_3(8)
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_3(9) = Round(((aTaxable7_3(8) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
'            Else
'                aTaxable7_3(9) = 0
'            End If
'
'            aTaxable7_3(10) = Round(aTaxable7_3(9) - oRecordSet("tax_wheld"), 2)
'            aTaxable7_3(11) = Round(oRecordSet("tax_wheld") - aTaxable7_3(9), 2)
'
'            aTaxable7_3(12) = Round(oRecordSet("tax_wheld") + aTaxable7_3(10), 2)
'
'            cSqlStmt = "insert into alpha7_1(sched1,sched2,sched3a,sched3b,sched3c,sched3d,sched3e,sched4a,sched4b,sched4c,sched4d,sched4e,sched4f,sched4g, " & _
'                       "sched4h,sched4i,sched4j,sched5a,sched5b,sched6,sched7,sched8,sched9,sched10a,sched10b,sched11,sched12)values(" & _
'                       nCtr & "," & _
'                       cQuote & oRecordSet("tin") & cQuote & "," & _
'                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & _
'                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
'                       cQuote & Format(oRecordSet("date_res"), "yyyy-mm-dd") & cQuote & "," & _
'                       aTaxable7_3(0) & "," & aTaxable7_3(1) & ",0," & oRecordSet("ded_amt") & ",0," & aTaxable7_3(3) & "," & aTaxable7_3(4) & "," & aTaxable7_3(5) & ",0," & aTaxable7_3(6) & "," & _
'                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_3(7) & ",0," & _
'                       aTaxable7_3(8) & "," & _
'                       aTaxable7_3(9) & "," & _
'                       oRecordSet("tax_wheld") & "," & _
'                       aTaxable7_3(10) & "," & aTaxable7_3(11) & "," & aTaxable7_3(12) & "," & _
'                       cQuote & "N" & cQuote & ")"
'
''                MsgBox cSqlStmt
'                QueryDBF cSqlStmt, objdbRs, True
'
'            nCtr = nCtr + 1
'            oRecordSet.MoveNext
'        Wend
'    End If
'
'    ShowProgress 4
'
'    ShowProgress 0
''7.5 regular
''    aTaxable7_5(0) = grosspay
''    aTaxable7_5(1) = basic per day
''    aTaxable7_5(2) = basic per month
''    aTaxable7_5(3) = basic per year
''    aTaxable7_5(4) = holiday
''    aTaxable7_5(5) = ot
''    aTaxable7_5(6) = night shift differential
''non-taxable
''    aTaxable7_5(7) = 13 month pay
''    aTaxable7_5(8) = nontax
''non-taxable
''    aTaxable7_5(9) = 13 month pay
'
''    aTaxable7_5(10) = total compensation
''    aTaxable7_5(11) = Tax Amount
''    aTaxable7_5(12) = Net taxable Compensation Income
''    aTaxable7_5(13) = Tax Due
'
''    aTaxable7_5(14) = Amount withheld and paid for in december
'
''    aTaxable7_5(15) = Over withheld tax refunded to employee
''    aTaxable7_5(16) = Amount of tax withheld as asjusted
''    aTaxable7_5(17) = Salaries and Other forms of compensation
'
'
'    aTaxable7_5 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    aTaxable = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'
'    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
'    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        While Not objdbRs.EOF
'            l13month = IIf(objdbRs("pclose") = 0, False, True)
'            objdbRs.MoveNext
'        Wend
'    End If
'
'    ' --> active employee minimun wage earner
'    nCtr = 1
'    cSqlStmt = "select a.empid, ifnull(c.lastname,a.lastname) as lastname,ifnull(c.firstname,a.firstname) as firstname,ifnull(c.mname,a.mname) as mname, c.tin, ifnull(e.taxcode,'S') as taxcode, c.date_hire, c.date_fin, " & _
'               "  round(a.gross_pay,2) as gross_pay, " & _
'               "  round(a.SA_NET_PAY,2) as gross_pay_sa, " & _
'               "  round((a.SUN_PAY + a.SUN_OT_PAY + a.SUN_COLA + a.SUN_ND_PAY + a.SUN_ND_OT_PAY),2) as suntot, " & _
'               "  round(a.reg_pay + a.ndiff_pay,2) as basic, " & _
'               "  round(a.hol_pay,2) as hol_pay, " & _
'               "  round(a.ndiff_pay+a.ndiff_ot_pay,2) as ndiff, " & _
'               "  round(b.ded_amt3,2) as non_tax, " & _
'               "  round(a.reg_ot_pay,2) as ot, " & _
'               "  round(c.ytd_wtax,2) as tax_wheld, " & _
'               "  (c.sl_avail - c.sl_use) + (c.vl_avail - c.vl_use) as leave_unuse, " & _
'               "  c.rate_amt " & _
'               "from pa87260 a left join pa87263 b on a.periodid=b.periodid and a.empid=b.empid and b.dedid='006' " & _
'               "  left join di3670 c on a.empid=c.empid " & _
'               "  left join pa8290 e on c.taxid=e.taxid " & _
'               "where a.active=0 and a.emp_stat=2 and a.periodid=" & cQuote & cPeriodID & cQuote & " and c.rate_amt = " & gBasicRate & _
'               " order by a.fullname"
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
'
'            cSqlStmt = " select empid,rate_amt,round(sum(gross_pay),2) as gross_pay, " & _
'                       " round(sum(SA_NET_PAY),2) as gross_pay_sa, " & _
'                       " round((sum(SUN_PAY) + sum(SUN_OT_PAY) + sum(SUN_COLA) + sum(SUN_ND_PAY) + sum(SUN_ND_OT_PAY)),2) as suntot, " & _
'                       " round(sum(hol_pay),2) as hol_pay, " & _
'                       " round(Sum(reg_pay) + Sum(ndiff_pay), 2) As basic, " & _
'                       " round(sum(ndiff_pay)+sum(ndiff_ot_pay),2) as ndiff, " & _
'                       " round(sum(reg_ot_pay),2) as ot, " & _
'                       " round(sum(wtax),2) as tax_wheld " & _
'                       " From pah87260 " & _
'                       " where periodid in (select periodid from pa7730 where year(date_start) = 2008 and 13month = 0) " & _
'                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
'                       " group by empid "
'
'            OpenQueryDNS cSqlStmt, oRSet, False
'
'            aTaxable7_5(1) = Round(oRSet("rate_amt"), 2)
'            aTaxable7_5(2) = Round(oRSet("basic") + oRecordSet("basic") / 6, 2)
'            aTaxable7_5(3) = Round(oRSet("basic") + oRecordSet("basic"), 2)
'            aTaxable7_5(4) = Round(oRSet("hol_pay") + oRecordSet("hol_pay"), 2)
'
'            If gCompanyID <> "0002" Then
'                aTaxable7_5(5) = Round(oRSet("ot") + oRecordSet("ot"), 2)
'            Else
'                aTaxable7_5(5) = Round((oRSet("gross_pay_sa") + oRecordSet("gross_pay_sa")) - oRSet("suntot") + oRecordSet("suntot"), 2)
'            End If
'
'            aTaxable7_5(6) = Round(oRSet("ndiff") + oRecordSet("ndiff"), 2)
'
'            cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
'                       " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                       "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
'
''            MsgBox cSqlStmt
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'
'                n13m_SLVL = Round((objdbRs("net_pay") + (oRecordSet("leave_unuse") * oRecordSet("rate_amt"))), 2)
'
'                If n13m_SLVL > 30000 Then
'                    aTaxable7_5(7) = 30000
'                    aTaxable7_5(9) = n13m_SLVL - 30000
'                Else
'                    aTaxable7_5(7) = n13m_SLVL
'                   aTaxable7_5(9) = 0
'                End If
'            End If
'
'            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                       " From pah87263 " & _
'                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                       " (dedid in (" & cDedID & ")) " & _
'                       " and (periodid in (select periodid from pa7730 where year(date_end)=2008)) " & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, objdbRs, False
'
'            aTaxable7_5(8) = objdbRs("non_tax")
'
'            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                       " From pa87263 " & _
'                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                       " (dedid in (" & cQuote & "001" & cQuote & "," & cQuote & "003" & cQuote & "," & cQuote & "005" & cQuote & "," & cQuote & "007" & cQuote & ")) " & _
'                       " and periodid=" & cQuote & cPeriodID & cQuote & _
'                       " group by empid "
'
'            OpenQueryDNS cSqlStmt, objdbRs, False
'
'
'            aTaxable7_5(8) = aTaxable7_5(8) + objdbRs("non_tax")
'
'            aTaxable7_5(0) = (aTaxable7_5(3) + aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(7)) - aTaxable7_5(8)
'
'
'            cSqlStmt = " select empid,rate_amt,round(sum(gross_pay),2) as gross_pay, " & _
'                       " round(sum(SA_NET_PAY),2) as gross_pay_sa, " & _
'                       " round((sum(SUN_PAY) + sum(SUN_OT_PAY) + sum(SUN_COLA) + sum(SUN_ND_PAY) + sum(SUN_ND_OT_PAY)),2) as suntot, " & _
'                       " round(sum(hol_pay),2) as hol_pay, " & _
'                       " round(Sum(reg_pay) + Sum(ndiff_pay), 2) As basic, " & _
'                       " round(sum(ndiff_pay)+sum(ndiff_ot_pay),2) as ndiff, " & _
'                       " round(sum(reg_ot_pay),2) as ot, " & _
'                       " round(sum(wtax),2) as tax_wheld " & _
'                       " From pah87260 " & _
'                       " where periodid in (select periodid from pa7730 where year(date_start) = 2008 and month(date_start) < 7 and 13month = 0) " & _
'                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, oRSet2, False
'            If oRSet2.RecordCount > 0 Then
'                aTaxable7_5(17) = Round((oRSet2("gross_pay") + (oRSet2("gross_pay_sa")) - oRSet2("suntot")), 2)
'            Else
'                aTaxable7_5(17) = 0
'            End If
'            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                       " From pah87263 " & _
'                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                       " (dedid in (" & cDedID & ")) " & _
'                       " and (periodid in (select periodid from pa7730 where year(date_end)=2008 and month(date_end)<7 )) " & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(17) = (aTaxable7_5(17) + aTaxable7_5(9)) - objdbRs("non_tax")
'            Else
'                aTaxable7_5(17) = (aTaxable7_5(17) + aTaxable7_5(9))
'
'            End If
'
'            aTaxable7_5(10) = Round(aTaxable7_5(0) + aTaxable7_5(9) + aTaxable7_5(17), 2)
'
''            'old taxcode n amount
''            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
''            OpenQueryDNS cSqlStmt, oTaxSet, False
''
'''            MsgBox right(oRecordSet("taxcode"), 1)
''            If InStr(oRecordSet("taxcode"), "S") = 1 Then
''                nTaxAmt = oTaxSet("S_AMT")
''            End If
''
''            If InStr(oRecordSet("taxcode"), "H") = 1 Then
''                nTaxAmt = oTaxSet("H_AMT")
''            End If
''
''            If InStr(oRecordSet("taxcode"), "M") = 1 Then
''                nTaxAmt = oTaxSet("M_AMT")
''            End If
''
''            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
''                nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
''            End If
'
'            'new taxcode amount
'            cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
'            OpenQueryDNS cSqlStmt, oTaxSet, False
'
'            If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                nTaxAmt2 = oTaxSet("S_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                nTaxAmt2 = oTaxSet("H_AMT")
'            End If
'
'            If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                nTaxAmt2 = oTaxSet("M_AMT")
'            End If
'
'            If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'            End If
'
''            aTaxable7_5(11) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
'            aTaxable7_5(11) = Round(nTaxAmt2 / 2, 2)
'
'
'            'aTaxable7_5(12) = round(aTaxable7_5(10) - aTaxable7_5(11), 2)
'            'aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(17)) - aTaxable7_5(11), 2)
'            aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(10)) - aTaxable7_5(11), 2)
'
'
'            'tax due
'            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
'                       " Where range1 <= " & aTaxable7_5(12) & " And range2 >=" & aTaxable7_5(12)
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(13) = Round(((aTaxable7_5(12) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
'            Else
'                aTaxable7_5(13) = 0
'            End If
'
'            aTaxable7_5(14) = Round(aTaxable7_5(13) - oRecordSet("tax_wheld"), 2)
'            aTaxable7_5(15) = Round(oRecordSet("tax_wheld") - aTaxable7_5(13), 2)
'
'            aTaxable7_5(16) = Round(oRecordSet("tax_wheld") + aTaxable7_5(14), 2)
'
'            cSqlStmt = "insert into alpha7_5(sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h,sched5i,sched5j,sched5k,sched5l, " & _
'                       " sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v,sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad, " & _
'                       " sched5ae,sched5af,sched5ag,sched6,sched6b,sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12)values(" & _
'                       nCtr & "," & _
'                       cQuote & oRecordSet("tin") & cQuote & "," & _
'                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & cQuote & "," & _
'                       "0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
'                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & cQuote & "," & _
'                       aTaxable7_5(0) & "," & _
'                       aTaxable7_5(1) & "," & aTaxable7_5(2) & "," & aTaxable7_5(3) & ",312," & _
'                       aTaxable7_5(4) & "," & aTaxable7_5(5) & "," & aTaxable7_5(6) & ",0," & _
'                       aTaxable7_5(7) & ",0," & aTaxable7_5(8) & ",0," & aTaxable7_5(9) & "," & aTaxable7_5(17) & "," & aTaxable7_5(10) & "," & aTaxable7_5(10) & "," & _
'                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_5(11) & ",0," & _
'                       aTaxable7_5(12) & "," & _
'                       aTaxable7_5(13) & ",0," & _
'                       oRecordSet("tax_wheld") & "," & _
'                       aTaxable7_5(14) & "," & _
'                       aTaxable7_5(15) & "," & _
'                       aTaxable7_5(16) & ")"
''                MsgBox cSqlStmt
'                QueryDBF cSqlStmt, objdbRs, True
'            nCtr = nCtr + 1
'            oRecordSet.MoveNext
'        Wend
'    End If
'
'    ShowProgress 4
'
'
'
'    ShowProgress 0
'    '7.5 Contractual
'
''    aTaxable7_5(0) = grosspay
''    aTaxable7_5(1) = basic per day
''    aTaxable7_5(2) = basic per month
''    aTaxable7_5(3) = basic per year
''    aTaxable7_5(4) = holiday
''    aTaxable7_5(5) = ot
''    aTaxable7_5(6) = night shift differential
''non-taxable
''    aTaxable7_5(7) = 13 month pay
''    aTaxable7_5(8) = nontax
''non-taxable
''    aTaxable7_5(9) = 13 month pay
'
''    aTaxable7_5(10) = total compensation
''    aTaxable7_5(11) = Tax Amount
''    aTaxable7_5(12) = Net taxable Compensation Income
''    aTaxable7_5(13) = Tax Due
'
''    aTaxable7_5(14) = Amount withheld and paid for in december
'
''    aTaxable7_5(15) = Over withheld tax refunded to employee
''    aTaxable7_5(16) = Amount of tax withheld as asjusted
''    aTaxable7_5(17) = Salaries and Other forms of compensation
'
'
'    aTaxable7_5 = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'    aTaxable = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
'
'    cSqlStmt = " SELECT * FROM pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote
'    OpenQueryDNS cSqlStmt, objdbRs, False
'    If objdbRs.RecordCount > 0 Then
'        While Not objdbRs.EOF
'            l13month = IIf(objdbRs("pclose") = 0, False, True)
'            objdbRs.MoveNext
'        Wend
'    End If
'
'    ' --> active employee minimun wage earner
'    nCtr = 1
'
'    cSqlStmt = " select a.paystatus,a.emp_stat,a.periodid,ifnull(b.rate_amt,0) as rate_amt,a.empid,ifnull(b.lastname,a.lastname) as lastname,ifnull(b.firstname,a.firstname) as firstname,ifnull(b.mname,a.mname) as mname, ifnull(b.tin,'') as tin,ifnull(c.taxcode,'S') as taxcode,ifnull(b.date_hire,'') as date_hire, ifnull(b.date_fin,'') as date_fin, " & _
'               " round(sum(ifnull(a.gross_pay,'')),2) as gross_pay, " & _
'               " round(sum(a.SA_NET_PAY),2) as gross_pay_sa, " & _
'               " round((sum(a.SUN_PAY) + sum(a.SUN_OT_PAY) + sum(a.SUN_COLA) + sum(a.SUN_ND_PAY) + sum(a.SUN_ND_OT_PAY)),2) as suntot, " & _
'               " round(sum(a.reg_pay) + sum(a.ndiff_pay),2) as basic, " & _
'               " round(sum(a.hol_pay),2) as hol_pay, " & _
'               " round(sum(a.ndiff_pay)+sum(a.ndiff_ot_pay),2) as ndiff, " & _
'               " round(sum(a.reg_ot_pay),2) as ot, " & _
'               " round(sum(a.sa_reg_pay)+sum(a.SUN_OT_PAY),2) as sa_sun_ot_pay, " & _
'               " round(sum(ifnull(b.ytd_wtax,0)), 2) As tax_wheld " & _
'               " from pah87260 a left join di3670 b on a.empid=b.empid " & _
'               "  left join pa8290 c on b.taxid=c.taxid or a.taxid=c.taxid " & _
'               " where a.periodid in (select periodid from pa7730 where year(date_start)=" & Combo1.Text & " and 13month=0) and a.emp_stat<>2" & _
'               " group by a.empid order by a.emp_stat,a.active,a.fullname "
'
'    OpenQueryDNS cSqlStmt, oRecordSet, False
'    If oRecordSet.RecordCount > 0 Then
'        While Not oRecordSet.EOF
'
'            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Retrieving data of " & oRecordSet("Lastname")
'
'            cSqlStmt = " select empid,rate_amt,round(sum(gross_pay),2) as gross_pay, " & _
'                       " round(SA_NET_PAY,2) as gross_pay_sa, " & _
'                       " round(SUN_PAY + SUN_OT_PAY + SUN_COLA + SUN_ND_PAY + SUN_ND_OT_PAY,2) as suntot, " & _
'                       " round(hol_pay,2) as hol_pay, " & _
'                       " round(reg_pay + ndiff_pay, 2) As basic, " & _
'                       " round(ndiff_pay + ndiff_ot_pay,2) as ndiff, " & _
'                       " round(reg_ot_pay,2) as ot, " & _
'                       " round(sa_reg_pay+SUN_OT_PAY,2) as sa_sun_ot_pay, " & _
'                       " round(wtax,2) as tax_wheld " & _
'                       " From pa87260 " & _
'                       " where periodid in (select periodid from pa7730 where year(date_start) = 2008 and 13month = 0) " & _
'                       " and empid = " & cQuote & oRecordSet("empid") & cQuote & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, oRSet, False
'
'            If oRSet.RecordCount > 0 Then
'                aTaxable7_5(1) = Round(oRSet("rate_amt"), 2)
'                aTaxable7_5(2) = Round(oRSet("basic") + oRecordSet("basic") / 6, 2)
'
'                aTaxable7_5(3) = Round(oRSet("basic") + oRecordSet("basic"), 2)
'
'                aTaxable7_5(4) = Round(oRSet("hol_pay") + oRecordSet("hol_pay"), 2)
'
'
'                If gCompanyID <> "0002" Then
'                    If gCompanyID = "0003" Then
'                        aTaxable7_5(5) = Round((oRSet("ot") + oRSet("sa_sun_ot_pay")) + (oRecordSet("ot") + oRecordSet("sa_sun_ot_pay")), 2)
'                    Else
'                        aTaxable7_5(5) = Round(oRSet("ot") + oRecordSet("ot"), 2)
'                    End If
'                Else
'                    aTaxable7_5(5) = Round((oRSet("gross_pay_sa") + oRecordSet("gross_pay_sa")) - oRSet("suntot") + oRecordSet("suntot"), 2)
'                End If
'
'                aTaxable7_5(6) = Round(oRSet("ndiff") + oRecordSet("ndiff"), 2)
'                aTaxable7_5(18) = (oRSet("gross_pay") + oRecordSet("gross_pay")) + (oRSet("gross_pay_sa") + oRecordSet("gross_pay_sa"))
'
'            Else
'
'                aTaxable7_5(1) = Round(oRecordSet("rate_amt"), 2)
'                aTaxable7_5(2) = Round(oRecordSet("basic") / 6, 2)
'                aTaxable7_5(3) = Round(oRecordSet("basic"), 2)
'                aTaxable7_5(4) = Round(oRecordSet("hol_pay"), 2)
'
'
'                If gCompanyID <> "0002" Then
'                    If gCompanyID = "0003" Then
'                        aTaxable7_5(5) = Round(oRecordSet("ot") + oRecordSet("sa_sun_ot_pay"), 2)
'                    Else
'                        aTaxable7_5(5) = Round(oRecordSet("ot"), 2)
'                    End If
'                Else
'                    aTaxable7_5(5) = Round((oRecordSet("gross_pay_sa")) - oRecordSet("suntot"), 2)
'                End If
'
'                aTaxable7_5(6) = Round(oRecordSet("ndiff"), 2)
'                aTaxable7_5(18) = oRecordSet("gross_pay") + oRecordSet("gross_pay_sa")
'
'            End If
'
'            If oRecordSet("emp_stat") = 2 Or oRecordSet("emp_stat") = 1 Then
'                If oRecordSet("paystatus") = 0 Then
'                    cSqlStmt = "select net_pay from " & IIf(l13month = "False", "pa87260", "pah87260") & _
'                               " where (periodid in (select periodid from pa7730 where 13month=1 and year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                               "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
'        '            Script2File cSqlStmt
'        '            MsgBox cSqlStmt
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'                    If objdbRs.RecordCount > 0 Then
'
'                        n13m_SLVL = Round(objdbRs("net_pay"), 2)
'
'                    Else
'
'                        cSqlStmt = "select net_pay from pah87260" & _
'                                   " where (periodid in (select periodid from pa7730 where 13month=1 year(date_end)=" & cQuote & Combo1.Text & cQuote & ")) " & _
'                                   "      and (empid=" & cQuote & oRecordSet("empid") & cQuote & ")"
'                        OpenQueryDNS cSqlStmt, objdbRs, False
'                        If objdbRs.RecordCount > 0 Then
'
'                            n13m_SLVL = Round(objdbRs("net_pay"), 2)
'
'                        Else
'                            aTaxable7_5(7) = 0
'                            aTaxable7_5(9) = 0
'                        End If
'                    End If
'
'                    If n13m_SLVL <> 0 Then
'                       If n13m_SLVL > 30000 Then
'                            aTaxable7_5(7) = 30000
'                            aTaxable7_5(9) = n13m_SLVL - 30000
'                        Else
'                            aTaxable7_5(7) = n13m_SLVL
'                            aTaxable7_5(9) = 0
'                        End If
'                    Else
'                        aTaxable7_5(7) = 0
'                        aTaxable7_5(9) = 0
'                    End If
'                Else
'                    n13m_SLVL = 0
'                End If
'            Else
'                n13m_SLVL = 0
'            End If
'
'            cSqlStmt = " select dedid, sum(ded_amt) as non_tax " & _
'                       " From pah87263 " & _
'                       " where (empid=" & cQuote & oRecordSet("empid") & cQuote & ") and " & _
'                       " (dedid in (" & cDedID & ")) " & _
'                       " and (periodid in (select periodid from pa7730 where year(date_end)=2008)) " & _
'                       " group by empid "
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(8) = objdbRs("non_tax")
'            Else
'                aTaxable7_5(8) = 0
'            End If
'
'            aTaxable7_5(0) = aTaxable7_5(3) + aTaxable7_5(4) + aTaxable7_5(5) + aTaxable7_5(6) + aTaxable7_5(7)
'
''            aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'
'
'
'            cSqlStmt = "select * from pa7730 where month(date_start) > 6  and periodid=" & cQuote & oRecordSet("periodid") & cQuote
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                ' non-taxable july - december
''                aTaxable7_5(18) = aTaxable7_5(0)
'                aTaxable7_5(18) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'                aTaxable7_5(17) = 0
'
'            Else
'                ' taxable January - June
'                aTaxable7_5(17) = aTaxable7_5(18) - (aTaxable7_5(7) + aTaxable7_5(8))
'                aTaxable7_5(18) = 0
''                aTaxable7_5(17) = aTaxable7_5(0)
'            End If
'
'
'            aTaxable7_5(10) = Round(aTaxable7_5(0) + aTaxable7_5(9) + aTaxable7_5(17) + aTaxable7_5(18), 2)
'
'            If Trim(oRecordSet("taxcode") <> "Z") Then
''                'old taxcode n amount
''                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pah4870"
''                OpenQueryDNS cSqlStmt, oTaxSet, False
''
''    '            MsgBox right(oRecordSet("taxcode"), 1)
''                If InStr(oRecordSet("taxcode"), "S") = 1 Then
''                    nTaxAmt = oTaxSet("S_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "H") = 1 Then
''                    nTaxAmt = oTaxSet("H_AMT")
''                End If
''
''                If InStr(oRecordSet("taxcode"), "M") = 1 Then
''                    nTaxAmt = oTaxSet("M_AMT")
''                End If
''
''                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
''                    nTaxAmt = nTaxAmt + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
''                End If
'
'                'new taxcode amount
'                cSqlStmt = "SELECT S_AMT, H_AMT, M_AMT, EX_AMT FROM pa4870"
'                OpenQueryDNS cSqlStmt, oTaxSet, False
'
'                If InStr(oRecordSet("taxcode"), "S") = 1 Then
'                    nTaxAmt2 = oTaxSet("S_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "H") = 1 Then
'                    nTaxAmt2 = oTaxSet("H_AMT")
'                End If
'
'                If InStr(oRecordSet("taxcode"), "M") = 1 Then
'                    nTaxAmt2 = oTaxSet("M_AMT")
'                End If
'
'                If (right(oRecordSet("taxcode"), 1) <> "F") And (right(oRecordSet("taxcode"), 1) <> "E") And (right(oRecordSet("taxcode"), 1) <> "S") Then
'                    nTaxAmt2 = nTaxAmt2 + (right(oRecordSet("taxcode"), 1) * oTaxSet("EX_AMT"))
'                End If
'            Else
'                nTaxAmt = 0
'                nTaxAmt2 = 0
'            End If
'
'
''            aTaxable7_5(11) = Round((nTaxAmt + nTaxAmt2) / 2, 2)
'            aTaxable7_5(11) = Round(nTaxAmt2 / 2, 2)
'
'
'            'aTaxable7_5(12) = round(aTaxable7_5(10) - aTaxable7_5(11), 2)
'            'aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(17)) - aTaxable7_5(11), 2)
'            aTaxable7_5(12) = Round((aTaxable7_5(9) + aTaxable7_5(10)) - aTaxable7_5(11), 2)
'
'
'            'tax due
'            cSqlStmt = " SELECT RANGE1, RANGE2, AMOUNT, PERCENT FROM pa4870 " & _
'                       " Where range1 <= " & aTaxable7_5(12) & " And range2 >=" & aTaxable7_5(12)
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            If objdbRs.RecordCount > 0 Then
'                aTaxable7_5(13) = Round(((aTaxable7_5(12) - objdbRs("range1")) * (objdbRs("percent") / 100)) + objdbRs("amount"), 2)
'            Else
'                aTaxable7_5(13) = 0
'            End If
'
'            aTaxable7_5(14) = Round(aTaxable7_5(13) - oRecordSet("tax_wheld"), 2)
'            aTaxable7_5(15) = Round(oRecordSet("tax_wheld") - aTaxable7_5(13), 2)
'
'            aTaxable7_5(16) = Round(oRecordSet("tax_wheld") + aTaxable7_5(14), 2)
'
'            cSqlStmt = "insert into alpha7_5(sched1,sched2,sched3a,sched3b,sched3c,sched4,sched5a,sched5b,sched5c,sched5d,sched5e,sched5f,sched5g,sched5h,sched5i,sched5j,sched5k,sched5l, " & _
'                       " sched5m,sched5n,sched5o,sched5p,sched5q,sched5r,sched5s,sched5t,sched5u,sched5v,sched5w,sched5x,sched5y,sched5z,sched5aa,sched5ab,sched5ac,sched5ad, " & _
'                       " sched5ae,sched5af,sched5ag,sched6,sched6b,sched7,sched8,sched9,sched10a,sched10b,sched11a,sched11b,sched12)values(" & _
'                       nCtr & "," & _
'                       cQuote & oRecordSet("tin") & cQuote & "," & _
'                       cQuote & oRecordSet("lastname") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & cQuote & oRecordSet("mname") & cQuote & "," & cQuote & "III" & cQuote & "," & _
'                       "0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
'                       cQuote & Format(oRecordSet("date_hire"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("date_fin"), "yyyy-mm-dd") & cQuote & "," & _
'                       aTaxable7_5(0) & "," & _
'                       aTaxable7_5(1) & "," & aTaxable7_5(2) & "," & aTaxable7_5(3) & ",312," & _
'                       aTaxable7_5(4) & "," & aTaxable7_5(5) & "," & aTaxable7_5(6) & ",0," & _
'                       aTaxable7_5(7) & ",0," & aTaxable7_5(8) & "," & aTaxable7_5(18) & "," & aTaxable7_5(9) & "," & aTaxable7_5(17) & "," & aTaxable7_5(10) & "," & aTaxable7_5(10) & "," & _
'                       cQuote & Replace(oRecordSet("taxcode"), "E", "") & cQuote & "," & _
'                       aTaxable7_5(11) & ",0," & _
'                       aTaxable7_5(12) & "," & _
'                       aTaxable7_5(13) & ",0," & _
'                       oRecordSet("tax_wheld") & "," & _
'                       aTaxable7_5(14) & "," & _
'                       aTaxable7_5(15) & "," & _
'                       aTaxable7_5(16) & ")"
''                MsgBox cSqlStmt
'                QueryDBF cSqlStmt, objdbRs, True
'
'            nCtr = nCtr + 1
'            oRecordSet.MoveNext
'        Wend
'    End If
'
'    ShowProgress 4
'
'    MsgBox "done"
'
'    Set oFileSys = Nothing
'    Set oRecordSet = Nothing
'    Set oRSet = Nothing
'    Set oRSet2 = Nothing
'    Set oTaxSet = Nothing
'    Set oTaxSet = Nothing
'
'End Sub

' + -->
' |     Procedure Name  :   GenSSSR1A
' |     Description     :   Generate SSS Employment Report (R-1A)
' |     Date Created    :   15 jan 2008
' + -->
Sub Create_SssR1A()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " create table tmpSSSR1A ( " & _
               " [DATE_HIRE] char(10),          [RATE_AMT] double, " & _
               " [EMPID] char(6),            [FIRSTNAME] char(100), " & _
               " [MNAME] char(100),          [LASTNAME] char(100), " & _
               " [CMPNAME] char(100),        [BIRTHDAY] char(10), " & _
               " [SSNUM] char(50),           [EMP_STAT] integer, " & _
               " [ACTIVE] integer,           [POSNAME] char(100), " & _
               " [DEPNAME] char(100),        [ADDRESS] char(100)," & _
               " [TELNO] char(50), " & _
               " [AREA_NO] char(50),           [EMPLR_ID] char(50), " & _
               " [TAXPAYID] char(50),        [POSTCODE] char(50), " & _
               " [CERT_BY] char(6),          [CERT_NAME] char(100), " & _
               " [CERT_POS] char(100))"
               

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpSSSR1A"
    QueryTemp cSqlStmt, oTempADO, True

End Sub

Sub GenSSSR1A()
    Dim cSqlStmt As String, _
        aUserInfo As Variant, _
        oRecordSet As New ADODB.Recordset
    
    aUserInfo = Array("")
    
    If Not ChkPersonnel(Text6) Then Exit Sub

    Create_SssR1A

    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If
    
     cSqlStmt = " select a.empid, a.ssnum, a.lastname, a.firstname, a.mname, a.birthday,a.rate_amt, " & _
               " a.date_hire,a.emp_stat,a.active,ifnull(b.linename,'') as linename, " & _
               " ifnull(c.posname,'') as posname from di3670 a " & _
               " left join di5463 b on a.depid=b.lineid " & _
               " left join di7670 c on a.posid=c.posid " & _
               " where (year(a.date_hire) = " & Combo1.Text & ")" & _
               " and (month(a.date_hire) = " & ListView1.SelectedItem & ") and (a.emp_stat <> 0) and (a.wap=0) and (a.paystatus <> 2) " & _
               " order by a.date_hire,a.lastname "
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
            cSqlStmt = " insert into tmpSSSR1A(DATE_HIRE,RATE_AMT,EMPID,FIRSTNAME,MNAME,LASTNAME,CMPNAME, BIRTHDAY,SSNUM,EMP_STAT,[ACTIVE]," & _
                       " POSNAME,DEPNAME,ADDRESS,TELNO,AREA_NO,EMPLR_ID,TAXPAYID,POSTCODE,CERT_BY,CERT_NAME,CERT_POS)values(" & _
                        cQuote & Format(oRecordSet("date_hire"), "mm/dd/yyyy") & cQuote & "," & _
                        oRecordSet("rate_amt") * 26 & "," & _
                        cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("firstname") & cQuote & "," & _
                        cQuote & oRecordSet("mname") & cQuote & "," & cQuote & oRecordSet("lastname") & cQuote & "," & _
                        cQuote & cCompany & cQuote & "," & _
                        cQuote & Format(oRecordSet("birthday"), "mm/dd/yyyy") & cQuote & "," & _
                        cQuote & oRecordSet("ssnum") & cQuote & "," & _
                        oRecordSet("emp_stat") & "," & oRecordSet("active") & "," & _
                        cQuote & oRecordSet("posname") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & _
                        cQuote & gAddress & cQuote & "," & cQuote & gTelNum & cQuote & "," & cQuote & gAreaNo & cQuote & "," & _
                        cQuote & gSSSNum & cQuote & "," & cQuote & gTINNum & cQuote & "," & _
                        cQuote & gPostal & cQuote & "," & _
                        cQuote & Text6.Text & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & _
                        cQuote & aUserInfo(0) & cQuote & ")"
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport "SSS R-1A Report", "PRVSSSR1A.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub


' + -->
' |     Procedure Name  :   GenLoadRpt
' |     Description     :   Generate Employee loan report
' |     Date Created    :   05 August 2008
' + -->
Sub Create_LoanRpt()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " create table TmpLoanRpt ( " & _
               " [EMPID] char(6),            [Fullname] char(100), " & _
               " [EMP_STAT] integer,         [ACTIVE] integer, " & _
               " [LineID] char(3),           [LINENAME] char(100), " & _
               " [POSID] char(3),            [POSNAME] char(100), " & _
               " [DEDID] char(3),            [DEDNAME] char(100), " & _
               " [DEDNAME2] char(100),       [SHORT_DESC] char(10), " & _
               " [DEF_AMT] double,           [CUT_OFF_AMT] double, " & _
               " [PERIOD1] integer,          [PERIOD2] integer, " & _
               " [ACC_AMT] double,           [LOAN_AMT] double, " & _
               " [CTRL_NO] char(10),         [DATE_GRANT] date, " & _
               " [DATE_START] date,          [DATE_END] date, " & _
               " [REMARK] char(100),         [DATE_FIN] date, " & _
               " [status] integer,           [DED_AMT] double, " & _
               " [PERIODID] char(5),         [DATE] char(100) )"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TmpLoanRpt"
    QueryTemp cSqlStmt, oTempADO, True

End Sub

Sub GenLoanRpt(ByVal cParam As String)
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        oRset1 As New ADODB.Recordset
    
    Create_LoanRpt
    
    ShowProgress 0
    
    cSqlStmt = " select a.EMPID, ifnull(concat(b.LASTNAME,', ',b.FIRSTNAME,' ', left(b.MNAME,1),'. '),'') as fullname, " & _
               " a.DEF_AMT, a.CUT_OFF_AMT, a.PERIOD1, a.PERIOD2, a.ACC_AMT, a.LOAN_AMT, a.CTRL_NO, a.DATE_GRANT, a.DATE_START, a.DATE_END, a.REMARK, a.DATE_FIN, a.status, " & _
               " ifnull(b.EMP_STAT,0) as emp_stat, " & _
               " ifnull(b.ACTIVE,0) as ACTIVE, " & _
               " ifnull(a.DEDID,'') as DEDID, ifnull(c.DEDNAME,'') as DEDNAME, ifnull(c.DEDNAME2,'') as DEDNAME2, ifnull(c.SHORT_DESC,'') as SHORT_DESC, " & _
               " ifnull(b.depid,'') as depid, ifnull(d.LINENAME,'') as LINENAME, " & _
               " ifnull(b.POSID,'')as POSID, ifnull(e.POSNAME,'') as POSNAME " & _
               " from di3673 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join pa3330 c on a.dedid=c.dedid " & _
               " left join di5463 d on b.depid=d.lineid " & _
               "  left join di7670 e on b.posid=e.posid " & _
               " Where a.Status = 0 " & IIf(cParam <> "", " And a.dedid in " & cParam, "")
'    MsgBox cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
        
            cSqlStmt = " select concat(date_format(b.date_start,'%b %d'),' - ',date_format(b.date_end,'%d, %Y')) as date, " & _
                       " a.ded_amt , a.periodid, 0, 1, b.date_start " & _
                       " from pah87263 a " & _
                       " left join pa7730 b on a.periodid=b.periodid " & _
                       " Where a.ctrl_no = " & cQuote & oRecordSet("ctrl_no") & cQuote & " and a.empid = " & cQuote & oRecordSet("empid") & cQuote & " and a.dedid = " & cQuote & oRecordSet("dedid") & cQuote & _
                       " order by  a.ctrl_no,a.periodid "
                       
'            MsgBox cSqlStmt
            OpenQueryDNS cSqlStmt, oRset1, False
            If oRset1.RecordCount > 0 Then
                While Not oRset1.EOF
                    'd2 na yung insert

                    cSqlStmt = " insert into TmpLoanRpt (EMPID, Fullname, EMP_STAT,[ACTIVE],LineID,LINENAME,POSID,POSNAME,DEDID,DEDNAME,DEDNAME2,SHORT_DESC, " & _
                               " DEF_AMT,CUT_OFF_AMT,PERIOD1,PERIOD2,ACC_AMT,LOAN_AMT,CTRL_NO,DATE_GRANT,DATE_START,DATE_END,REMARK,DATE_FIN,STATUS,DED_AMT,[DATE],PERIODID)values(" & _
                               cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("fullname") & cQuote & "," & _
                               oRecordSet("emp_stat") & "," & oRecordSet("active") & "," & _
                               cQuote & oRecordSet("depid") & cQuote & "," & cQuote & oRecordSet("linename") & cQuote & "," & _
                               cQuote & oRecordSet("posid") & cQuote & "," & cQuote & oRecordSet("posname") & cQuote & "," & _
                               cQuote & oRecordSet("dedid") & cQuote & "," & cQuote & oRecordSet("dedname") & cQuote & "," & _
                               cQuote & oRecordSet("dedname2") & cQuote & "," & cQuote & oRecordSet("short_desc") & cQuote & "," & _
                               oRecordSet("def_amt") & "," & oRecordSet("cut_off_amt") & "," & _
                               oRecordSet("period1") & "," & oRecordSet("period2") & "," & _
                               oRecordSet("acc_amt") & "," & oRecordSet("loan_amt") & "," & _
                               oRecordSet("ctrl_no") & "," & _
                               cQuote & Format(oRecordSet("date_grant"), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & Format(oRecordSet("date_start"), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & Format(oRecordSet("date_end"), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & oRecordSet("remark") & cQuote & "," & _
                               cQuote & Format(oRecordSet("DATE_FIN"), "yyyy-mm-dd") & cQuote & "," & _
                               oRecordSet("status") & "," & oRset1("ded_amt") & "," & _
                               cQuote & oRset1("date") & cQuote & "," & _
                               cQuote & oRset1("periodid") & cQuote & ")"
                               
'                    MsgBox cSqlStmt
                    
                    QueryTemp cSqlStmt, objdbRs, True
                    
                    oRset1.MoveNext
                Wend
            End If
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        

        GenerateReport "Employee Loan Report", "PRV3570.RPT", , True

        ShowProgress 4
        
    Else
        ShowProgress 3
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
        ShowProgress 4
    End If
        
End Sub

Sub GenSLVLPayRoll(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
    Dim cSqlStmt, _
        cPeriodName, _
        cDedParam, _
        cDedValue As String, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        oRSet3 As New ADODB.Recordset, _
        nActive, _
        nCtr As Integer, _
        aOtherInfo As Variant, _
        aPosInfo As Variant, _
        aDedName As Variant, aDedAmt As Variant, _
        aMonthName As Variant

    If SSTab1.TabVisible(3) Then
        If Not ChkPersonnel(Text6) Then Exit Sub
        If Not ChkPersonnel(Text5) Then Exit Sub
        If Not ChkPersonnel(Text1) Then Exit Sub
        If Not ChkPersonnel(Text7) Then Exit Sub
        If Not ChkPersonnel(Text3) Then Exit Sub
        If Not ChkPersonnel(Text8) Then Exit Sub
    End If

    ' --> process active employee first here...
    nActive = IIf(Tag = 15, 1, 0)
    
    aPosInfo = Array("", "", "", "", "", "", "")
    
    aMonthName = Array("Enero", "Pebrero", "Marso", "Abril", "Mayo", "Hunyo", "Hulyo", "Agosto", "Setyembre", "Oktubre", "Nobyembre", "Disyembre")
    
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aPosInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aPosInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text1.Text & "'"
        aPosInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text7.Text & "'"
        aPosInfo(3) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text3.Text & "'"
        aPosInfo(4) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text8.Text & "'"
        aPosInfo(5) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If

    If Trim(cParam) <> "" Then
        cParam = "a.depid IN " & cParam
    End If
    
    If Check3.Value = vbChecked Then
        CreateTmpPaySlip 2  ' --> payroll sheet deduction
    Else
        CreateTmpPaySlip 0  ' --> header
        If (nFilter = 0) Then CreateTmpPaySlip 1 ' --> detail
    End If
    
loopd2:

    aOtherInfo = Array("", "", "", "")

    OpenQueryDNS "select * from pa7730 where periodid=" & cQuote & cPeriodID & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If Tag = 33 Then
            cPeriodName = "SLVL Pay " & Year(objdbRs("date_start"))
        Else
            If Check5.Value = vbChecked Then
                If Combo1.ListIndex <> 2 Then
                    cPeriodName = "Para sa sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & Format(objdbRs("date_start"), " d, yyyy") & " hanggang " & aMonthName(Month(objdbRs("date_end")) - 1) & Format(objdbRs("date_end"), " d, yyyy")
                Else
                    cPeriodName = "Sahod mula " & aMonthName(Month(objdbRs("date_start")) - 1) & " " & Day(objdbRs("date_start")) & "-" & Day(objdbRs("date_end")) & ", " & Year(objdbRs("date_end"))
                End If
            Else
                cPeriodName = "For the " & IIf(Tag = 5, "SLVL Payroll", "") & " period " & Format(objdbRs("date_start"), "mmm d, yyyy") & " to " & Format(objdbRs("date_end"), "mmm d, yyyy")
            End If
        End If
    End If
    
    If (Check4.Value = vbChecked) Then
        cSqlStmt = " select a.BACCNTNO,a.periodid, a.seq_no, " & _
                   " a.empid, a.firstname, a.lastname, a.mname, a.emp_stat,  a.fullname, " & _
                   " a.depid, a.rate_amt, " & _
                   " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as LEAVE_PAY, " & _
                   " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as gross_pay, " & _
                   " ((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt as net_pay, " & _
                   " a.active " & _
                   " from pa87260 a " & _
                   " left join di3670 b on a.empid=b.empid"
    Else
        cSqlStmt = "select a.BACCNTNO,a.periodid, a.p_day, a.p_holiday, a.depid, count(a.empid) as manpower,sum(truncate(a.cola_amt*(a.reg_day+a.ndiff_day),2)) as cola_amt, " & _
                   " sum(((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt) as leave_pay, " & _
                   " sum(((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt) as gross_pay, " & _
                   " sum(((b.sl_avail+b.vl_avail)-(b.sl_use+b.vl_use))* b.rate_amt) as net_pay" & _
                   " from pa87260 a " & _
                   " left join di3670 b on a.empid=b.empid "
    End If
    
   
  
    If gAgency = "0" Then
         If gCompanyID <> "0001" Then
             cSqlStmt = cSqlStmt & " where (a.active = 0)" & IIf(cParam <> "", " and (" & cParam & ")", "") & " and a.emp_stat = 2 " & _
                    " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                    " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                    IIf((Check4.Value <> vbChecked), " group by a.depid", "")
         Else
             cSqlStmt = cSqlStmt & " where (a.active = 0)" & IIf(cParam <> "", " and (" & cParam & ")", "") & " and a.emp_stat <> 0 and b.sl_avail <> 0 " & _
                    " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                    " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                    IIf((Check4.Value <> vbChecked), " group by a.depid", "")
    
          End If
     Else
         cSqlStmt = cSqlStmt & " where (a.active = 0)" & IIf(cParam <> "", " and (" & cParam & ")", "") & " and a.emp_stat = 1 " & _
                    " and (a.periodid=" & cQuote & cPeriodID & cQuote & ")" & _
                    " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                    IIf((Check4.Value <> vbChecked), " group by a.depid", "")
    
    End If
   
   
   
'    MsgBox cSqlStmt
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
'        OpenQueryDNS "select posid, posname from DI7670 order by posid", oRSet1, False
        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRset1, False
        
        ShowProgress 0
        
        While Not oTempADO.EOF
        
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
            
'            If oTempADO("emp_stat") = 2 Then
            
                '---> need to kasama to
                If oRset1.RecordCount > 0 Then
                    oRset1.Requery adAsyncFetch
                    oRset1.Find "lineid='" & oTempADO("depid") & "'"
                    aOtherInfo(1) = IIf(oRset1.EOF, "", oRset1("linename"))
                End If
                
                If Check4.Value <> vbChecked Then
                    cSqlStmt = "insert into tmp7297655(periodname, p_day, p_holiday, depid, deptname," & IIf(nActive = 1, " depid2, deptname2,", "") & _
                               " rate_amt, leave_pay, gross_pay, net_pay,BACCNTNO," & _
                               " signatory1,signatory2,signatory3,signatory4,signatory5,signatory6,signatory7, " & _
                               " posname1,posname2,posname3,posname4,posname5,posname6,posname7)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("p_day") & "," & oTempADO("p_holiday") & "," & _
                               cQuote & IIf(nActive = 1, IIf(Tag = 15, oTempADO("depid"), "999"), oTempADO("depid")) & cQuote & "," & cQuote & IIf(nActive = 1, IIf(Tag = 15, DecodeStr(EncodeStr2(aOtherInfo(1))), "Resigned/FC"), DecodeStr(EncodeStr2(aOtherInfo(1)))) & cQuote & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               oTempADO("manpower") & "," & oTempADO("leave_pay") & "," & oTempADO("gross_pay") & "," & _
                               oTempADO("net_pay") & "," & _
                               cQuote & oTempADO("BACCNTNO") & cQuote & "," & _
                               cQuote & EncodeStr2(Label8.Caption) & cQuote & "," & cQuote & EncodeStr2(Label6.Caption) & cQuote & "," & cQuote & EncodeStr2(Label4.Caption) & cQuote & "," & _
                               cQuote & EncodeStr2(Label15.Caption) & cQuote & "," & cQuote & EncodeStr2(Label14.Caption) & cQuote & "," & cQuote & EncodeStr2(Label16.Caption) & cQuote & "," & cQuote & cQuote & "," & _
                               cQuote & aPosInfo(0) & cQuote & "," & cQuote & aPosInfo(1) & cQuote & "," & cQuote & aPosInfo(2) & cQuote & "," & cQuote & aPosInfo(3) & cQuote & "," & cQuote & aPosInfo(4) & cQuote & "," & cQuote & aPosInfo(5) & cQuote & "," & cQuote & aPosInfo(6) & cQuote & ")"
                Else
                    cSqlStmt = "insert into tmp7297655(periodname, seq_no, depid, deptname," & _
                               IIf(nActive = 1, " depid2, deptname2,", "") & _
                               " empid, emp_stat, [active], fullname, fname, mname, lname, " & _
                               " rate_amt, LEAVE_PAY, gross_pay, net_pay,BACCNTNO)values(" & _
                               cQuote & DecodeStr(EncodeStr2(cPeriodName)) & cQuote & "," & oTempADO("seq_no") & "," & _
                               IIf(nActive = 1, "999", cQuote & oTempADO("depid") & cQuote) & "," & _
                               IIf(nActive = 1, "Resigned/FC", cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote) & "," & _
                               IIf(nActive = 1, cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & ",", "") & _
                               cQuote & oTempADO("empid") & cQuote & "," & oTempADO("emp_stat") & "," & oTempADO("active") & "," & _
                               cQuote & DecodeStr(EncodeStr2(oTempADO("fullname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("firstname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("mname"))) & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("lastname"))) & cQuote & "," & _
                               oTempADO("rate_amt") & "," & oTempADO("LEAVE_PAY") & "," & oTempADO("gross_pay") & "," & _
                               oTempADO("net_pay") & "," & _
                               cQuote & oTempADO("BACCNTNO") & cQuote & ")"
                End If
'                    MsgBox cSqlStmt
                QueryTemp cSqlStmt, objdbRs, True
'            End If
            
            oTempADO.MoveNext
        Wend
        
        ShowProgress 4
    End If
    
    ShowProgress 0
    
    QueryTemp "select * from " & IIf(Check3.Value <> vbChecked, "tmp7297655", "tmp7297655d2"), objdbRs, False
    If objdbRs.RecordCount > 0 Then
    
        ShowProgress 3
        
         Select Case Tag
            Case 32     ' --> slvl
                GenerateReport IIf(Check5.Value <> vbChecked, "PAYSLIP", "TALAAN NG KINITA"), IIf(Check5.Value = vbChecked, "rpt7297547T.rpt", "rpt7297547.rpt")
            Case 33     ' --> SLVL payroll sheets
                GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "") & cPeriodName, IIf(Check4.Value = vbChecked, "rpt13667.rpt", "rpt13667s.rpt")
            Case 34      ' --> Acknowledgement
                GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "") & "SLVL Acknowledgement Report", "rpt734748.rpt"
                
        End Select
        
    Else
        ShowProgress 3

        MsgBox "No report to generate!", vbCritical, "System Advisory"
    End If
    
    ShowProgress 4
    
    Set oRset1 = Nothing
    Set oRSet2 = Nothing
    Set oRSet3 = Nothing
End Sub

Sub CreateTmpManPower()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = "CREATE TABLE tmpManPowerLst(" & _
               "[LINEID] char(3),           [LINENAME] char(100)," & _
               "[EMPID] char(6),            [FULLNAME] char(100)," & _
               "[TCID] char(6),             [date_hire] date," & _
               "[status] char (4),          [emp_stat] integer," & _
               "[paystatus] integer,        [active] integer," & _
               "[date_fin] date,            [position] char(100)," & _
               "[WAP] integer )"
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpManPowerLst", oTempADO, True
End Sub

Sub GenManPowerList(ByVal cPeriod As String, cParam As String, nMode As Integer)
    Dim cSqlStmt As String, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        aOtherInfo As Variant, _
        cString As String
        
    aOtherInfo = Array("", "", "", "")
    
    CreateTmpManPower
    
    ShowProgress 0
    
    
    cString = " and a.active " & IIf(nMode <> 0, "> 0", "=0")
    
    If Trim(cParam) <> "" Then
        cParam = " and a.depid IN " & cParam
    End If

    cSqlStmt = " SELECT a.date_hire,b.date_fin,a.date_res,a.active,a.empid,b.tcid, a.firstname, a.mname, a.lastname, concat(a.lastname,', ',a.firstname, ', ',if(trim(b.mname)='','',concat(left(b.mname,1),'.'))) as fullname, b.posid, a.depid, " & _
               " a.date_hire,b.date_fin,a.date_res,if(a.paystatus=2,'EM',if(a.emp_stat=0,'WAP',if(a.emp_stat=1,if(a.wap=1,'WAP','C'),if(a.emp_stat=2,'R','')))) as status,a.emp_stat, a.paystatus, a.active,a.wap  "
            
 
    OpenQueryDNS "select pclose from pa7730 where periodid =" & cQuote & cPeriod & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If objdbRs("pclose") <> 1 Then
            
            cSqlStmt = cSqlStmt & " FROM pa87260 a left join di3670 b on a.empid=b.empid "
        
        Else
            cSqlStmt = cSqlStmt & " FROM pah87260 a left join di3670 b on a.empid=b.empid "

        End If
    End If
        
    cSqlStmt = cSqlStmt & "where a.periodid=" & cQuote & cPeriod & cQuote & cString & cParam
'    Script2File cSqlStmt
'    MsgBox cSqlStmt
'            cSqlStmt = "select a.date_hire, a.birthday, a.date_fin, a.date_res, a.status, a.rate_amt, a.cola_amt, a.pos_allow, a.isunion, a.emp_stat, a.paystatus, a.active, " & _
'                       " a.empid, a.firstname, a.mname, a.lastname, concat(a.lastname,', ',a.firstname) as fullname, a.sex, a.posid, a.depid, a.taxid, a.shiftid," & _
'                       " a.ssnum, a.pagibigno, a.tin from di3670 a"
               
'    cSqlStmt = cSqlStmt & " where a.active=" & nmode & IIf(Trim(cParam) = "", "", " and " & cParam)
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        
        OpenQueryDNS "select posid, posname from DI7670 order by posid", oRset1, False
        OpenQueryDNS "select lineid, linename from di5463 order by lineid", oRSet2, False

        While Not oTempADO.EOF
            
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100, , , "Processing data for " & oTempADO("fullname")
            
            If oRset1.RecordCount > 0 Then
                oRset1.Requery adAsyncFetch
                oRset1.Find "posid='" & oTempADO("posid") & "'"
                aOtherInfo(0) = IIf(oRset1.EOF, "", oRset1("posname"))
            End If
            
            If oRSet2.RecordCount > 0 Then
                oRSet2.Requery adAsyncFetch
                oRSet2.Find "lineid='" & oTempADO("depid") & "'"
                aOtherInfo(1) = IIf(oRSet2.EOF, "", oRSet2("linename"))
            End If
            
               cSqlStmt = "insert into tmpManPowerLst(empid,tcid,fullname,lineid,linename,date_hire,date_fin,[status]," & _
                       " emp_stat, paystatus,[active],[position],[WAP])values(" & _
                       cQuote & oTempADO("empid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(oTempADO("tcid"))) & cQuote & "," & _
                       cQuote & EncodeStr2(DecodeStr(oTempADO("fullname"))) & cQuote & "," & _
                       cQuote & oTempADO("depid") & cQuote & "," & cQuote & DecodeStr(EncodeStr2(aOtherInfo(1))) & cQuote & "," & _
                       cQuote & Format(oTempADO("date_hire"), "mm/dd/yyyy") & cQuote & "," & cQuote & Format(oTempADO("date_fin"), "mm/dd/yyyy") & cQuote & "," & _
                       cQuote & oTempADO("status") & cQuote & "," & oTempADO("emp_stat") & "," & _
                       oTempADO("paystatus") & "," & oTempADO("active") & "," & _
                       cQuote & DecodeStr(EncodeStr2(aOtherInfo(0))) & cQuote & "," & _
                       oTempADO("wap") & ")"


'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
            
            oTempADO.MoveNext
        Wend
    End If
    
    ShowProgress 3
    
    QueryTemp "select * from tmpManPowerLst", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        GenerateReport " Manpower Report Listing", "lstManpower.rpt"
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    ShowProgress 4
    
EndGenManPower:
    Set oRset1 = Nothing
    Set oRSet2 = Nothing

    Exit Sub
    
ErrGenManPower:
    ErrorMsg Err.Number, Err.Description, "Manpower Report Listing", Name
    
    Resume EndGenManPower
End Sub

'2009-08-12 ATM Excell

Sub createBackupATMxExell(nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cFieldName As String, _
        nCtr As Integer
        
    Select Case nMode
        Case 0
            cSqlStmt = "CREATE TABLE NETCACCN (" & _
                       " [ACCNTNO] char(16), " & _
                       " [AMOUNT] double, " & _
                       " [PAYNAME] char(99))"
        Case 1
            cSqlStmt = "CREATE TABLE NETCARD (" & _
                       " [CARDNO] char(20), " & _
                       " [AMOUNT] double, " & _
                       " [PAYNAME] char(99))"
        Case 2
            cSqlStmt = "CREATE TABLE NETTRAN (" & _
                       " [ACCNTNO] char(16), " & _
                       " [AMOUNT] double, " & _
                       " [CARDNO] char(20))"
        Case 3
            cSqlStmt = "CREATE TABLE SACACCN (" & _
                       " [ACCNTNO] char(16), " & _
                       " [AMOUNT] double, " & _
                       " [PAYNAME] char(99))"
        Case 4
            cSqlStmt = "CREATE TABLE SACARD (" & _
                       " [CARDNO] char(20), " & _
                       " [AMOUNT] double, " & _
                       " [PAYNAME] char(99))"
        Case 5
            cSqlStmt = "CREATE TABLE SATRAN (" & _
                       " [ACCNTNO] char(16), " & _
                       " [AMOUNT] double, " & _
                       " [CARDNO] char(20))"
    End Select
    
    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    Select Case nMode
        Case 0
            cSqlStmt = "DELETE FROM NETCACCN"
        Case 1
            cSqlStmt = "DELETE FROM NETCARD"
        Case 2
            cSqlStmt = "DELETE FROM NETTRAN"
        Case 3
            cSqlStmt = "DELETE FROM SACACCN"
        Case 4
            cSqlStmt = "DELETE FROM SACARD"
        Case 5
            cSqlStmt = "DELETE FROM SATRAN"
    End Select
    
    QueryDBF cSqlStmt, oTempADO, True
End Sub

Sub GenManATMExcell(cPeriod As String, cParam As String)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cWant As String, _
        cPayPath As String, _
        cDuration As String, _
        cDurationPath As String, _
        cHistoryPath As String, _
        FileSys As FileSystemObject, _
        nCtr As Integer

    Set FileSys = New FileSystemObject

    cSqlStmt = " SELECT duration FROM pa7730 " & _
               " Where periodid = " & cQuote & cPeriod & cQuote

    OpenQueryDNS cSqlStmt, oRecordSet, False
    cDuration = IIf(oRecordSet.RecordCount > 0, oRecordSet("duration"), "")

    cDurationPath = CheckPath(Text4.Text) & "Backup"
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath

    cDurationPath = cDurationPath & "\" & cDuration
    cDurationPath = CheckPath(cDurationPath)
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath

    If Check4.Value = vbChecked Then
        If (FileSys.FileExists(cDurationPath & "NETCACCN.dbf") = True) Or _
           (FileSys.FileExists(cDurationPath & "NETCARD.dbf") = True) Or _
           (FileSys.FileExists(cDurationPath & "NETTRAN.dbf") = True) Then
            cWant = MsgBox("Files are Already Exsisting... do you want to backup previously created file ?", vbYesNoCancel + vbCritical, App.Title)
            If cWant = vbYes Then
                checkK1ATMPath cDurationPath, cDuration
            ElseIf cWant = vbCancel Then
                GoTo ulit
            End If
    
            FileSys.DeleteFile cDurationPath & "NETCACCN.dbf"
            FileSys.DeleteFile cDurationPath & "NETCARD.dbf"
            FileSys.DeleteFile cDurationPath & "NETTRAN.dbf"
        End If
    Else
        If (FileSys.FileExists(cDurationPath & "SACACCN.dbf") = True) Or _
           (FileSys.FileExists(cDurationPath & "SACARD.dbf") = True) Or _
           (FileSys.FileExists(cDurationPath & "SATRAN.dbf") = True) Then
            cWant = MsgBox("Files are Already Exsisting... do you want to backup previously created file ?", vbYesNoCancel + vbCritical, App.Title)
            If cWant = vbYes Then
                checkK1ATMPath cDurationPath, cDuration
            ElseIf cWant = vbCancel Then
                GoTo ulit
            End If
        
            FileSys.DeleteFile cDurationPath & "SACACCN.dbf"
            FileSys.DeleteFile cDurationPath & "SACARD.dbf"
            FileSys.DeleteFile cDurationPath & "SATRAN.dbf"
        End If
    End If
    
    GenbackupATMExcell cDurationPath, cPeriod, cParam
ulit:
End Sub

Sub GenbackupATMExcell(cPath As String, ByVal cPeriod As String, cParam As String)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cString, _
        cFieldName As String, _
        cBAccntNo As String, _
        nCtr As Integer

    DetectDBF cPath
    If Check4.Value = vbChecked Then
        createBackupATMxExell 0
        createBackupATMxExell 1
        createBackupATMxExell 2
    Else
        createBackupATMxExell 3
        createBackupATMxExell 4
        createBackupATMxExell 5
    End If
    
    'card
    ShowProgress 0
    OpenQueryDNS "select * from di2660 where cmpid=" & cQuote & gCompanyID & cQuote, objdbRs, False
    cString = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    
    cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME,B.MNAME, round(a.GROSS_PAY,2) as GROSS_PAY, round(a.NET_PAY,2) as NET_PAY, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " and b.paystatus = 0 and b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & Combo1.ListIndex) & _
               " order by b.emp_stat desc, a.lastname, a.firstname "
'    Script2File cSqlStmt
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF
                
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            
            If (IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY"))) > 0 Then
                cSqlStmt = " INSERT INTO " & IIf(Check4.Value = vbChecked, "NETCARD", "SACARD") & _
                           " (CARDNO,[AMOUNT],PAYNAME)VALUES(" & _
                           cQuote & Replace(oRecordSet("BACCNTNO"), "-", "", 1, Len(oRecordSet("BACCNTNO")), vbTextCompare) & cQuote & "," & _
                           IIf(Check4.Value = vbChecked, Round(oRecordSet("NET_PAY"), 2), Round(oRecordSet("SA_NET_PAY"), 2)) & "," & _
                           cQuote & Replace(oRecordSet("Lastname"), "ñ", "", 1, Len(oRecordSet("Lastname")), vbTextCompare) & ", " & _
                           Replace(oRecordSet("firstname"), "ñ", "", 1, Len(oRecordSet("firstname")), vbTextCompare) & " " & _
                           Replace(oRecordSet("mname"), "ñ", "", 1, Len(oRecordSet("mname")), vbTextCompare) & "." & cQuote & ")"
    '            MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True
            End If
            oRecordSet.MoveNext
        Wend

        ShowProgress 4
    Else
        ShowProgress 4
        MsgBox "Data not found...!!!", vbInformation, App.Title
        Exit Sub
    End If

    'Transmital
    ShowProgress 0
    cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME, round(a.GROSS_PAY,2) as GROSS_PAY, round(a.NET_PAY,2) as NET_PAY, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 and b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & Combo1.ListIndex) & _
               " order by b.emp_stat desc, a.lastname, a.firstname "
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF
                
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            
            If (IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY"))) > 0 Then
                cSqlStmt = " INSERT INTO " & IIf(Check4.Value = vbChecked, "NETTRAN", "SATRAN") & _
                           " (ACCNTNO,[AMOUNT],CARDNO)VALUES(" & _
                           cQuote & Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)) & cQuote & "," & _
                           IIf(Check4.Value = vbChecked, Round(oRecordSet("NET_PAY"), 2), Round(oRecordSet("SA_NET_PAY"), 2)) & "," & _
                           cQuote & Replace(oRecordSet("BACCNTNO"), "-", "", 1, Len(oRecordSet("BACCNTNO")), vbTextCompare) & cQuote & ")"
                           
    '            MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True
            End If
            oRecordSet.MoveNext

        Wend

        ShowProgress 4
    Else
        ShowProgress 4
        MsgBox "Data not found...!!!", vbInformation, App.Title
        Exit Sub
    End If

    ShowProgress 0
    
    'Account
    cSqlStmt = " select a.EMPID,round(sum(a.NET_PAY),2) as NET_PAY, round(sum(a.SA_NET_PAY),2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 And b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & Combo1.ListIndex) & _
               " group by periodid "
               
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "..."
            
            If (IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY"))) > 0 Then
                cSqlStmt = " INSERT INTO " & IIf(Check4.Value = vbChecked, "NETCACCN", "SACACCN") & _
                           " (ACCNTNO,[AMOUNT],PAYNAME)VALUES(" & _
                           cQuote & Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)) & cQuote & "," & _
                           IIf(Check4.Value = vbChecked, Round(oRecordSet("NET_PAY"), 2), Round(oRecordSet("SA_NET_PAY"), 2)) & "," & _
                           cQuote & cString & cQuote & ")"
    '            MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True
            End If
            oRecordSet.MoveNext

        Wend

        ShowProgress 4

        MsgBox "File Process Done... Press [OK] to continue...", vbInformation, App.Title
    Else
        MsgBox "Data not found...!!!", vbInformation, App.Title
    End If
    
    Set oRecordSet = Nothing
    
End Sub

Sub checkK1ATMPath(cPathcheck As String, cDuration As String)
    Dim cWant As String, _
        cHistoryPath As String, _
        FileSys As FileSystemObject, _
        nCtr As Integer
        
    Set FileSys = New FileSystemObject
    
    cHistoryPath = ""
                    
    cHistoryPath = CheckPath(Text4.Text) & "history"
    
    cHistoryPath = CheckPath(cHistoryPath)
    
    If Dir(cHistoryPath, vbDirectory) = "" Then MkDir cHistoryPath

    cHistoryPath = cHistoryPath & cDuration & "\"
    
    If Dir(cHistoryPath, vbDirectory) = "" Then MkDir cHistoryPath
    
    cHistoryPath = CheckPath(cHistoryPath)

Loop1:
    If Dir(cHistoryPath) = "" Then
        If Check4.Value = vbChecked Then
            FileSys.CopyFile cPathcheck & "NETCACCN.dbf", cHistoryPath & "NETCACCN.dbf"
            FileSys.CopyFile cPathcheck & "NETCARD.dbf", cHistoryPath & "NETCARD.dbf"
            FileSys.CopyFile cPathcheck & "NETTRAN.dbf", cHistoryPath & "NETTRAN.dbf"
        Else
            FileSys.CopyFile cPathcheck & "SACACCN.dbf", cHistoryPath & "SACACCN.dbf"
            FileSys.CopyFile cPathcheck & "SACARD.dbf", cHistoryPath & "SACARD.dbf"
            FileSys.CopyFile cPathcheck & "SATRAN.dbf", cHistoryPath & "SATRAN.dbf"
        End If
    Else
        If Check4.Value = vbChecked Then
            If (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "NETCACCN" & ".dbf") = True) Or _
               (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "NETCARD" & ".dbf") = True) Or _
               (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "NETTRAN" & ".dbf") = True) Then
                nCtr = nCtr + 1
                GoTo Loop1
            Else
                FileSys.CopyFile cPathcheck & "NETCACCN.dbf", cHistoryPath & nCtr & "NETCACCN.dbf"
                FileSys.CopyFile cPathcheck & "NETCARD.dbf", cHistoryPath & nCtr & "NETCARD.dbf"
                FileSys.CopyFile cPathcheck & "NETTRAN.dbf", cHistoryPath & nCtr & "NETTRAN.dbf"
            End If
        Else
            If (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "SACACCN" & ".dbf") = True) Or _
               (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "SACARD" & ".dbf") = True) Or _
               (FileSys.FileExists(cHistoryPath & IIf(nCtr <> 0, nCtr, "") & "SATRAN" & ".dbf") = True) Then
                nCtr = nCtr + 1
                GoTo Loop1
            Else
                FileSys.CopyFile cPathcheck & "SACACCN.dbf", cHistoryPath & nCtr & "SACACCN.dbf"
                FileSys.CopyFile cPathcheck & "SACARD.dbf", cHistoryPath & nCtr & "SACARD.dbf"
                FileSys.CopyFile cPathcheck & "SATRAN.dbf", cHistoryPath & nCtr & "SATRAN.dbf"
            End If
        End If
    End If
End Sub


Sub GenManATMIX1(cPeriod As String, cParam As String)
    Dim oTextFile As New FileSystemObject, _
            oTxtStream As TextStream, _
            FileSys As FileSystemObject, _
            oFile As File, _
            oRecordSet As New ADODB.Recordset, _
            cWant As String, _
            cDuration As String, _
            cDurationPath As String, _
            cIX1 As String, _
            cCmpname As String, _
            cSqlStmt As String, _
            cString As String, _
            cRecCnt As String
        
    OpenQueryDNS "select * from di2660 where cmpid=" & cQuote & gCompanyID & cQuote, objdbRs, False
    cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")


    Set FileSys = New FileSystemObject

    cSqlStmt = " SELECT duration FROM pa7730 " & _
               " Where periodid = " & cQuote & cPeriod & cQuote

    OpenQueryDNS cSqlStmt, objdbRs, False
    cDuration = IIf(objdbRs.RecordCount > 0, objdbRs("duration"), "")

    cDurationPath = CheckPath(Text4.Text) & "Backup"
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath

    cDurationPath = cDurationPath & "\" & cDuration
    cDurationPath = CheckPath(cDurationPath)
'    MsgBox Chr(241)
    
    If Dir(cDurationPath, vbDirectory) = "" Then MkDir cDurationPath

    cIX1 = CheckPath(cDurationPath) & "S" & gRCBCNo & Day(Now) & ".IX1"

    If FileSys.FileExists(cIX1) = True Then
        cWant = MsgBox(cIX1 & " Already Exsist... do you want to overwrite the file ?", vbYesNo + vbCritical, App.Title)
        If cWant = vbYes Then
            FileSys.DeleteFile cIX1
        Else
            Exit Sub
        End If
    End If

    If Dir(cIX1) = "" Then
        Set oTxtStream = oTextFile.CreateTextFile(cIX1, True)
    Else
        Set oFile = oTextFile.GetFile(cIX1)
        Set oTxtStream = oFile.OpenAsTextStream(ForAppending)
    End If
    
    cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME,left(B.MNAME,1) as MNAME, round(a.GROSS_PAY,2) as GROSS_PAY, round(a.NET_PAY,2) as NET_PAY, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 and b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & cQuote & Combo1.ListIndex & cQuote) & _
               " order by a.lastname desc, a.firstname desc "
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, objdbRs, False
    cRecCnt = PadStr(Str(objdbRs.RecordCount + 1), "0", 5)

    cString = "H" & Format(Now, "mmddyy") & cRecCnt
    oTxtStream.WriteLine cString
    
    'Account
    ShowProgress 0
    
    cSqlStmt = " select a.EMPID,round(sum(a.NET_PAY),2) as NET_PAY, round(sum(a.SA_NET_PAY),2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 And b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & Combo1.ListIndex) & _
               " group by periodid "
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF

            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "..."
            
            cString = ""
            cString = Replace(Round(IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY")), 2), ".", "", 1, Len(Round(IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY")), 2)), vbTextCompare)
            cString = PadStr(cString, "0", 15)
            cString = "D" & PadStr(Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)), "0", 14) & "  " & _
                      cString & _
                      UCase(cCmpname)
                      
            oTxtStream.WriteLine cString
            
            oRecordSet.MoveNext

        Wend

        ShowProgress 4

    Else
    
        ShowProgress 4
        
        MsgBox "Data not found...!!!", vbInformation, App.Title
        Exit Sub
    End If
    
    'card
    ShowProgress 0
    
    OpenQueryDNS "select * from di2660 where cmpid=" & cQuote & gCompanyID & cQuote, objdbRs, False
    cString = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    
    cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME,left(B.MNAME,1) as MNAME, round(a.GROSS_PAY,2) as GROSS_PAY, round(a.NET_PAY,2) as NET_PAY, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 and b.active = 0 " & _
               IIf(Combo1.ListIndex > 2, "", " and b.emp_stat=" & cQuote & Combo1.ListIndex & cQuote) & _
               " order by a.lastname desc, a.firstname desc "
'    MsgBox cSqlStmt
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF
                
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
            If (IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY"))) > 0 Then
                cString = ""
                cString = Replace(Round(IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY")), 2), ".", "", 1, Len(Round(IIf(Check4.Value = vbChecked, oRecordSet("NET_PAY"), oRecordSet("SA_NET_PAY")), 2)), vbTextCompare)
                cString = PadStr(cString, "0", 15)
                cString = "C" & PadStr(Replace(IIf(oRecordSet("BACCNTNO") <> "", oRecordSet("BACCNTNO"), "000000000000000"), "-", "", 1, Len(IIf(oRecordSet("BACCNTNO") <> 0, oRecordSet("BACCNTNO"), "000000000000000")), vbTextCompare), "0", 14) & _
                          cString & _
                          Replace(oRecordSet("Lastname"), "ñ", "", 1, Len(oRecordSet("Lastname")), vbTextCompare) & ", " & _
                          Replace(oRecordSet("firstname"), "ñ", "", 1, Len(oRecordSet("firstname")), vbTextCompare) & " " & _
                          Replace(oRecordSet("mname"), "ñ", "", 1, Len(oRecordSet("mname")), vbTextCompare) & "."
                 
                oTxtStream.WriteLine cString
            End If
            oRecordSet.MoveNext

        Wend

        ShowProgress 4
    Else
    
        ShowProgress 4
        
        MsgBox "Data not found...!!!", vbInformation, App.Title
        Exit Sub
    End If
    
    cString = ""
    cString = "T88960227135"
    oTxtStream.WriteLine cString
    
    oTxtStream.Close
    
    MsgBox "File Process Done... Press [OK] to continue...", vbInformation, App.Title

    Set oTxtStream = Nothing
    Set oTextFile = Nothing
    Set oFile = Nothing
    
End Sub

Sub Create_ATMTrans()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " create table TmpATMTrans ( " & _
               " [CMPNAME] char(100),        [BACCNTNO] char(16), " & _
               " [GATM_TOT] double,         [CARDNO] char(16), " & _
               " [CATM_TOT] double,         [PAYNAME] char(100), " & _
               " [HASH] char(15))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM TmpATMTrans"
    QueryTemp cSqlStmt, oTempADO, True

End Sub


Sub GenManATMTrans(ByVal cPeriod As String, cParam As String)
    Dim cSqlStmt As String, _
        oRecordSet As New ADODB.Recordset, _
        nNet_pay As Double, _
        nSA_Net_pay As Double, _
        nSA_Reg As Double, _
        cCmpname As String, _
        oRset1 As New ADODB.Recordset
    
    Create_ATMTrans
    
    ShowProgress 0
    
    OpenQueryDNS "select * from di2660 where cmpid=" & cQuote & gCompanyID & cQuote, objdbRs, False
    cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    
'para sa SA
    If Combo1.ListIndex = 0 Then
        cSqlStmt = " select a.EMPID,round(sum(a.SA_NET_PAY),2) as SA_NET_PAY " & _
                   " from pa87260 a " & _
                   " left join di3670 b on a.empid=b.empid " & _
                   " left join di5463 c on b.depid=c.lineid " & _
                   " Where a.periodid = " & cQuote & cPeriod & cQuote & _
                   IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
                   " And b.paystatus = 0 And b.active = 0 " & _
                   " and b.emp_stat <>0 " & _
                   " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                   " group by periodid "
        OpenQueryDNS cSqlStmt, objdbRs, False
                
        nSA_Reg = IIf(objdbRs.RecordCount > 0, objdbRs("SA_Net_pay"), 0)
    Else
        nSA_Reg = 0
    End If
    
    cSqlStmt = " select a.EMPID,round(sum(a.NET_PAY),2) as NET_PAY, round(sum(a.SA_NET_PAY),2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 And b.active = 0 " & _
               IIf(Combo1.ListIndex = 0, " and b.emp_stat =0 ", " and b.emp_stat <>0 ") & _
               " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
               " group by periodid "
               
'                    .AddItem "WAP"
'                    .AddItem "Regular"
               
   
    OpenQueryDNS cSqlStmt, objdbRs, False
    nNet_pay = IIf(objdbRs.RecordCount > 0, objdbRs("Net_pay"), 0)
    nSA_Net_pay = IIf(objdbRs.RecordCount > 0, objdbRs("SA_Net_pay"), 0)
    
    OpenQueryDNS "select * from di2660 where cmpid=" & cQuote & gCompanyID & cQuote, objdbRs, False
    cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    
    If Combo1.ListIndex = 0 Then
        cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME,B.MNAME, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
                   " from pa87260 a " & _
                   " left join di3670 b on a.empid=b.empid " & _
                   " left join di5463 c on b.depid=c.lineid " & _
                   " Where a.periodid = " & cQuote & cPeriod & cQuote & _
                   IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
                   " And b.paystatus = 0 and b.active = 0 " & _
                   " and b.emp_stat <>0 " & _
                   " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
                   " order by a.lastname, a.firstname "
    '    Script2File cSqlStmt
        
        OpenQueryDNS cSqlStmt, oRecordSet, False
        If oRecordSet.RecordCount > 0 Then
    
            While Not oRecordSet.EOF
                    
                ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
                If oRecordSet("SA_NET_PAY") <> 0 Then
                    cSqlStmt = " INSERT INTO TmpATMTrans (CMPNAME,BACCNTNO,GATM_TOT,CARDNO,CATM_TOT,PAYNAME,HASH)VALUES(" & _
                               cQuote & cCmpname & cQuote & "," & _
                               cQuote & Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)) & cQuote & "," & _
                               Round(nSA_Reg, 2) + Round(nNet_pay, 2) + Round(nSA_Net_pay, 2) & "," & _
                               cQuote & Replace(oRecordSet("BACCNTNO"), "-", "", 1, Len(oRecordSet("BACCNTNO")), vbTextCompare) & cQuote & "," & _
                               Round(oRecordSet("SA_NET_PAY"), 2) & "," & _
                               cQuote & Replace(oRecordSet("Lastname"), "ñ", "", 1, Len(oRecordSet("Lastname")), vbTextCompare) & ", " & _
                               Replace(oRecordSet("firstname"), "ñ", "", 1, Len(oRecordSet("firstname")), vbTextCompare) & " " & _
                               Replace(oRecordSet("mname"), "ñ", "", 1, Len(oRecordSet("mname")), vbTextCompare) & "." & cQuote & "," & _
                               cQuote & "88960227135" & cQuote & ")"
        '            MsgBox cSqlStmt
                    QueryTemp cSqlStmt, objdbRs, True
                End If
                oRecordSet.MoveNext
            Wend
        End If
    End If
    
    cSqlStmt = " select a.EMPID,b.BACCNTNO,a.FULLNAME,B.LASTNAME,B.FIRSTNAME,B.MNAME, round(a.GROSS_PAY,2) as GROSS_PAY, round(a.NET_PAY,2) as NET_PAY, round(a.SA_NET_PAY,2) as SA_NET_PAY " & _
               " from pa87260 a " & _
               " left join di3670 b on a.empid=b.empid " & _
               " left join di5463 c on b.depid=c.lineid " & _
               " Where a.periodid = " & cQuote & cPeriod & cQuote & _
               IIf(cParam <> "", " and c.lineid in " & cParam, "") & _
               " And b.paystatus = 0 and b.active = 0 " & _
               IIf(Combo1.ListIndex = 0, " and b.emp_stat =0 ", " and b.emp_stat <>0 ") & _
               " and (a.BACCNTNO " & IIf(Check6.Value <> vbChecked, "<>", "=") & cQuote & "" & cQuote & ")" & _
               " order by a.lastname, a.firstname "
'    Script2File cSqlStmt
    
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF
                
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
'            cSqlStmt = " INSERT INTO TmpATMTrans (CMPNAME,BACCNTNO,GATM_TOT,CARDNO,CATM_TOT,PAYNAME,HASH)VALUES(" & _
'                       cQuote & cCmpname & cQuote & "," & _
'                       cQuote & Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)) & cQuote & "," & _
'                       IIf(Check4.Value = vbChecked, Round(nNet_pay, 2), Round(nSA_Net_pay, 2)) & "," & _
'                       cQuote & Replace(oRecordSet("BACCNTNO"), "-", "", 1, Len(oRecordSet("BACCNTNO")), vbTextCompare) & cQuote & "," & _
'                       IIf(Check4.Value = vbChecked, Round(oRecordSet("NET_PAY"), 2), Round(oRecordSet("SA_NET_PAY"), 2)) & "," & _
'                       cQuote & Replace(oRecordSet("Lastname"), "ñ", "", 1, Len(oRecordSet("Lastname")), vbTextCompare) & ", " & _
'                       Replace(oRecordSet("firstname"), "ñ", ""k, 1, Len(oRecordSet("firstname")), vbTextCompare) & " " & _
'                       Replace(oRecordSet("mname"), "ñ", "", 1, Len(oRecordSet("mname")), vbTextCompare) & "." & cQuote & "," & _
'                       cQuote & "88960227135" & cQuote & ")"
            
            cSqlStmt = " INSERT INTO TmpATMTrans (CMPNAME,BACCNTNO,GATM_TOT,CARDNO,CATM_TOT,PAYNAME,HASH)VALUES(" & _
                       cQuote & cCmpname & cQuote & "," & _
                       cQuote & Trim(Replace(gBAccntNo, "-", "", 1, Len(gBAccntNo), vbTextCompare)) & cQuote & "," & _
                       IIf(Combo1.ListIndex = 0, Round(nSA_Reg, 2) + Round(nNet_pay, 2) + Round(nSA_Net_pay, 2), IIf(Check4.Value = vbChecked, Round(nNet_pay, 2), Round(nSA_Net_pay, 2))) & "," & _
                       cQuote & Replace(oRecordSet("BACCNTNO"), "-", "", 1, Len(oRecordSet("BACCNTNO")), vbTextCompare) & cQuote & "," & _
                       IIf(Combo1.ListIndex = 0, Round(oRecordSet("NET_PAY"), 2) + Round(oRecordSet("SA_NET_PAY"), 2), IIf(Check4.Value = vbChecked, Round(oRecordSet("NET_PAY"), 2), Round(oRecordSet("SA_NET_PAY"), 2))) & "," & _
                       cQuote & Replace(oRecordSet("Lastname"), "ñ", "", 1, Len(oRecordSet("Lastname")), vbTextCompare) & ", " & _
                       Replace(oRecordSet("firstname"), "ñ", "", 1, Len(oRecordSet("firstname")), vbTextCompare) & " " & _
                       Replace(oRecordSet("mname"), "ñ", "", 1, Len(oRecordSet("mname")), vbTextCompare) & "." & cQuote & "," & _
                       cQuote & "88960227135" & cQuote & ")"
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True

            oRecordSet.MoveNext
        Wend

        ShowProgress 3
        GenerateReport IIf(Check6.Value = vbChecked, "NO ATM ", "") & "TRANSMITTAL SHEET", "RCBCTrans.rpt", , True
        ShowProgress 4
        
    Else
        ShowProgress 3
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
        ShowProgress 4
    End If
        
End Sub

Sub CreateGrandOT()
        On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmpGrandOT(" & _
               " [DATE] date,             [GOT_REG] double, " & _
               " [GOT_CON] double,       [GOT_WAP] double ) "
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpGrandOT"
    QueryTemp cSqlStmt, oTempADO, True
End Sub

Sub GenGrandOT(ByVal cPeriod As String)
    Dim oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        oRset1 As New ADODB.Recordset, _
        oRSet2 As New ADODB.Recordset, _
        cString As String, _
        cSqlStmt As String, _
        nGot_Reg As Double, _
        nGot_Con As Double, _
        nGot_Wap As Double, _
        nCtr As Integer, _
        lperiod As Boolean
 
    CreateGrandOT

    ShowProgress 0
    
    cSqlStmt = "SELECT PERIODID, DURATION,pclose FROM PA7730 where periodid = " & cQuote & cPeriod & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        lperiod = IIf(objdbRs("pclose") = 1, True, False)
    End If

'    cSqlStmt = " select date from " & IIf(lperiod = True, "dih36770", "di36770") & " where periodid = " & cQuote & cPeriod & cQuote & _
'               " group by date order by date "

    cSqlStmt = " select date from  " & IIf(lperiod = True, "dih36770", "di36770") & "  where periodid = " & cQuote & cPeriod & cQuote & _
               " group by date order by date "
    
'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then

        While Not oRecordSet.EOF
        
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
                
            'reg
            cSqlStmt = " select a.date,sum(a.reg_ot_hr) as reg_ot_hr,sum(a.sa_reg_ot) as sa_reg_ot,sum(a.nd_ot_hr) as nd_ot_hr,sum(a.sa_nd_ot) as sa_nd_ot,sum(a.sun_ot_hr) as sun_ot_hr,sum(a.sun_nd_ot) as sun_nd_ot,(sum(a.reg_ot_hr)+sum(a.sa_reg_ot)+sum(a.nd_ot_hr)+sum(a.sa_nd_ot)+sum(a.sun_ot_hr)+sum(a.sun_nd_ot)) as GOT_REG " & _
                       " from  " & IIf(lperiod = True, "dih36770", "di36770") & "  a left join di3670 b on a.empid=b.empid " & _
                       " Where a.periodid = " & cQuote & cPeriod & cQuote & _
                       " And b.emp_stat = 2 And b.paystatus <> 1 And Date = " & cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & " group by a.periodid, a.date "
            OpenQueryDNS cSqlStmt, oRSet, False
            nGot_Reg = IIf(oRSet.RecordCount > 0, oRSet("GOT_REG"), 0)
            
            
            'cons
            cSqlStmt = " select a.date,sum(a.reg_ot_hr) as reg_ot_hr,sum(a.sa_reg_ot) as sa_reg_ot,sum(a.nd_ot_hr) as nd_ot_hr,sum(a.sa_nd_ot) as sa_nd_ot,sum(a.sun_ot_hr) as sun_ot_hr,sum(a.sun_nd_ot) as sun_nd_ot,(sum(a.reg_ot_hr)+sum(a.sa_reg_ot)+sum(a.nd_ot_hr)+sum(a.sa_nd_ot)+sum(a.sun_ot_hr)+sum(a.sun_nd_ot)) as GOT_CON " & _
                       " from  " & IIf(lperiod = True, "dih36770", "di36770") & "  a left join di3670 b on a.empid=b.empid " & _
                       " Where a.periodid = " & cQuote & cPeriod & cQuote & _
                       " And b.emp_stat = 1 And b.paystatus <> 1 And Date = " & cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & " group by a.periodid, a.date "
            OpenQueryDNS cSqlStmt, oRset1, False
            nGot_Con = IIf(oRset1.RecordCount > 0, oRset1("GOT_CON"), 0)
            
'            'wap
            cSqlStmt = " select a.date,sum(a.reg_ot_hr) as reg_ot_hr,sum(a.sa_reg_ot) as sa_reg_ot,sum(a.nd_ot_hr) as nd_ot_hr,sum(a.sa_nd_ot) as sa_nd_ot,sum(a.sun_ot_hr) as sun_ot_hr,sum(a.sun_nd_ot) as sun_nd_ot,(sum(a.reg_ot_hr)+sum(a.sa_reg_ot)+sum(a.nd_ot_hr)+sum(a.sa_nd_ot)+sum(a.sun_ot_hr)+sum(a.sun_nd_ot)) as GOT_WAP " & _
                       " from  " & IIf(lperiod = True, "dih36770", "di36770") & "  a left join di3670 b on a.empid=b.empid " & _
                       " Where a.periodid = " & cQuote & cPeriod & cQuote & _
                       " And b.emp_stat = 0 And b.paystatus <> 1 And Date = " & cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & " group by a.periodid, a.date "
            OpenQueryDNS cSqlStmt, oRSet2, False
            nGot_Wap = IIf(oRSet2.RecordCount > 0, oRSet2("GOT_WAP"), 0)

            cSqlStmt = " insert into tmpGrandOT (`DATE`,GOT_REG,GOT_CON,GOT_WAP) values (" & _
                       cQuote & Format(oRecordSet("date"), "yyyy-mm-dd") & cQuote & "," & _
                       Round(nGot_Reg, 2) & "," & Round(nGot_Con, 2) & "," & Round(nGot_Wap, 2) & ")"
'            MsgBox cSqlStmt
            QueryTemp cSqlStmt, objdbRs, True
         
            oRecordSet.MoveNext
        Wend
        
        ShowProgress 3
        GenerateReport "OT GRAND TOTAL REPORT ", "rpt4680.rpt"
'
        ShowProgress 4

    Else
        ShowProgress 3
        MsgBox "No report(s) to generate!", vbExclamation, App.Title
        ShowProgress 4
    End If
    
    Set oRecordSet = Nothing
    Set oRSet = Nothing
    Set oRset1 = Nothing
    Set oRSet2 = Nothing
    
End Sub

' + -->
' |     Procedure Name  :   GenPHILhealth
' |     Description     :   Generate Philhealth Report (Er2)
' |     Date Created    :   11 Sept 2009
' + -->
Sub Create_PhilH()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " create table tmpPhilH ( " & _
               " [PHEALTHNUM] char(50),      [EMPID] char(6), " & _
               " [FULLNAME] char(100),       [POSNAME] char(100), " & _
               " [RATE_AMT] double,          [DATE_HIRE] char(50), " & _
               " [EMP_STAT] integer,         [CMPNAME] char(100), " & _
               " [ADDRESS] char(100),        [POSTCODE] char(50), " & _
               " [EMPLR_ID] char(100),       [CERT_BY] char(6),  " & _
               " [CERT_NAME] char(100),      [CERT_POS] char(100))"
               
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpPhilH"
    QueryTemp cSqlStmt, oTempADO, True

End Sub

Sub GenPHILh()
    Dim cSqlStmt As String, _
        aUserInfo As Variant, _
        cCmpname As String, _
        cCmpAdd As String, _
        oRecordSet As New ADODB.Recordset
        
    Dim cPhilno As String
    
    aUserInfo = Array("")
    
    If Not ChkPersonnel(Text6) Then Exit Sub

    Create_PhilH

    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If
    
    cSqlStmt = " select ifnull(a.phealthnum, '') as phealthnum,a.empid, CONCAT(a.LASTNAME,', ',a.FIRSTNAME,if(trim(a.mname)='','',concat(' ',left(mname,1),'. '))) as FULLNAME, " & _
               " ifnull(c.posname,'') as posname, a.rate_amt, a.date_hire, a.emp_stat " & _
               " from di3670 a " & _
               " left join di7670 c on a.posid=c.posid " & _
               " where (year(a.date_hire) = " & Combo1.Text & ")" & _
               " and (month(a.date_hire) = " & ListView1.SelectedItem & ") and (a.emp_stat <> 0) and (a.wap=0) and (a.paystatus <> 2) " & _
               " order by a.date_hire,a.lastname "
'         Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
    
        ShowProgress 0
        
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
'            cSqlStmt = "select * from di2660 where cmpid = " & gCompanyID
'            OpenQueryDNS cSqlStmt, objdbRs, False
'            cCmpName = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
'            cCmpAdd = IIf(objdbRs.RecordCount > 0, objdbRs("cmpaddress1"), "")

            ' ---> revision for phil no if it is blank search and post sssno
            cPhilno = oRecordSet("PHEALTHNUM")
            
            If cPhilno = "" Then
                OpenQueryDNS "select ssnum from di3670 where empid = " & cQuote & oRecordSet("EMPID") & cQuote, objdbRs, False
                cPhilno = "SSS NO - " & IIf(objdbRs.RecordCount > 0, objdbRs("SSNUM"), "")
            Else
                cPhilno = "PHIL NO - " & oRecordSet("PHEALTHNUM")
            End If
            
            
            cSqlStmt = " insert into tmpPhilH(PHEALTHNUM,EMPID,FULLNAME,POSNAME,RATE_AMT,DATE_HIRE,EMP_STAT,CMPNAME," & _
                       " ADDRESS,POSTCODE,EMPLR_ID,CERT_BY,CERT_NAME,CERT_POS)values(" & _
                        cQuote & cPhilno & cQuote & "," & _
                        cQuote & oRecordSet("EMPID") & cQuote & "," & _
                        cQuote & oRecordSet("FULLNAME") & cQuote & "," & _
                        cQuote & oRecordSet("POSNAME") & cQuote & "," & _
                        oRecordSet("RATE_AMT") & "," & _
                        cQuote & Format(oRecordSet("DATE_HIRE"), "mmddyyyy") & cQuote & "," & _
                        oRecordSet("EMP_STAT") & "," & _
                        cQuote & cCompany & cQuote & "," & _
                        cQuote & gAddress & cQuote & "," & _
                        cQuote & gPostal & cQuote & "," & _
                        cQuote & gPHealthNum & cQuote & "," & _
                        cQuote & Text6.Text & cQuote & "," & cQuote & EncodeStr2(DecodeStr(Label8.Caption)) & cQuote & "," & _
                        cQuote & aUserInfo(0) & cQuote & ")"
                        
                        
'                        oRecordSet("RATE_AMT") * 26 & "," & _

            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        GenerateReport "PHILHEALTH Er2 Report", "PRVPHEALTHEr2.RPT", , True

        ShowProgress 4
        
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub

Sub GenOLDDTRRPT(ByVal cPeriodID As String, ByVal cParam As String)
    Dim cSqlStmt, _
        oRecordSet As New ADODB.Recordset

    If Trim(cParam) <> "" Then
        cParam = "a.depid IN " & cParam
    End If
Set oRecordSet = Nothing
End Sub
' + -->
' |     Procedure Name  :   GenTMSRpt
' |     Description     :   Generate TMS Report
' |     Date Created    :   21 Mar 2011
' + -->
Sub Create_TMSRpt()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmpdtr( " & _
                       " [EMPID] char(6),       [paystatus] integer, " & _
                       " [FULLNAME] char(100),  [POSITION] char(100)," & _
                       " [DEPID] char(3),       [DEPTNAME] char(100)," & _
                       " [EMP_STAT] integer,    [active] integer," & _
                       " [SDATE] date,          [EDATE] date," & _
                       " [REG_DAY] double,      [REG_OT_HR] double,     [SA_OT_HR] double,          [TOT_OT] double," & _
                       " [ND_DAY] double,       [ND_OT_HR] double,      [ND_TOT_OT] double,         [SAND_OT_HR] double," & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [SUN_ND] double,       [SUN_ND_OT] double, " & _
                       " [HOLIDAY] double, " & _
                       " [SIGNATORY1] char(50),     [POSNAME1] char(50)," & _
                       " [SIGNATORY2] char(50),     [POSNAME2] char(50)," & _
                       " [SIGNATORY3] char(50),     [POSNAME3] char(50))"
                       
                       

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpdtr"
    QueryTemp cSqlStmt, oTempADO, True

End Sub


Sub GenTMSRpt(ByVal cPeriodID As String, ByVal cParam As String)
    Dim cSqlStmt As String, _
        aUserInfo As Variant, _
        cCmpname As String, _
        cCmpAdd As String, _
        oRecordSet As New ADODB.Recordset
    Dim D_Start, D_End As String
    Dim nCtr As Integer
    aUserInfo = Array("", "", "")

    If Not ChkPersonnel(Text6) Then Exit Sub
    'If Not ChkPersonnel(Text5) Then Exit Sub
    If Not ChkPersonnel(Text1) Then Exit Sub
    
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aUserInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text1.Text & "'"
        aUserInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If

    Create_TMSRpt
    
    ShowProgress 0
    
    OpenQueryDNS "SELECT DATE_START, DATE_END FROM PA7730 WHERE PERIODID=" & cQuote & cPeriodID & cQuote, objdbRs, False
    
    If objdbRs.RecordCount > 0 Then
        D_Start = objdbRs("DATE_START")
        D_End = objdbRs("DATE_END")
    Else
        D_Start = ""
        D_End = ""
    End If

    cSqlStmt = "select a.tcid, " & _
               "       a.empid, " & _
               "       concat(a.lastname,', ',a.firstname,' ',if(trim(a.mname)='',' ',concat(left(a.mname,1),'.'))) as fullname, " & _
               "       ifnull(b.posname,'') as position, a.emp_stat, " & _
               "       a.firstname, a.lastname, ifnull(c.linename,'') as linename, " & _
               "       a.paystatus, a.active, " & _
               "       round(ifnull(sum(d.reg_hr)/8,0),3) as reg_day, " & _
               "       round(ifnull(sum(d.reg_ot_hr),0),3) as reg_ot, " & _
               "       round(ifnull(sum(d.sa_reg_ot),0),3) as sa_reg_ot, " & _
               "       round(ifnull(sum(d.tot_ot),0),3) as tot_ot, " & _
               "       round(ifnull(sum(d.nd_hr)/8,0),3) as nd_day, " & _
               "       round(ifnull(sum(d.nd_ot_hr),0),3) as nd_ot, " & _
               "       round(ifnull(sum(d.sa_nd_ot),0),3) as sa_nd_ot, " & _
               "       round(ifnull(sum(d.nd_tot_ot),0),3) as nd_tot_ot, " & _
               "       round(ifnull(sum(d.sun_hr),0),3) as sun_hr, round(ifnull(sum(d.sun_ot_hr),0),3) as sun_ot, " & _
               "       a.depid,a.wap, " & _
               "       round(ifnull(sum(d.sun_nd),0),3) as sun_nd, round(ifnull(sum(d.sun_nd_ot),0),3) as sun_nd_ot, " & _
               "       round(ifnull(sum(d.inc_hr),0),3) as inc_hr " & _
               "from di3670 a   left join " & IIf(Check1.Value <> 0, "di36770", "dih36770") & " d on a.empid=d.empid and d.date between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & _
               " left join di7670 b on a.posid=b.posid " & _
               " left join di5463 c on a.depid=c.lineid " & _
               " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ")))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & "))"
    
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt & " group by a.empid order by a.lastname,a.firstname", oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst
        ShowProgress 0
        
        While Not oRecordSet.EOF
            
            ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100
            
                    cSqlStmt = "insert into tmpdtr(empid, fullname, [position], " & _
                               " deptname, paystatus, emp_stat, [active], sdate, edate, " & _
                               " reg_day, reg_ot_hr, sa_ot_hr, tot_ot, nd_day, nd_ot_hr, sand_ot_hr, nd_tot_ot, sun, sunot, sun_nd, sun_nd_ot, holiday, " & _
                               " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                               cQuote & oRecordSet("EMPID") & cQuote & "," & _
                               cQuote & EncodeStr2(oRecordSet("FULLNAME")) & cQuote & "," & _
                               cQuote & EncodeStr2(oRecordSet("POSITION")) & cQuote & "," & _
                               cQuote & EncodeStr2(oRecordSet("LINENAME")) & cQuote & "," & _
                               Val(oRecordSet("PAYSTATUS")) & "," & _
                               Val(oRecordSet("EMP_STAT")) & "," & _
                               Val(oRecordSet("ACTIVE")) & "," & _
                               cQuote & Format(D_Start, "mm/dd/yyyy") & cQuote & "," & _
                               cQuote & Format(D_End, "mm/dd/yyyy") & cQuote & "," & _
                               Val(oRecordSet("REG_DAY")) & "," & _
                               Val(oRecordSet("REG_OT")) & "," & _
                               Val(oRecordSet("SA_REG_OT")) & "," & _
                               Val(oRecordSet("TOT_OT")) & "," & _
                               Val(oRecordSet("ND_DAY")) & "," & _
                               Val(oRecordSet("ND_OT")) & "," & _
                               Val(oRecordSet("SA_ND_OT")) & "," & _
                               Val(oRecordSet("ND_TOT_OT")) & "," & _
                               Val(oRecordSet("SUN_HR")) & "," & _
                               Val(oRecordSet("SUN_OT")) & "," & _
                               Val(oRecordSet("SUN_ND")) & "," & _
                               Val(oRecordSet("SUN_ND_OT")) & ",0," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                        
                        
            QueryTemp cSqlStmt, objdbRs, True
            
            oRecordSet.MoveNext
            
        Wend
        
        ShowProgress 3
        
        Select Case Combo1.ListIndex
            Case 0
                GenerateReport "Daily Time Report (Summary)", "RPT387.RPT", , True
            Case 1
                GenerateReport "Daily Time Report (Summary)", IIf(Check7.Value = vbChecked, "RPT387AR.RPT", "RPT387AR_SUN.RPT"), , True
            Case 2
                GenerateReport "Extension Daily Time Report", "rpt387E.rpt", , True
                
        End Select
        ShowProgress 4
        
    Else
        MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    End If
    
    Set oRecordSet = Nothing
End Sub

' + -->
' |     Procedure Name  :   GenTMSSumRpt
' |     Description     :   Generate TMS Summary Report
' |     Date Created    :   22 Mar 2011
' + -->
Sub Create_TMSSumRpt(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt, cTableName As String
    
    Select Case nMode
        Case 1
            cSqlStmt = " CREATE TABLE tmp84650( " & _
                       " [wap] integer,         [paystatus] integer, " & _
                       " [EMPID] char(6),       [TRAN_NO] char(10)," & _
                       " [FULLNAME] char(100),  [DEPTNAME] char(100)," & _
                       " [DAY_DATE] date,       [DAY_NAME] char(20)," & _
                       " [RegHour] double,      [OTHour] double, " & _
                       " [SAOT] double,         [NDiff] double, " & _
                       " [NDiffOT] double,      [SANDOT] double, " & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [SUN_ND] double,       [SUN_ND_OT] double, " & _
                       " [LOGDATE] date,        [TRANSDATE] date," & _
                       " [SHIFTDESC] char(100), [REMARK] char(100)," & _
                       " [TIME1] char(15),      [TIME2] char(15), " & _
                       " [SEQ_NO] integer,      [emp_stat] integer, " & _
                       " [TOT_OT] double,       [ND_TOT_OT] double, " & _
                       " [SIGNATORY1] char(50), [POSNAME1] char(50)," & _
                       " [SIGNATORY2] char(50), [POSNAME2] char(50)," & _
                       " [SIGNATORY3] char(50), [POSNAME3] char(50))"
            
            cTableName = "tmp84650"
        
        Case 2
             cSqlStmt = " CREATE TABLE tmpDTRD(  [paystatus] integer, " & _
                       " [emp_stat] integer,    [wap] integer," & _
                       " [EMPID] char(6),       [TRAN_NO] char(10), " & _
                       " [FULLNAME] char(100),  [DEPTNAME] char(100), " & _
                       " [DAY_DATE] date,       [DAY_NAME] char(20), " & _
                       " [RegHour] double,      [OTHour] double, " & _
                       " [SAOT] double,         [NDiff] double, " & _
                       " [NDiffOT] double,      [SANDOT] double, " & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [SUN_ND] double,       [SUN_ND_OT] double, " & _
                       " [LOGDATE] date,        [TRANSDATE] date," & _
                       " [outtrantime] char(15),[intrantime] char(15), " & _
                       " [SHIFTDESC] char(100), [REMARK] char(100)," & _
                       " [TIME1] char(15),      [TIME2] char(15)," & _
                       " [SEQ_NO] integer,      [tag] integer, " & _
                       " [periodid] char(5),    [Duration] char(100)," & _
                       " [TOT_OT] double,       [ND_TOT_OT] double, " & _
                       " [SIGNATORY1] char(50), [POSNAME1] char(50)," & _
                       " [SIGNATORY2] char(50), [POSNAME2] char(50)," & _
                       " [SIGNATORY3] char(50), [POSNAME3] char(50))"
            
            cTableName = "tmpDTRD"

                    
    End Select
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM " & cTableName
    QueryTemp cSqlStmt, oTempADO, True

End Sub


Sub GenTMSSumRpt(ByVal cPeriodID As String, ByVal cParam As String)
    Dim cSqlStmt As String, _
        aUserInfo As Variant, _
        cCmpname As String, _
        cCmpAdd As String, _
        oRecordSet As New ADODB.Recordset
    Dim D_Start, D_End As String
    Dim nCtr As Integer
    
    Dim oRSet As New ADODB.Recordset
    
    Dim cDepid As String, _
        nCtr1 As Integer, _
        aTInfo As Variant, _
        aTimeInfo As Variant, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        aTimeDtrVal As Variant, _
        dLogDate As Date, _
        lWap As Boolean
        
    Dim oRset1 As New ADODB.Recordset

    aUserInfo = Array("", "", "")

    If Not ChkPersonnel(Text6) Then Exit Sub
'    If Not ChkPersonnel(Text5) Then Exit Sub
    If Not ChkPersonnel(Text1) Then Exit Sub
    
    OpenQueryDNS "SELECT * FROM PA2360 ORDER BY USERID", objdbRs, False
    If objdbRs.RecordCount > 0 Then
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text6.Text & "'"
        aUserInfo(0) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text5.Text & "'"
        aUserInfo(1) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    
        objdbRs.Requery adAsyncFetch
        objdbRs.Find "USERID='" & Text1.Text & "'"
        aUserInfo(2) = IIf(Not objdbRs.EOF, objdbRs("POSITION"), "")
    End If

    Create_TMSSumRpt nTagSelect
    
    ShowProgress 0
    
    OpenQueryDNS "SELECT DATE_START, DATE_END FROM PA7730 WHERE PERIODID=" & cQuote & cPeriodID & cQuote, objdbRs, False
    
    If objdbRs.RecordCount > 0 Then
        D_Start = objdbRs("DATE_START")
        D_End = objdbRs("DATE_END")
    Else
        D_Start = ""
        D_End = ""
    End If
    
    cSqlStmt = "select a.tcid, " & _
               "       a.empid, " & _
               "       concat(a.lastname,', ',a.firstname,' ',if(trim(a.mname)='',' ',concat(left(a.mname,1),'.'))) as fullname, " & _
               "       ifnull(b.posname,'') as position, a.emp_stat, " & _
               "       a.firstname, a.lastname, c.linename, " & _
               "       a.paystatus, a.active, " & _
               "       round(ifnull(sum(d.reg_hr)/8,0),3) as reg_day, " & _
               "       round(ifnull(sum(d.reg_ot_hr),0),3) as reg_ot, " & _
               "       round(ifnull(sum(d.sa_reg_ot),0),3) as sa_reg_ot, " & _
               "       round(ifnull(sum(d.tot_ot),0),3) as tot_ot, " & _
               "       round(ifnull(sum(d.nd_hr)/8,0),3) as nd_day, " & _
               "       round(ifnull(sum(d.nd_ot_hr),0),3) as nd_ot, " & _
               "       round(ifnull(sum(d.sa_nd_ot),0),3) as sa_nd_ot, " & _
               "       round(ifnull(sum(d.nd_tot_ot),0),3) as nd_tot_ot, " & _
               "       round(ifnull(sum(d.sun_hr),0),3) as sun_hr, round(ifnull(sum(d.sun_ot_hr),0),3) as sun_ot, " & _
               "       a.depid,a.wap, " & _
               "       round(ifnull(sum(d.sun_nd),0),3) as sun_nd, round(ifnull(sum(d.sun_nd_ot),0),3) as sun_nd_ot, " & _
               "       round(ifnull(sum(d.inc_hr),0),3) as inc_hr " & _
               "from di3670 a   left join " & IIf(Check1.Value <> 0, "di36770", "dih36770") & " d on a.empid=d.empid and d.date between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & _
               " left join di7670 b on a.posid=b.posid " & _
               " left join di5463 c on a.depid=c.lineid " & _
               " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ")))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & "))"

    
    'Script2File cSqlStmt
    OpenQueryDNS cSqlStmt & " group by a.empid order by a.lastname,a.firstname", oRSet, False
    If oRSet.RecordCount > 0 Then
        'oRSet.MoveFirst
        
        While oRSet.EOF = False
            
'            If oRSet("empid") = "255625" Then MsgBox "stop"
            ShowProgress 2, (oRSet.AbsolutePosition / oRSet.RecordCount) * 100
        
            If (InStr(1, cParam, oRSet("DEPID"), vbTextCompare)) Or (Trim(cParam) = "") Then
                
                aShiftInfo = Array("", "", "", "")
                aTrantype = Array("", "", "", "")
                
                aTimeDtrVal = Array(0#, 0#)
                
                If nTagSelect = 1 Then
                    cSqlStmt = " select distinct a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2 from " & IIf(Check1.Value <> 0, "pa", "pah") & "84650 a " & _
                                       " left join pa74380 b on a.shiftid = b.shiftid " & _
                                       " where (a.empid=" & cQuote & oRSet("EMPID") & cQuote & _
                                       ") and (a.logdate between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ")"
                Else
                    cSqlStmt = " select a.empid, a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2," & _
                               " a.tran_no,a.transdate,date_format(a.transdate,'%a - %b %e, %Y') as `day`,trantype,if(a.trantype=0,'In','Out') as trn_type,a.trantime " & _
                               " from " & IIf(Check1.Value <> 0, "pa", "pah") & "84650 a left join pa74380 b on a.shiftid = b.shiftid " & _
                               " where (a.empid=" & cQuote & oRSet("EMPID") & cQuote & _
                               ") and (a.logdate between " & cQuote & Format(D_Start, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(D_End, "yyyy-mm-dd") & cQuote & ")" & _
                               " order by a.logdate,a.transdate, a.trantime"
                End If
                
'                Script2File cSqlStmt
                OpenQueryDNS cSqlStmt, oRset1, False
    
                If oRset1.RecordCount > 0 Then
                    oRset1.MoveFirst
                    
                     While Not oRset1.EOF
                        cSqlStmt = "select a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.tot_ot, a.nd_hr, a.nd_ot_hr, a.nd_tot_ot, a.sun_hr, a.sun_ot_hr, " & _
                                                       " 0, 0, 0, a.tag, a.sa_nd_ot, a.sun_nd, a.sun_nd_ot, a.remark,b.description,b.time1,b.time2 " & _
                                                       "from " & IIf(Check1.Value <> 0, "di", "dih") & "36770  a " & _
                                                       " left join pa74380 b on a.shiftid = b.shiftid " & _
                                                       "where (a.empid=" & cQuote & oRSet("EMPID") & cQuote & ")" & _
                                                       " and (a.date=" & cQuote & Format(oRset1("logdate"), "yyyy-mm-dd") & cQuote & ")"
'                        Script2File cSqlStmt
                        OpenQueryDNS cSqlStmt, objdbRs, False
                
                          aTimeInfo = Array(IIf(objdbRs.RecordCount = 0, 0, objdbRs("reg_hr")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("reg_ot_hr")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sa_reg_ot")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("nd_hr")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("nd_ot_hr")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sun_hr")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sun_ot_hr")), _
                                                          0, _
                                                          0, _
                                                          0, _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("tag")), _
                                                          IIf(objdbRs.RecordCount = 0, "", objdbRs("remark")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sa_nd_ot")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sun_nd")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("sun_nd_ot")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("tot_ot")), _
                                                          IIf(objdbRs.RecordCount = 0, 0, objdbRs("nd_tot_ot")))
                
                        If nTagSelect = 1 Then
                             cSqlStmt = " insert into tmp84650(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                           " RegHour," & _
                                                           " OTHour," & _
                                                           " SAOT," & _
                                                           " NDiff," & _
                                                           " NDiffOT," & _
                                                           " SANDOT," & _
                                                           " SUN,SUNOT," & _
                                                           " SUN_ND,SUN_ND_OT," & _
                                                           " TOT_OT,ND_TOT_OT," & _
                                                           "SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no, " & _
                                                           " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                                           cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                                           cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                                           cQuote & Format(oRset1("logdate"), "mm/dd/yyyy") & cQuote & "," & cQuote & Format(oRset1("logdate"), "dddd") & cQuote & "," & _
                                                           aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                                           aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                                           aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                                           aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                                           aTimeInfo(15) & "," & aTimeInfo(16) & "," & _
                                                           cQuote & EncodeStr2(oRset1("description")) & cQuote & "," & _
                                                           cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                                           cQuote & objdbRs("time1") & cQuote & "," & cQuote & objdbRs("time2") & cQuote & "," & _
                                                           cQuote & Format(oRset1("logdate"), "mm/dd/yyyy") & cQuote & "," & _
                                                           nCtr & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                                           
                            QueryTemp cSqlStmt, objdbRs, True
                                                       
                        Else
                            
                            aTrantype(3) = oRset1("TRANSDATE")
                            If oRset1("trantype") = 0 Then

                                If Trim(aTrantype(1)) = "" Then
                                    aTrantype(0) = oRset1("trantype")
                                    aTrantype(1) = oRset1("trantime")
                                    dLogDate = oRset1("logdate")
                                End If

                            Else
                                aTrantype(0) = oRset1("trantype")
                                aTrantype(2) = oRset1("trantime")
'                                If gCompanyID <> "0003" Then
                                    aShiftInfo(0) = oRset1("description")
                                    aShiftInfo(1) = oRset1("time1")
                                    aShiftInfo(2) = oRset1("time2")
'                                Else
'                                    aShiftInfo(0) = objdbRs("description")
'                                    aShiftInfo(1) = objdbRs("time1")
'                                    aShiftInfo(2) = objdbRs("time2")
'                                End If
                                dLogDate = oRset1("logdate")
                            End If
                        End If
                        
                       oRset1.MoveNext
                       
                       If nTagSelect <> 1 Then
                        
                            If Not oRset1.EOF Then
                                If dLogDate = oRset1("logdate") Then
                                    If (oRset1("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                        If gDepid <> oRSet("DEPID") Then
                                           
                                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                       " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                       " LOGDATE,TRANSDATE, " & _
                                                       " intrantime,outtrantime," & _
                                                       " SHIFTDESC,REMARK," & _
                                                       " TIME1,TIME2," & _
                                                       " tag,SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                                       " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                                       cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                                       cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                       aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                                       aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                                       aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                                       aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                       cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                       cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                       cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                                       cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                       aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                                       

                                            QueryTemp cSqlStmt, objdbRs, True
                                            aTrantype = Array("", "", "", "")
                                        Else
                                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                       " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                       " LOGDATE,TRANSDATE, " & _
                                                       " intrantime,outtrantime," & _
                                                       " SHIFTDESC,REMARK," & _
                                                       " TIME1,TIME2," & _
                                                       " tag,SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                                       " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                                       cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                                       cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                       aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                                       aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                                       aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                                       aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                                       cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                       cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                       cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                       cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                                       cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                       aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                                       
                                            QueryTemp cSqlStmt, objdbRs, True
                                            aTrantype = Array("", "", "", "")
                                        End If
'                                        aTrantype = Array("", "", "", "")
                                    End If
'                                    aTrantype = Array("", "", "", "")
                                Else
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                               " LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime," & _
                                               " SHIFTDESC,REMARK," & _
                                               " TIME1,TIME2," & _
                                               " tag, SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                               " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                               cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                               cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                               aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                               aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                               aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                               aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                               aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                               
                                    QueryTemp cSqlStmt, objdbRs, True
                                    aTrantype = Array("", "", "", "")

                                End If
                            Else
                                If gDepid <> oRSet("DEPID") Then
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME, " & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                               " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                               cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                               cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                               aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                               aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                               aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                               aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                               nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                               
                                    QueryTemp cSqlStmt, objdbRs, True
                                    aTrantype = Array("", "", "", "")
                                Else
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                               " LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime," & _
                                               " SHIFTDESC,REMARK," & _
                                               " TIME1,TIME2," & _
                                               " tag,SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                               " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                               cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                               cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                               aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                               aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                               aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                               aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                               aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & "," & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                               
                                    QueryTemp cSqlStmt, objdbRs, True
                                    aTrantype = Array("", "", "", "")
                                End If
                                
                                aTrantype = Array("", "", "", "")
                            End If
                        End If
                       
                       
                    Wend
                Else
                    If nTagSelect = 1 Then

                        cSqlStmt = " insert into tmp84650(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME,RegHour," & _
                                       " OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,TOT_OT,ND_TOT_OT,SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no, " & _
                                       " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                       cQuote & oRSet("EMPID") & cQuote & "," & _
                                       cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                       cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & _
                                       Val(oRSet("EMP_STAT")) & "," & _
                                       Val(oRSet("WAP")) & "," & _
                                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                                       "'',0,0,0,0,0,0,0,0,0,0,0,0,'','','',''," & _
                                       cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & nCtr & ", " & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"

                        QueryTemp cSqlStmt, objdbRs, True
                    Else
                        If gCompanyID <> "0003" Then '20080328 custom setting for mico only
                            If oRSet("EMP_STAT") <> 0 And oRSet("PAYSTATUS") = 0 Then
                                cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME, " & _
                                           " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,LOGDATE,TRANSDATE, " & _
                                           " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO,TOT_OT,ND_TOT_OT, " & _
                                           " [SIGNATORY1], [SIGNATORY2], [SIGNATORY3], [POSNAME1], [POSNAME2], [POSNAME3])values(" & _
                                           cQuote & oRSet("EMPID") & cQuote & "," & cQuote & oRSet("FULLNAME") & cQuote & "," & _
                                           cQuote & oRSet("LINENAME") & cQuote & "," & Val(oRSet("PAYSTATUS")) & "," & Val(oRSet("EMP_STAT")) & "," & Val(oRSet("WAP")) & "," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                           "0,0,0,0,0,0,0,0,0,0," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & _
                                           "'','','','','',''," & _
                                           nCtr & ",0,0, " & cQuote & Label8.Caption & cQuote & "," & cQuote & Label6.Caption & cQuote & "," & cQuote & Label4.Caption & cQuote & "," & cQuote & aUserInfo(0) & cQuote & "," & cQuote & aUserInfo(1) & cQuote & "," & cQuote & aUserInfo(2) & cQuote & ")"
                                           
                                QueryTemp cSqlStmt, objdbRs, True
                                aTrantype = Array("", "", "", "")
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            oRSet.MoveNext
            
            
        Wend
      End If
        
        ShowProgress 3
        
        QueryTemp "select * from " & IIf(nTagSelect = 1, "tmp84650", "tmpDTRD"), objdbRs, False
        
        If objdbRs.RecordCount > 0 Then
        
            If nTagSelect <> 1 Then
                QueryTemp "select * from tmpDTRD where intrantime=''", objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    QueryTemp "delete from tmpDTRD where intrantime=''", objdbRs, True
                End If

                QueryTemp "select * from tmpDTRD where outtrantime=''", objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    QueryTemp "delete from tmpDTRD where outtrantime=''", objdbRs, True
                End If
                
                Select Case Combo1.ListIndex
                    Case 0
                        GenerateReport "Daily Time Report", "PRV376.RPT", , True
                    Case 1
                        GenerateReport "Daily Time Report", IIf(Check7.Value = vbChecked, "PRV376AR.RPT", "PRV376AR_SUN.RPT"), , True
                    Case 2
                        GenerateReport "Extension Daily Time Report", "PRV376E.rpt", , True
                        
                End Select
            Else
                Select Case Combo1.ListIndex
                    Case 0
                        GenerateReport "Daily Time Report", "PRV377.RPT", , True
                    Case 1
                        GenerateReport "Daily Time Report", IIf(Check7.Value = vbChecked, "PRV377AR.RPT", "PRV377ARSUN.RPT"), , True
                    Case 2
                        GenerateReport "Extension Daily Time Report", "PRV377E.rpt", , True
                        
                End Select
            End If
        
        Else
            MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
        End If
        
        ShowProgress 4
    
    Set oRecordSet = Nothing
End Sub


' + -->
' |     Procedure Name  :   GenERPSAL
' |     Description     :   Print Utility for ERP SALARY
' |     Date Created    :   10 Sep 2012
' + -->

Sub CreateTmpPayERPSAL()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
        
    cSqlStmt = " CREATE TABLE tmpERPSAL( [BUKRS] CHAR(4), " & _
               " [GJAHR] CHAR(4), " & _
               " [BUDAT] CHAR(100), " & _
               " [BLDAT] CHAR(100), " & _
               " [WAERS] CHAR(10), " & _
               " [FLAG_A] CHAR(5), " & _
               " [FLAG_B] CHAR(100), " & _
               " [BKTXT] CHAR(100), " & _
               " [WRBTR] double, " & _
               " [SGTXT] CHAR(100), " & _
               " [KOSTL] CHAR(100), " & _
               " [SEQ_NO] INTEGER)"
    
    
    'MsgBox cSqlStmt
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    cSqlStmt = "DELETE FROM tmpERPSAL"
    QueryTemp cSqlStmt, oTempADO, True
    
End Sub


Sub GenERPSAL(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
    Dim cSqlStmt As String, _
        cPeriodName As String, _
        cSalaryName As String, _
        cDateSalary As String
    Dim nSubGross, _
        nSubded, _
        nGtotal As Double
        
    Dim cParam1, cParam2
        
    CreateTmpPayERPSAL

    If Trim(cParam) <> "" Then
        cParam1 = "a.costcenterid IN " & cParam
    End If


    'Regular and contractual and SA

    cSqlStmt = " select a.PERIODID,a.COSTCENTERID, ifnull(b.DESCRIPTION,'') as DESCRIPTION, " & _
               " sum(a.GROSS_PAY) as GROSS_PAY, sum(a.NET_PAY) as NET_PAY, sum(a.SA_NET_PAY) as SA_NET_PAY, " & _
               " a.WORKCENTERID from pa87260 a " & _
               " left join pa37722 b on a.costcenterid=b.costcenterid " & _
               " where periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam1 <> "", " and (" & cParam1 & ")", "") & _
               " and a.emp_stat <> 0 and a.paystatus=0 " & _
               " group by a.emp_stat, a.costcenterid " & _
               " order by a.emp_stat "

'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    If oTempADO.RecordCount > 0 Then


        ShowProgress 0

        While Not oTempADO.EOF

            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100

            cSqlStmt = "SELECT PERIODID, DATE_START, DATE_END, DURATION, STATUS FROM pa7730 where periodid=" & cQuote & oTempADO("periodid") & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                If objdbRs("status") = 1 Then
                    cPeriodName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Doc"
                    cSalaryName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Payment " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
                    cDateSalary = Year(objdbRs("date_start")) & "." & Format(objdbRs("date_start"), "MM") & ".25"
                    
'                    MsgBox cDateSalarys
                    
                Else
                    cPeriodName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Doc"
                    cSalaryName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Payment " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
                    cDateSalary = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & ".10"
                End If
            End If

            If oTempADO("gross_pay") <> 0 Then

                cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                            cQuote & nCompCode & cQuote & "," & _
                            cQuote & Format(Now, "YYYY") & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & "PHP" & cQuote & "," & _
                            cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                            cQuote & " REG_CON " & cQuote & "," & _
                            cQuote & cPeriodName & cQuote & "," & _
                            IIf(gAgency = 0, oTempADO("gross_pay"), oTempADO("net_pay")) & "," & _
                            cQuote & cSalaryName & cQuote & "," & _
                            cQuote & oTempADO("costcenterid") & cQuote & ",1)"

                QueryTemp cSqlStmt, objdbRs, True
            End If

            If oTempADO("SA_NET_PAY") <> 0 Then
                cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                            cQuote & nCompCode & cQuote & "," & _
                            cQuote & Format(Now, "YYYY") & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & "PHP" & cQuote & "," & _
                            cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                            cQuote & " REG_CON_SA " & cQuote & "," & _
                            cQuote & cPeriodName & cQuote & "," & _
                            oTempADO("SA_NET_PAY") & "," & _
                            cQuote & cSalaryName & cQuote & "," & _
                            cQuote & oTempADO("costcenterid") & cQuote & ",2)"

                QueryTemp cSqlStmt, objdbRs, True

            End If
            oTempADO.MoveNext

        Wend

    Else
        GoTo FINANCEERP

    End If
    
'***************************************************************************************************************************************

    'WAP and WAP SA

'    If gAgency = 0 Then

        cSqlStmt = " select a.PERIODID,a.COSTCENTERID, ifnull(b.DESCRIPTION,'') as DESCRIPTION, " & _
                   " sum(a.GROSS_PAY) as GROSS_PAY, sum(a.NET_PAY) as NET_PAY, sum(a.SA_NET_PAY) as SA_NET_PAY, " & _
                   " a.WORKCENTERID from pa87260 a " & _
                   " left join pa37722 b on a.costcenterid=b.costcenterid " & _
                   " where periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam1 <> "", " and (" & cParam1 & ")", "") & _
                   " and a.emp_stat = 0 and a.paystatus=0 " & _
                   " group by a.emp_stat, a.costcenterid " & _
                   " order by a.emp_stat "

        'Script2File cSqlStmt
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then


            ShowProgress 0

            While Not oTempADO.EOF

                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100

                cSqlStmt = "SELECT PERIODID, DATE_START, DATE_END, DURATION, STATUS FROM pa7730 where periodid=" & cQuote & oTempADO("periodid") & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    If objdbRs("status") = 1 Then
                        cPeriodName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Doc"
                        cSalaryName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Payment"
                        cDateSalary = Year(objdbRs("date_start")) & "." & Format(objdbRs("date_start"), "MM") & ".25"
                    Else
                        cPeriodName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Doc"
                        cSalaryName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Payment " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
                        cDateSalary = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & ".10"
                        
                    End If
                End If

                If oTempADO("gross_pay") <> 0 Then

                    cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                                cQuote & nCompCode & cQuote & "," & _
                                cQuote & Format(Now, "YYYY") & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & "PHP" & cQuote & "," & _
                                cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                                cQuote & " WAP " & cQuote & "," & _
                                cQuote & cPeriodName & cQuote & "," & _
                                oTempADO("gross_pay") & "," & _
                                cQuote & cSalaryName & cQuote & "," & _
                                cQuote & oTempADO("costcenterid") & cQuote & ",3)"

                    QueryTemp cSqlStmt, objdbRs, True

                End If

                If oTempADO("SA_NET_PAY") <> 0 Then
                    cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                                cQuote & nCompCode & cQuote & "," & _
                                cQuote & Format(Now, "YYYY") & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & "PHP" & cQuote & "," & _
                                cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                                cQuote & " WAP_SA " & cQuote & "," & _
                                cQuote & cPeriodName & cQuote & "," & _
                                oTempADO("SA_NET_PAY") & "," & _
                                cQuote & cSalaryName & cQuote & "," & _
                                cQuote & oTempADO("costcenterid") & cQuote & ",4)"

                    QueryTemp cSqlStmt, objdbRs, True
                End If
                oTempADO.MoveNext

            Wend
        End If


    '***************************************************************************************************************************************

        'Emergency and Emergency SA

        cSqlStmt = " select a.PERIODID,a.COSTCENTERID, ifnull(b.DESCRIPTION,'') as DESCRIPTION, " & _
                   " sum(a.GROSS_PAY) as GROSS_PAY, sum(a.NET_PAY) as NET_PAY, sum(a.SA_NET_PAY) as SA_NET_PAY, " & _
                   " a.WORKCENTERID from pa87260 a " & _
                   " left join pa37722 b on a.costcenterid=b.costcenterid " & _
                   " where periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam1 <> "", " and (" & cParam1 & ")", "") & _
                   " and a.paystatus=2 " & _
                   " group by a.emp_stat, a.costcenterid " & _
                   " order by a.emp_stat "

'        Script2File cSqlStmt
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then


            ShowProgress 0

            While Not oTempADO.EOF

                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100

                cSqlStmt = "SELECT PERIODID, DATE_START, DATE_END, DURATION, STATUS FROM pa7730 where periodid=" & cQuote & oTempADO("periodid") & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    If objdbRs("status") = 1 Then
                        cPeriodName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Doc"
                        cSalaryName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Payment"
                        cDateSalary = Year(objdbRs("date_start")) & "." & Format(objdbRs("date_start"), "MM") & ".25"
                    Else
                        cPeriodName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Doc"
                        cSalaryName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Payment " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
                        cDateSalary = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & ".10"
                        
                    End If
                End If

'                If oTempADO("gross_pay") <> 0 Then
'
'
'                    cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
'                                cQuote & nCompCode & cQuote & "," & _
'                                cQuote & Format(Now, "YYYY") & cQuote & "," & _
'                                cQuote & cDateSalary & cQuote & "," & _
'                                cQuote & cDateSalary & cQuote & "," & _
'                                cQuote & "PHP" & cQuote & "," & _
'                                cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
'                                cQuote & " Emergency " & cQuote & "," & _
'                                cQuote & cPeriodName & cQuote & "," & _
'                                oTempADO("gross_pay") & "," & _
'                                cQuote & cSalaryName & cQuote & "," & _
'                                cQuote & oTempADO("costcenterid") & cQuote & ",5)"
'
'                    QueryTemp cSqlStmt, objdbRs, True
'                End If

                 If oTempADO("gross_pay") <> 0 Then
                 
                     If gAgency = 0 Then
                     
                        cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                                    cQuote & nCompCode & cQuote & "," & _
                                    cQuote & Format(Now, "YYYY") & cQuote & "," & _
                                    cQuote & cDateSalary & cQuote & "," & _
                                    cQuote & cDateSalary & cQuote & "," & _
                                    cQuote & "PHP" & cQuote & "," & _
                                    cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                                    cQuote & " Emergency " & cQuote & "," & _
                                    cQuote & cPeriodName & cQuote & "," & _
                                    oTempADO("gross_pay") & "," & _
                                    cQuote & cSalaryName & cQuote & "," & _
                                    cQuote & oTempADO("costcenterid") & cQuote & ",5)"
    
                        QueryTemp cSqlStmt, objdbRs, True
                    Else
                        cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                                    cQuote & nCompCode & cQuote & "," & _
                                    cQuote & Format(Now, "YYYY") & cQuote & "," & _
                                    cQuote & cDateSalary & cQuote & "," & _
                                    cQuote & cDateSalary & cQuote & "," & _
                                    cQuote & "PHP" & cQuote & "," & _
                                    cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                                    cQuote & " Emergency " & cQuote & "," & _
                                    cQuote & cPeriodName & cQuote & "," & _
                                    oTempADO("net_pay") & "," & _
                                    cQuote & cSalaryName & cQuote & "," & _
                                    cQuote & oTempADO("costcenterid") & cQuote & ",5)"
    
                        QueryTemp cSqlStmt, objdbRs, True
                    End If
                    
                End If





                If oTempADO("SA_NET_PAY") <> 0 Then
                    cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                                cQuote & nCompCode & cQuote & "," & _
                                cQuote & Format(Now, "YYYY") & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & cDateSalary & cQuote & "," & _
                                cQuote & "PHP" & cQuote & "," & _
                                cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                                cQuote & " Emergency_SA " & cQuote & "," & _
                                cQuote & cPeriodName & cQuote & "," & _
                                oTempADO("SA_NET_PAY") & "," & _
                                cQuote & cSalaryName & cQuote & "," & _
                                cQuote & oTempADO("costcenterid") & cQuote & ",6)"

                    QueryTemp cSqlStmt, objdbRs, True
                End If

                oTempADO.MoveNext

            Wend
        End If

'    End If

    '*********************************

    If Trim(cParam) <> "" Then
        cParam2 = "b.costcenterid IN " & cParam
    End If

    If gAgency = 0 Then
        cSqlStmt = " select a.PERIODID, b.COSTCENTERID,a.DEDID,ifnull(c.DEDNAME,'') as DEDNAME,c.DEDERPID, " & _
                   " sum(a.DED_AMT) as DED_AMT, sum(a.DED_AMT2) as DED_AMT2, sum(a.DED_AMT3) as DED_AMT3, " & _
                   " sum(b.GROSS_PAY) as GROSS_PAY, sum(b.NET_PAY) as NET_PAY, sum(b.SA_NET_PAY) as SA_NET_PAY " & _
                   " from pa87263 a " & _
                   " left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
                   " left join pa3330 c on a.dedid=c.dedid " & _
                   " where a.periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam2 <> "", " and (" & cParam2 & ")", "") & _
                   IIf(gAgency = 0, " group by a.dedid", " group by a.periodid")
    
    '    Script2File cSqlStmt
    '    MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, oTempADO, False
        If oTempADO.RecordCount > 0 Then
    
    
            ShowProgress 0
    
            While Not oTempADO.EOF
    
                ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100
    
                cSqlStmt = "SELECT PERIODID, DATE_START, DATE_END, DURATION, STATUS FROM pa7730 where periodid=" & cQuote & oTempADO("periodid") & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    If objdbRs("status") = 1 Then
                        cPeriodName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Doc"
                        cSalaryName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st " & IIf(gAgency = 0, oTempADO("DEDNAME"), "Decution (" & cODBC & ")")
                        cDateSalary = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) & ".25"
                    Else
                        cPeriodName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Doc"
                        cSalaryName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd " & IIf(gAgency = 0, oTempADO("DEDNAME"), "Decution (" & cODBC & ")")
                        cDateSalary = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & ".10"
                    End If
                End If
    
                cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                            cQuote & nCompCode & cQuote & "," & _
                            cQuote & Format(Now, "YYYY") & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & cDateSalary & cQuote & "," & _
                            cQuote & "PHP" & cQuote & "," & _
                            cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                            cQuote & " " & cQuote & "," & _
                            cQuote & cPeriodName & cQuote & "," & _
                            -oTempADO("DED_AMT") & "," & _
                            cQuote & cSalaryName & cQuote & "," & _
                            cQuote & IIf(gAgency = 0, IIf(oTempADO("DEDID") <> "", oTempADO("DEDID"), oTempADO("DEDERPID")), "998") & cQuote & ",7)"
    
                QueryTemp cSqlStmt, objdbRs, True
    
                oTempADO.MoveNext
    
            Wend
    
    '        ShowProgress 3
    '
    '        GenerateReport "Daily Time Report (Summary)", "rptFinance.RPT", , True
    '
    '        ShowProgress 4
        End If
    End If

    cSqlStmt = " select (sum(a.GROSS_PAY) + sum(a.SA_NET_PAY)) as GROSS_PAY,(sum(a.NET_PAY) + sum(a.SA_NET_PAY)) as NET_PAY from pa87260 a " & _
               " left join pa37722 b on a.costcenterid=b.costcenterid " & _
               " where periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam1 <> "", " and (" & cParam1 & ")", "") & _
               " group by a.periodid"

'    MsgBox cSqlStmt
    OpenQueryDNS cSqlStmt, oTempADO, False
    nSubGross = IIf(oTempADO.RecordCount > 0, IIf(gAgency = 0, oTempADO("GROSS_PAY"), oTempADO("NET_PAY")), 0)


    cSqlStmt = " select sum(a.DED_AMT) as DED_AMT from pa87263 a " & _
               " left join pa87260 b on a.periodid=b.periodid and a.empid=b.empid " & _
               " left join pa3330 c on a.dedid=c.dedid " & _
               " where a.periodid = " & cQuote & cPeriodID & cQuote & IIf(cParam2 <> "", " and (" & cParam2 & ")", "") & _
               " group by a.periodid"

    OpenQueryDNS cSqlStmt, oTempADO, False
    nSubded = IIf(oTempADO.RecordCount > 0, oTempADO("DED_AMT"), 0)

    If gAgency = 0 Then
        nGtotal = nSubGross - nSubded
    Else
        nGtotal = nSubGross
    End If
    

    cSqlStmt = "SELECT PERIODID, DATE_START, DATE_END, DURATION, STATUS FROM pa7730 where periodid=" & cQuote & cPeriodID & cQuote
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        If objdbRs("status") = 1 Then
            cPeriodName = Format(objdbRs("date_start"), "YYYY.MM") & " 1st Salary Doc"
            cSalaryName = Format(objdbRs("date_start"), "YYYY.MM") & " Other AP - Salary " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
        Else
            cPeriodName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " 2nd Salary Doc"
            cSalaryName = Year(objdbRs("date_start")) & "." & Month(objdbRs("date_start")) + 1 & " Other AP - Salary " & IIf(gAgency = 1, " (" & cODBC & ") ", "")
        End If
    End If

    cSqlStmt = "insert into tmpERPSAL(BUKRS,GJAHR,BUDAT,BLDAT,WAERS,FLAG_A,FLAG_B,BKTXT,WRBTR,SGTXT,KOSTL,SEQ_NO)values(" & _
                cQuote & nCompCode & cQuote & "," & _
                cQuote & Format(Now, "YYYY") & cQuote & "," & _
                cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                cQuote & "PHP" & cQuote & "," & _
                cQuote & IIf(gAgency = 0, "R", "A") & cQuote & "," & _
                cQuote & " " & cQuote & "," & _
                cQuote & cPeriodName & cQuote & "," & _
                IIf(gAgency = 0, -nGtotal, nGtotal) & "," & _
                cQuote & cSalaryName & cQuote & "," & _
                cQuote & "999" & cQuote & ",8)"

    QueryTemp cSqlStmt, objdbRs, True



    ShowProgress 3

    GenerateReport "Salary Payment (Summary)", "rptFinance.RPT", , True

    ShowProgress 4
    
    
    GoTo ERPEnd:
    
    
FINANCEERP:

    ShowProgress 4
    
    MsgBox "No report to generate!", vbInformation, "System Advisory!!!"
    
    Set oTempADO = Nothing

ERPEnd:
    Set oTempADO = Nothing
End Sub

Sub GenEmpMasterDataDBF()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cFieldName As String, _
        nCtr As Integer
        
    cSqlStmt = "CREATE TABLE K1Master (" & _
               " [EMPID] char(6),                     [TCID] char(5), " & _
               " [BCID] char(2),                      [BACCNTNO] char(16),                [RATE_AMT] decimal(18,4),            [DATEREG] date, " & _
               " [DATE_HIRE] date,                    [DATE_FIN] date,                    [DATE_RES] date,                     [FIRSTNAME] char(50), " & _
               " [MNAME] char(50),                    [LASTNAME] char(50),                [FULLNAME] char(50),                 [BIRTHDAY] date, " & _
               " [ADD_NO] char(230),                  [ADD_BRGY] char(230),               [ADD_CITY] char(230),                [TEL_NUM] char(15), " & _
               " [POS_ALLOW] decimal(18,4),           [PHEALTHNUM] char(15),              [PAGIBIGNO] char(15),                [SSNUM] char(15), " & _
               " [TIN] char(15),                      [SEX] char(50),                     " & _
               " [EMP_STAT] char(50),                 [ACTIVE] char(50),                  [PAYSTATUS] char(50),                [CCID] char(16), " & _
               " [CCIDDESC] char(100),                [WCID] char(16),                    [WCIDDESC] char(100),                [REMARK] char(100), " & _
               " [S_REMARK] char(100),                [LINENAME] char(100),               [POSNAME] char(50),                  [TAXCODE] char(5), " & _
               " [ISUNION] char(50),                   [SL_AVAIL] integer,                 [VL_AVAIL] integer,                  [UL_AVAIL] integer, " & _
               " [SL_USE] integer,                    [VL_USE] integer,                   [UL_USE] integer  ) "
               
               
        
    
    oDBFConn.Execute cSqlStmt
    While oDBFConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...

    cSqlStmt = "DELETE FROM K1Master"
    QueryDBF cSqlStmt, oTempADO, True


End Sub

Sub GenEmpMasterData(ByVal cPeriodID As String, ByVal cParam As String, nFilter As Integer)
    Dim oRecordSet As New ADODB.Recordset, _
        cSqlStmt As String, _
        cWant As String, _
        cDurationPath As String, _
        FileSys As FileSystemObject, _
        nCtr As Integer, _
        cParam2 As String

    Dim ntag As Integer, _
        aOtherInfo As Variant
        
    aOtherInfo = Array("", "", "", "", "", "", "", "", "", "")
    '0 Liname
    '1 posname
    '2 taxcode
    '3 coscenter name
    '4 workcenter name
    '5 SEX
    '6 EMPSTAT
    '7 ACTIVE
    '8 PAYSTATUS
    '9 Union member
        
    Set FileSys = New FileSystemObject
    
    cDurationPath = CheckPath(Text4.Text)
    
    If FileSys.FileExists(cDurationPath & "K1Master.dbf") = True Then
        cWant = MsgBox("K1Master.dbf Already Exsist... do you want to Overwrite the file ?", vbYesNo + vbCritical, App.Title)
        If cWant = vbYes Then
            FileSys.DeleteFile cDurationPath & "K1Master.dbf"
        End If
    Else
        cWant = vbYes
    End If
    
    If cWant = vbYes Then
        DetectDBF cDurationPath
        
        ShowProgress 0
        
        GenEmpMasterDataDBF
        
        If Check5.Value <> vbChecked Then
        
            If Trim(cParam) <> "" Then
                cParam = " and b.depid IN " & cParam
            End If
            
            'check is pclose is =1
            cSqlStmt = "select PERIODID, DATE_START, DATE_END, DURATION, PCLOSE, DATE_CLOSE  from pa7730 where periodid= " & cQuote & cPeriodID & cQuote
            OpenQueryDNS cSqlStmt, oTempADO, False
            ntag = IIf(oTempADO.RecordCount > 0, oTempADO("PCLOSE"), "")
            
            Select Case Combo1.ListIndex
                Case 0
                    cParam2 = " and b.active = 0 "
                Case 1
                    cParam2 = " and b.active = 1 and b.date_res between " & cQuote & Format(oTempADO("date_start"), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(oTempADO("date_end"), "yyyy-mm-dd") & cQuote
                Case 2
                    cParam2 = " and b.active = 2 and b.date_fin between " & cQuote & Format(oTempADO("date_start"), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(oTempADO("date_end"), "yyyy-mm-dd") & cQuote
                Case 3
                    cParam2 = " and b.active = 3 and b.date_res between " & cQuote & Format(oTempADO("date_start"), "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(oTempADO("date_end"), "yyyy-mm-dd") & cQuote
            End Select
            
            cSqlStmt = " select a.PERIODID, a.date, a.EMPID, b.TCID, b.BCID, b.BACCNTNO, b.RATE_AMT, b.DATEREG, b.DATE_HIRE, b.DATE_FIN, b.DATE_RES, " & _
                       " b.FIRSTNAME, b.MNAME, b.LASTNAME, concat(b.LASTNAME,', ',b.FIRSTNAME, ' ', left(b.MNAME,1),'. ') as fullname, " & _
                       " b.BIRTHDAY,b.ADD_NO,b.ADD_BRGY,b.ADD_CITY,b.TEL_NUM,b.POS_ALLOW,b.PHEALTHNUM,b.PAGIBIGNO,b.SSNUM,b.TIN,b.REMARK,b.S_REMARK, " & _
                       " b.SEX, b.EMP_STAT, b.Active, b.PAYSTATUS, " & _
                       " b.depid, b.posid, b.COSTCENTERID, b.WORKCENTERID, b.taxid, " & _
                       " b.ISUNION, b.SL_AVAIL, b.VL_AVAIL, b.UL_AVAIL, b.SL_USE, b.VL_USE, b.UL_USE " & _
                       " from " & IIf(ntag = 1, "dih36770", "di36770") & " a  left join di3670 b on a.empid=b.empid " & _
                       " Where a.periodid = " & cQuote & cPeriodID & cQuote & _
                       cParam & cParam2 & _
                       " group by a.empid "
        Else
            
            cSqlStmt = " select EMPID, TCID, BCID, BACCNTNO, RATE_AMT, DATEREG, DATE_HIRE, DATE_FIN, DATE_RES," & _
                       " FIRSTNAME, MNAME, LASTNAME, concat(LASTNAME,', ',FIRSTNAME, ' ', left(MNAME,1),'. ') as fullname, " & _
                       " BIRTHDAY, ADD_NO, ADD_BRGY, ADD_CITY, TEL_NUM, POS_ALLOW, PHEALTHNUM, PAGIBIGNO, SSNUM, TIN, REMARK, S_REMARK, " & _
                       " SEX, EMP_STAT, Active, PAYSTATUS, " & _
                       " depid, posid, COSTCENTERID, WORKCENTERID, taxid, " & _
                       " ISUNION , SL_AVAIL, VL_AVAIL, UL_AVAIL, SL_USE, VL_USE, UL_USE " & _
                       " From di3670 "
            
        End If
    '    Script2File cSqlStmt
    '    MsgBox cSqlStmt
        OpenQueryDNS cSqlStmt, oRecordSet, False
        If oRecordSet.RecordCount > 0 Then
            While Not oRecordSet.EOF
                ShowProgress 2, (oRecordSet.AbsolutePosition / oRecordSet.RecordCount) * 100, , , "Copying " & Trim(oRecordSet("EMPID")) & "... " & Trim(oRecordSet("Lastname") & ", " & oRecordSet("firstname")) & "..."
                
                'Linaname
                OpenQueryDNS " select lineid, linename from di5463 where lineid=" & cQuote & oRecordSet("depid") & cQuote, objdbRs, False
                aOtherInfo(0) = IIf(objdbRs.RecordCount > 0, objdbRs("linename"), "")
                
                'Posname
                OpenQueryDNS " select posid,posname from di7670 where posid = " & cQuote & oRecordSet("posid") & cQuote, objdbRs, False
                aOtherInfo(1) = IIf(objdbRs.RecordCount > 0, objdbRs("posname"), "")
            
                'Posname
                OpenQueryDNS " select TAXID, TAXCODE  from PA8290 where taxid = " & cQuote & oRecordSet("taxid") & cQuote, objdbRs, False
                aOtherInfo(2) = IIf(objdbRs.RecordCount > 0, objdbRs("TAXCODE"), "")
            
                'costcentername
                OpenQueryDNS " select COSTCENTERID, DESCRIPTION from pa37722 where COSTCENTERID = " & cQuote & oRecordSet("COSTCENTERID") & cQuote, objdbRs, False
                aOtherInfo(3) = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
                'workcentername
                OpenQueryDNS " select WORKCENTERID, DESCRIPTION  from pa97722 where WORKCENTERID = " & cQuote & oRecordSet("WORKCENTERID") & cQuote, objdbRs, False
                aOtherInfo(4) = IIf(objdbRs.RecordCount > 0, objdbRs("DESCRIPTION"), "")
    
                aOtherInfo(5) = IIf(oRecordSet("SEX") = 0, "Male", "Female")
                
                aOtherInfo(6) = IIf(oRecordSet("EMP_STAT") = 0, "Wap", IIf(oRecordSet("EMP_STAT") = 1, "Contractual", "Regular"))
                
                aOtherInfo(7) = IIf(oRecordSet("ACTIVE") = 0, "Active", IIf(oRecordSet("ACTIVE") = 1, "Resigned", IIf(oRecordSet("ACTIVE") = 2, "Finished", "Terminated")))
           
                aOtherInfo(8) = IIf(oRecordSet("PAYSTATUS") = 0, "Daily", IIf(oRecordSet("PAYSTATUS") = 1, "Monthly", "Emergency"))
                
                aOtherInfo(9) = IIf(oRecordSet("ISUNION") = 0, "Not Union Member", "Union Member")
                
                cSqlStmt = " INSERT INTO K1Master.dbf ( EMPID, TCID, BCID, BACCNTNO, RATE_AMT, DATEREG, DATE_HIRE, DATE_FIN, DATE_RES, FIRSTNAME, MNAME, LASTNAME, FULLNAME, " & _
                           " BIRTHDAY, ADD_NO, ADD_BRGY, ADD_CITY, TEL_NUM, POS_ALLOW, PHEALTHNUM, PAGIBIGNO, SSNUM, TIN, SEX, EMP_STAT, ACTIVE, PAYSTATUS, CCID, " & _
                           " CCIDDESC, WCID, WCIDDESC, REMARK, S_REMARK, LINENAME, POSNAME, TAXCODE, ISUNION, SL_AVAIL, VL_AVAIL, UL_AVAIL, SL_USE, VL_USE, UL_USE)VALUES(" & _
                           cQuote & oRecordSet("EMPID") & cQuote & "," & cQuote & oRecordSet("TCID") & cQuote & "," & cQuote & oRecordSet("BCID") & cQuote & "," & cQuote & oRecordSet("BACCNTNO") & cQuote & "," & cQuote & oRecordSet("RATE_AMT") & cQuote & "," & _
                           cQuote & Format(oRecordSet("DATEREG"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("DATE_HIRE"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("DATE_FIN"), "yyyy-mm-dd") & cQuote & "," & cQuote & Format(oRecordSet("DATE_RES"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & oRecordSet("FIRSTNAME") & cQuote & "," & cQuote & oRecordSet("MNAME") & cQuote & "," & cQuote & oRecordSet("LASTNAME") & cQuote & "," & cQuote & oRecordSet("FullName") & cQuote & "," & cQuote & Format(oRecordSet("BIRTHDAY"), "YYYY-MM-DD") & cQuote & "," & _
                           cQuote & oRecordSet("ADD_NO") & cQuote & "," & cQuote & oRecordSet("ADD_BRGY") & cQuote & "," & cQuote & oRecordSet("ADD_CITY") & cQuote & "," & cQuote & oRecordSet("TEL_NUM") & cQuote & "," & cQuote & oRecordSet("POS_ALLOW") & cQuote & "," & _
                           cQuote & oRecordSet("PHEALTHNUM") & cQuote & "," & cQuote & oRecordSet("PAGIBIGNO") & cQuote & "," & cQuote & oRecordSet("SSNUM") & cQuote & "," & cQuote & oRecordSet("TIN") & cQuote & "," & _
                           cQuote & aOtherInfo(5) & cQuote & "," & cQuote & aOtherInfo(6) & cQuote & "," & cQuote & aOtherInfo(7) & cQuote & "," & cQuote & aOtherInfo(8) & cQuote & "," & _
                           cQuote & oRecordSet("CostCenterid") & cQuote & "," & cQuote & aOtherInfo(3) & cQuote & "," & cQuote & oRecordSet("workCenterid") & cQuote & "," & cQuote & aOtherInfo(4) & cQuote & "," & _
                           cQuote & oRecordSet("REMARK") & cQuote & "," & cQuote & oRecordSet("S_REMARK") & cQuote & "," & cQuote & aOtherInfo(0) & cQuote & "," & _
                           cQuote & aOtherInfo(1) & cQuote & "," & cQuote & aOtherInfo(2) & cQuote & "," & cQuote & aOtherInfo(9) & cQuote & "," & _
                           oRecordSet("SL_AVAIL") & "," & oRecordSet("VL_AVAIL") & "," & oRecordSet("UL_AVAIL") & "," & oRecordSet("SL_USE") & "," & oRecordSet("VL_USE") & "," & oRecordSet("UL_USE") & ")"
                           
    '            MsgBox cSqlStmt
                QueryDBF cSqlStmt, objdbRs, True
    
                oRecordSet.MoveNext
                
                
            Wend
            ShowProgress 4
            MsgBox "Data Generation Done...", vbInformation, App.Title
        Else
            ShowProgress 4
            MsgBox "Data not found...!!!", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
End Sub


Private Sub Check1_Click()
    Dim lProceed As Boolean
'    Dim nCtr As Integer
'
'    ListView1.Enabled = Check1.Value <> 1
'    For nCtr = 1 To ListView1.ListItems.Count
'        ListView1.ListItems(nCtr).Checked = Check1.Value = vbChecked
'    Next nCtr
    
'    addnew 2015-05-07
    Select Case Tag
'        Case 3, 4, 5, 6, 10, 15, 22, 34, 36, 44, 45
        Case 6, 15
            If Check1.Value = vbChecked Then
                ' --> added security as of 2015-05-07
                If (gUserGroup = 0) And Not lSuperUser Then
                        frmManager.Tag = 1
                        frmManager.Show 1
                        If ModalResult = mrCancel Then Exit Sub
                        lProceed = ModalResult = mrOk
                Else
                    lProceed = True
                End If
                
                If Not lProceed Then
                    MsgBox "You are not allowed to listup close payroll period", vbExclamation, "System Advisory!!!"
                    Exit Sub
                Else
                    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID DESC", objdbRs, False
                End If
                ' --> end of added security...
            
            
'                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID DESC", objdbRs, False
            Else
                OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 WHERE (13month=0) and (PCLOSE=0) and (DATE_START < CURDATE() and DATE_END < CURDATE()) ORDER BY PERIODID", objdbRs, False
            End If
        Case Else
            OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 " & IIf(Check1.Value = vbChecked, " WHERE PCLOSE=0 AND 13MONTH<>1  ", "WHERE 13MONTH<>1 ") & " ORDER BY PERIODID DESC", objdbRs, False
    End Select
    
    add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
            
'    OpenQueryDNS "SELECT PERIODID, DURATION FROM PA7730 " & IIf(Check1.Value = vbChecked, " WHERE PCLOSE=0 AND 13MONTH<>1  ", "WHERE 13MONTH<>1 ") & " ORDER BY PERIODID DESC", objdbRs, False
'    add2LstBox objdbRs, ListView1, Array("DURATION", "PERIODID")
    
End Sub

Private Sub Check2_Click()
    Dim nCtr As Integer
    
    ListView2.Enabled = Check2.Value <> 1
    For nCtr = 1 To ListView2.ListItems.Count
        ListView2.ListItems(nCtr).Checked = Check2.Value = vbChecked
    Next nCtr
End Sub

Private Sub Check3_Click()
'    If (Tag <> 9) And (Tag <> 10) Then
'        Combo1.ListIndex = 0
'        Combo1.Visible = Check3.Value = vbUnchecked
'        Label2.Visible = Check3.Value = vbUnchecked
'    End If
End Sub

Private Sub Check4_Click()
    Select Case Tag
        Case 6, 7, 15, 20, 28, 29, 33
            SSTab1.TabVisible(3) = Check4.Value <> vbChecked
        Case 14
            Label10.Caption = IIf(Check4.Value = vbChecked, "Checked By", "Prepared By")
            Text5.Visible = Check4.Value <> vbChecked
            Label5.Visible = Check4.Value <> vbChecked
            Label6.Visible = Check4.Value <> vbChecked
            Command14.Visible = Check4.Value <> vbChecked
            If Check4.Value = vbChecked Then
                Check7.Visible = False
            Else
                Check7.Visible = True
            End If
                        
    End Select
End Sub


Private Sub Check5_Click()
    If Tag = 51 Then
        If Check5.Value <> vbChecked Then
            With SSTab1
                .TabVisible(0) = True
                .TabVisible(1) = True
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = True
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Department"
                .TabCaption(4) = "Generate File"
            End With
        Else
            With SSTab1
                .TabVisible(0) = False
                .TabVisible(1) = False
                .TabVisible(2) = False
                .TabVisible(3) = False
                .TabVisible(4) = True
                .TabCaption(0) = "Select Period"
                .TabCaption(1) = "Department"
                .TabCaption(4) = "Generate File"
            End With
        End If
    End If
    Combo1.Enabled = Check5.Value <> vbChecked
End Sub

Private Sub Combo1_Click()
'If gCompanyID = "0002" Then
'    Check3.Enabled = IIf(Tag <> 9, Combo1.ListIndex < 1, True)
'End If
'    If Combo1.ListIndex <> 4 Then
'        Check6.Enabled = True
'    Else
'        Check6.Enabled = False
'    End If
    If Tag >= 47 And Tag <= 49 Then
    
        SSTab1.TabVisible(1) = True
        Check7.Visible = True
        
        If Tag = 47 Then
            If Combo1.ListIndex <> 1 Then
               SSTab1.TabVisible(1) = False
               Check7.Visible = False
            End If
        Else
             If Combo1.ListIndex <> 1 Then Check7.Visible = False
        End If
    End If
End Sub

Private Sub Command1_Click()
    cmdClick Text2, Label9
    Text2.SetFocus
End Sub

Private Sub Command13_Click()
    cmdClick Text6, Label8
    Text6.SetFocus
End Sub

Private Sub Command14_Click()
    cmdClick Text5, Label6
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

Private Sub Command2_Click()
    cmdClick Text3, Label14
    Text3.SetFocus
End Sub

Private Sub Command5_Click()
    cmdClick Text1, Label4
    Text1.SetFocus
End Sub

Private Sub Dir1_Change()
    Text4.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 3, Text1.Text, Label4
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 6, Text2.Text, Label9
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 7, Text3.Text, Label14
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 2, Text5.Text, Label6
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 1, Text6.Text, Label8
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 4, Text7.Text, Label15
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtKeyDown 5, Text8.Text, Label16
End Sub

Private Sub XPButton3_Click()
    Dim nCtr As Integer, _
        cParam As String, _
        cParam2 As String, _
        cPeriod As String
        
    If (Tag <> 12) And ((Tag < 19) Or (Tag > 20)) And ((Tag < 32) Or (Tag > 33)) Then
        If (Tag <> 31) And (Tag <> 21) Then
            If (ListView1.Visible = True) Then cPeriod = ListView1.SelectedItem.Text
        End If
    End If
    
    If (Tag = 17) Or (Tag = 37) Or (Tag = 38) Then
        If (Text4.Text = "") Then 'Or (Text9.Text = "") Then
            MsgBox "Select Path/Location...", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    If Tag <= 2 Then
        If Check1.Value = vbUnchecked Then
            For nCtr = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(nCtr).Checked Then cParam = cParam & cQuote & ListView1.ListItems(nCtr).Text & cQuote & ","
            Next nCtr
            
            If Trim(cParam) <> "" Then cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
        End If
    Else
        If Tag <> 9 Then
            If Check2.Value = vbUnchecked Then
                For nCtr = 1 To ListView2.ListItems.Count
                    If ListView2.ListItems(nCtr).Checked Then cParam = cParam & cQuote & ListView2.ListItems(nCtr).Text & cQuote & ","
                Next nCtr
                
                If Trim(cParam) <> "" Then cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
            End If
            If Tag = 10 Then
                If Check2.Value = vbUnchecked Then
                    For nCtr = 1 To ListView2.ListItems.Count
                        If ListView2.ListItems(nCtr).Checked Then cParam2 = cParam2 & cQuote & ListView2.ListItems(nCtr).Text & cQuote & ","
                    Next nCtr
                    
                    If Trim(cParam2) <> "" Then
                        cParam2 = " and (b.depid IN (" & left(cParam2, Len(cParam2) - 1) & "))"
                    Else
                        MsgBox "Please specify a Department/Line to preview!", vbInformation, "Monthly Production Report - " & App.Title
                        Exit Sub
                    End If
                End If
            End If
            If Tag = 12 Then
                If Check1.Value = vbUnchecked Then
                    For nCtr = 1 To ListView1.ListItems.Count
                        If ListView1.ListItems(nCtr).Checked Then cParam = cParam & cQuote & ListView1.ListItems(nCtr).Text & cQuote & ","
                    Next nCtr
                    
                    If Trim(cParam) <> "" Then cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
                End If
            End If
        Else
            If Check1.Value = vbUnchecked Then
                For nCtr = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(nCtr).Checked Then cParam = cParam & cQuote & ListView1.ListItems(nCtr).Text & cQuote & ","
                Next nCtr
                
                If Trim(cParam) <> "" Then
                    cParam = " ((month(b.date_fin) IN (" & left(cParam, Len(cParam) - 1) & ")) and (year(b.date_fin) = " & Combo1.Text & "))"
                Else
                    MsgBox "Please specify a month to preview!", vbInformation, "Monthly Production Report - " & App.Title
                    Exit Sub
                End If
            End If
      
            If Check2.Value = vbUnchecked Then
                For nCtr = 1 To ListView2.ListItems.Count
                    If ListView2.ListItems(nCtr).Checked Then cParam2 = cParam2 & cQuote & ListView2.ListItems(nCtr).Text & cQuote & ","
                Next nCtr
                
                If Trim(cParam2) <> "" Then
                    cParam2 = " and (b.depid IN (" & left(cParam2, Len(cParam2) - 1) & "))"
                Else
                    MsgBox "Please specify a Department/Line to preview!", vbInformation, "Monthly Production Report - " & App.Title
                    Exit Sub
                End If
            End If
        End If
    End If
        
    Select Case Tag
        Case 1      ' --> Process Payroll Transaction
'            ProcessTransaction
            
            GetUserRights PadStr(frmMain.mnuTransaction.Name, " ", 100, PadRight), gUserID
            With frmTransaction
                .lblPeriod.Caption = cPeriod
                .lblDuration.Caption = ListView1.SelectedItem.SubItems(1) & " Payroll"
                .Show
            End With
            
        Case 2      ' --> Generate PEZA Report
            GenPEZA cPeriod
            
        Case 3      ' --> Employee Listing
            If Check4.Value = vbChecked Then
                GenEmpList cPeriod, cParam, Combo1.ListIndex
            Else
                GenEmpListSum cPeriod, cParam, Combo1.ListIndex
            End If
            
        Case 4, 5, 6, 15, 44, 45
            GenPayRoll cPeriod, cParam, Combo1.ListIndex
            
        Case 7, 23, 41   ' --> Denomination Report
            GenDenom cPeriod, 0, Tag
            
        Case 8      ' --> Close Period
            ClosePeriod cPeriod
            
        Case 9      ' --> Monthly Finish Contract Report
            MFCon cParam, cParam2, 0
            
        Case 10
            MFCon cPeriod, cParam2, 1
        
        Case 11     ' --> Leave Report by Period
            Genleaverep cPeriod, Combo1.ListIndex
        
        Case 12   ' --> new medicare 20080710
            GenPhilHealth cParam, Combo1.ListIndex
        
        Case 13     ' --> Withholding Tax Report
            GenWithRep cPeriod
            
        Case 14     ' --> Pag-Ibig Remittance Report
            GenPagIbig cPeriod
            
        Case 16     ' --> SSS Premium Remittance Report
            SSSPData cPeriod
            
        Case 17     ' --> Generate Backup
            GenBackupPay cPeriod
'            GenBackup cPeriod
            
        Case 18, 46    ' --> Salary Division
            GenSalDiv cPeriod, Tag
            
        Case 19, 20 ' --> generate payslip/paysheet for 13th month
            GenPayRoll Text2.Text, cParam, 0
            
        Case 21     ' --> backup 13th month
            Gen13mobackup
            
        Case 22     ' --> 13th month acknowledgement...
            GenPayRoll cPeriod, cParam, 0
            
        Case 24     ' --> Alphalist (Annual Withholding Tax)
            GenAlphaList cPeriod
            
        Case 25     ' --> Alphalist for Contractual (Annual Withholding Tax)
'            GenAlphaList2 cperiod
            
        Case 26     ' --> SS Employment Report (R-1A)
            GenSSSR1A
            
        Case 27, 28, 29, 30   ' --> Employee Incenive payroll
            GenIncPayroll cPeriod, cParam, Combo1.ListIndex
            
        Case 31   ' --> Loan Report
            GenLoanRpt cParam
            
        Case 32, 33  ' --> payslip/payshet for SLVLPayRoll 20081206
            GenSLVLPayRoll Text2.Text, cParam, 0
            
        Case 34     ' --> SLVL acknowledgement...
            GenSLVLPayRoll cPeriod, cParam, 0
            
        Case 35     ' --> SLVL Denomination Report
            GenDenom cPeriod, 2, Tag
        Case 36      ' --> Manpower Listing
            GenManPowerList cPeriod, cParam, 0
            
'        Case 37      ' --> RCBC ATM EXCELL FILE
'            GenManATMExcell cPeriod, cParam
'
'        Case 38      ' --> RCBC ATM EXCELL FILE
'            GenManATMIX1 cPeriod, cParam
            
        Case 39      ' --> RCBC ATM Transmittal report
            GenManATMTrans cPeriod, cParam
            
         Case 40   ' --> Payroll Analysis
            GenGrandOT cPeriod
            
        Case 42   ' --> PHILHEALTH REPORT Er2
            GenPHILh
        Case 43     ' --> SLVL acknowledgement...
            GenSLVLPayRoll cPeriod, cParam, 0
        Case 47 '--> TMS Report
            GenTMSRpt cPeriod, cParam
        Case 48, 49 '--> TMS Summary Report
            GenTMSSumRpt cPeriod, cParam
        Case 50
            GenERPSAL cPeriod, cParam, Combo1.ListIndex
        Case 51 '--> Employee Master Data Generation
            GenEmpMasterData cPeriod, cParam, Combo1.ListIndex
            
            
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Me.Name, "OPEN"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log2Audit Me.Name, "CLOSE"
End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
    Set oDBFConn = Nothing
    Set oSSSConn = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ListView1_DblClick()
    XPButton3_Click
End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub

Private Sub XPButton2_Click()
    Dim cString As String
    
    OpenQueryDNS "SELECT * FROM PA87260 WHERE PERIODID=" & cQuote & ListView1.SelectedItem.Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        cString = "A transaction had been detected for Period ID #" & ListView1.SelectedItem.Text & Chr$(13) & Chr$(10) & _
                  "Would you like to re-process the transaction?"
        If MsgBox(cString, vbYesNo, "System Advisory") = vbYes Then
            OpenQueryDNS "DELETE FROM PA87260 WHERE PERIODID=" & cQuote & ListView1.SelectedItem.Text & cQuote, objdbRs, True
            Script2File "DELETE FROM PA87260 WHERE PERIODID=" & cQuote & ListView1.SelectedItem.Text & cQuote
            
            OpenQueryDNS "DELETE FROM PA87263 WHERE PERIODID=" & cQuote & ListView1.SelectedItem.Text & cQuote, objdbRs, True
            Script2File "DELETE FROM PA87263 WHERE PERIODID=" & cQuote & ListView1.SelectedItem.Text & cQuote
            
            Log2Audit Me.Name, "Re-Process Transaction for Period ID#" & ListView1.SelectedItem.Text
            
            ProcessTransaction
        End If
    Else
        ProcessTransaction
    End If
End Sub
