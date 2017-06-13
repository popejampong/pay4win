VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaXPFrame.ocx"
Begin VB.Form frmTMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Management System"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14685
   Begin ciaXPPanel.XPPanel XPPanel3 
      Height          =   1875
      Left            =   6390
      TabIndex        =   66
      Top             =   3675
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   3307
      LicValid        =   -1  'True
      Begin VB.CommandButton Command12 
         Caption         =   "Close"
         Height          =   435
         Index           =   3
         Left            =   90
         TabIndex        =   70
         Top             =   1350
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Extension Report"
         Height          =   435
         Index           =   2
         Left            =   90
         TabIndex        =   69
         Top             =   930
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Regular Report"
         Height          =   435
         Index           =   1
         Left            =   90
         TabIndex        =   68
         Top             =   510
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Complete Report"
         Height          =   435
         Index           =   0
         Left            =   90
         TabIndex        =   67
         Top             =   90
         Width           =   1800
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel2 
      Height          =   5715
      Left            =   4830
      TabIndex        =   60
      Top             =   1725
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   10081
      HasTopBorder    =   0   'False
      HasBottomBorder =   0   'False
      LicValid        =   -1  'True
      Begin VB.CheckBox Check1 
         Caption         =   "Select &All"
         Height          =   255
         Left            =   105
         TabIndex        =   63
         Top             =   5445
         Width           =   1290
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4515
         Picture         =   "frmTMS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "15"
         Top             =   285
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Close"
         Height          =   660
         Left            =   4515
         Picture         =   "frmTMS.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   61
         Tag             =   "21"
         Top             =   1065
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5190
         Left            =   90
         TabIndex        =   64
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   9155
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
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         BorderStyle     =   0  'Transparent
         Height          =   5715
         Left            =   4290
         Top             =   15
         Width           =   1320
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   65
         Top             =   30
         Width           =   3600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   12450
      TabIndex        =   56
      Top             =   6855
      Visible         =   0   'False
      Width           =   1230
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   450
         Index           =   2
         Left            =   75
         TabIndex        =   59
         Top             =   1035
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Detailed"
         Height          =   450
         Index           =   1
         Left            =   75
         TabIndex        =   58
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Summary"
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   57
         Top             =   165
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1995
      Left            =   10890
      TabIndex        =   51
      Top             =   6420
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton Command7 
         Caption         =   "Ca&ncel"
         Height          =   450
         Index           =   3
         Left            =   75
         TabIndex        =   55
         Top             =   1470
         Width           =   1410
      End
      Begin VB.CommandButton Command7 
         Caption         =   "by E&mployee"
         Height          =   450
         Index           =   2
         Left            =   75
         TabIndex        =   54
         Top             =   1035
         Width           =   1410
      End
      Begin VB.CommandButton Command7 
         Caption         =   "by &Department"
         Height          =   450
         Index           =   1
         Left            =   75
         TabIndex        =   53
         Top             =   600
         Width           =   1410
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&TMS Report"
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   52
         Top             =   165
         Width           =   1410
      End
   End
   Begin MSComCtl2.DTPicker dtFlex 
      Height          =   375
      Left            =   5430
      TabIndex        =   50
      Top             =   7215
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   63766528
      CurrentDate     =   38381
   End
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmTMS.frx":3304
      Left            =   5400
      List            =   "frmTMS.frx":330E
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   7710
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Save"
      Height          =   660
      Left            =   12720
      Picture         =   "frmTMS.frx":331B
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "20"
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   660
      Left            =   13695
      Picture         =   "frmTMS.frx":4C9D
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "21"
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Edit"
      Height          =   660
      Left            =   11865
      Picture         =   "frmTMS.frx":661F
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "16"
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Preview"
      Height          =   660
      Left            =   10890
      Picture         =   "frmTMS.frx":7FA1
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "15"
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Generate"
      Height          =   660
      Left            =   9915
      Picture         =   "frmTMS.frx":9923
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "22"
      Top             =   8400
      Width           =   855
   End
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   1230
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   2170
      Alignment       =   2
      Caption         =   " Select a Period "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.CheckBox Check2 
         Caption         =   "Exclude Close Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   3795
         TabIndex        =   1
         Top             =   510
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   1335
         TabIndex        =   12
         Top             =   180
         Width           =   390
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
         Left            =   615
         TabIndex        =   0
         Top             =   180
         Width           =   690
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   285
         Left            =   600
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   495
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   63766528
         CurrentDate     =   38623
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   600
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   795
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   63766528
         CurrentDate     =   38623
      End
      Begin VB.Label lblPClose 
         BackStyle       =   0  'Transparent
         Caption         =   "Period Close"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3660
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "End"
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
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
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
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   525
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1785
         TabIndex        =   15
         Top             =   240
         Width           =   3045
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7035
      Left            =   75
      TabIndex        =   2
      Top             =   1260
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   12409
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
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
      TabIndex        =   20
      Top             =   8325
      Width           =   4425
      _ExtentX        =   7805
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
         ItemData        =   "frmTMS.frx":B2A5
         Left            =   645
         List            =   "frmTMS.frx":B2BB
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   90
         Width           =   2205
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
         Left            =   645
         TabIndex        =   21
         Top             =   420
         Width           =   3705
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
         Height          =   285
         Left            =   90
         TabIndex        =   24
         Top             =   150
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
         TabIndex        =   23
         Top             =   465
         Width           =   1575
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame2 
      Height          =   1230
      Left            =   5100
      TabIndex        =   25
      Top             =   30
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   2170
      Alignment       =   2
      Caption         =   " Employee Information "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1080
         TabIndex        =   31
         Top             =   510
         Width           =   4005
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1080
         TabIndex        =   30
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   210
         Width           =   990
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   1080
         TabIndex        =   29
         Top             =   810
         Width           =   4005
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   4095
         TabIndex        =   28
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   210
         Width           =   990
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   5925
         TabIndex        =   27
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   210
         Width           =   3450
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   5925
         TabIndex        =   26
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   510
         Width           =   3450
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fullname"
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   570
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Enroll Number"
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
         Height          =   255
         Left            =   3030
         TabIndex        =   35
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Height          =   255
         Left            =   5235
         TabIndex        =   34
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   255
         Left            =   5235
         TabIndex        =   33
         Top             =   570
         Width           =   990
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Resigned/FC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5940
         TabIndex        =   32
         Top             =   840
         Width           =   1830
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5670
      Left            =   5085
      TabIndex        =   3
      Top             =   1260
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   10001
      _Version        =   393216
      GridColor       =   12640511
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   1335
      Left            =   5085
      TabIndex        =   4
      Top             =   6960
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   2355
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
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   1350
      Left            =   9750
      TabIndex        =   39
      Top             =   6945
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   2381
      LicValid        =   -1  'True
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   315
         Left            =   1695
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   450
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
         Left            =   1035
         TabIndex        =   5
         Top             =   135
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift ID"
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
         Left            =   150
         TabIndex        =   47
         Top             =   195
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Height          =   255
         Left            =   150
         TabIndex        =   46
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
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
         Height          =   255
         Left            =   150
         TabIndex        =   45
         Top             =   765
         Width           =   1110
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "End Time"
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
         Height          =   255
         Left            =   150
         TabIndex        =   44
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   1050
         TabIndex        =   43
         Top             =   480
         Width           =   3450
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
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
         Left            =   1050
         TabIndex        =   42
         Top             =   765
         Width           =   2865
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
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
         Left            =   1050
         TabIndex        =   41
         Top             =   1035
         Width           =   2865
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   810
      Left            =   4545
      TabIndex        =   48
      Top             =   8325
      Width           =   5160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   885
      Left            =   9750
      Top             =   8310
      Width           =   4950
   End
End
Attribute VB_Name = "frmTMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmTMS
' description   :   Revised Time Management System
' programmer    :   _-=[ srm ]=-_
' date          :   24 May 2007

Option Explicit

Dim oTempADO As New ADODB.Recordset, _
    nTagSelect, nAdd As Integer, _
    myArray, myArray2, myArray3, aPeriodInfo As Variant, _
    lShow As Boolean

Sub CheckGrid(ByVal nRow As Integer)
    Dim aTimeInfo As Variant
    
    With MSHFlexGrid1
        aTimeInfo = ComputeDays(.TextMatrix(nRow, 2), _
                                Array(DTPicker5.Value, DTPicker1.Value, 0), _
                                Array(Val(.TextMatrix(nRow, 10)), Val(.TextMatrix(nRow, 11))), _
                                Val(lblPClose.Caption) = 1)
        .TextMatrix(nRow, 12) = aTimeInfo(0)
        .TextMatrix(nRow, 13) = aTimeInfo(1)
        .TextMatrix(nRow, 14) = aTimeInfo(2)
        .TextMatrix(nRow, 15) = aTimeInfo(3)
        .TextMatrix(nRow, 16) = aTimeInfo(4)
        .TextMatrix(nRow, 17) = aTimeInfo(12)
        .TextMatrix(nRow, 18) = aTimeInfo(5)
        .TextMatrix(nRow, 19) = aTimeInfo(6)
    End With
End Sub

Sub BtnEnable(ByVal nMode As Integer)
    Dim nCtr As Integer
    
    Select Case nMode
        Case 0
            Command4.Enabled = False
            Command6.Enabled = False
            Command8.Enabled = False
            Command10.Enabled = False
        
        Case 1      ' --> enable generate button
            Command4.Enabled = False
            Command6.Enabled = True
            Command8.Enabled = False
            Command10.Enabled = False
            
            Label6.Caption = ""
            Label7.Caption = ""
            Label8.Caption = ""
            Label27.Caption = ""
            MSHFlexGrid1.Width = 14515
            XPFrame2.Visible = False
            
            For nCtr = 0 To 5
                Text3(nCtr).Text = ""
            Next nCtr
            
            SetGridColumn myArray, MSHFlexGrid1
            SetGridColumn myArray2, MSHFlexGrid2
            SetGridColumn myArray3, MSHFlexGrid3
    
        Case 2      ' --> disable generate button
            CtrlPanel Me, nAdd
            Command6.Enabled = False
            If Val(lblPClose.Caption) = 1 Then Command8.Enabled = False
    End Select
End Sub

Function ChkPeriod() As Boolean
    Dim cString
    If (Text1.Text <> aPeriodInfo(0)) Or _
       (DTPicker5.Value <> aPeriodInfo(1)) Or _
       (DTPicker1.Value <> aPeriodInfo(2)) Then
        aPeriodInfo = Array(Text1.Text, DTPicker5.Value, DTPicker1.Value)
        BtnEnable 1
        cString = "Warning!!!" & vbCrLf & _
                  "The selected period is already close. Revision are not anymore allowed. " & _
                  "You can still generate and preview the TMS for reference/archival purposes."
        Label18.Caption = IIf(Val(lblPClose.Caption) = 0, "", cString)
    End If
End Function

Private Sub cmbFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid3
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 4) = cmbFlex.ListIndex
                .TextMatrix(.Row, 5) = cmbFlex.Text
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
    Command5.Cancel = True
End Sub

Private Sub Combo1_Click()
    With MSHFlexGrid1
        .Redraw = False
'            myArray = Array("TXT:1[TCID]:6:True", _
'                            "TXT:2[Emp ID]:8:True", _
'                            "TXT:3[Fullname]:30:True", _
'                            "TXT:4[Position]:20:True", _
'                            "TXT:5[FName]:20:False", _
'                            "TXT:6[LName]:20:False", _
'                            "NUM:7[Dep ID]:3:False", _
'                            "TXT:8[Department]:20:True", _
'                            "NUM:9[Active]:1:False", _
'                            "NUM:0[emp stat]:1:False", _
'                            "NUM:1[WAP Status]:1:False", _
'                            "NUM:2[Days Work]:10:True", _
'                            "NUM:3[OT Hour]:9:True", _
'                            "NUM:4[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
'                            "NUM:5[NDiff]:9:True", _
'                            "NUM:6[NDiff OT]:9:True", _
'                            "NUM:7[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
'                            "NUM:8[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
'                            "NUM:9[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"))
        Select Case Combo1.ListIndex
            Case 0
                .Col = 2
            Case 1
                .Col = 1
            Case 2
                .Col = 3
            Case 3
                .Col = 5
            Case 4
                .Col = 6
            Case 5
                .Col = 8
        End Select
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With
End Sub

Private Sub Command1_Click()
    frmLookup.showPopup 5, IIf(Check2.Value = vbChecked, " where pclose=0", "")
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text1.Text = cResult
        OpenQueryDNS "SELECT * FROM PA7730 where periodid=" & cQuote & cResult & cQuote, objdbRs, False
        Label1.Caption = EncodeStr(objdbRs("duration"))
        lblPClose.Caption = objdbRs("pclose")
        DTPicker5.Value = objdbRs("date_start")
        DTPicker1.Value = objdbRs("date_end")
        ChkPeriod
    End If
    
    Text1.SetFocus
End Sub

Function ChkHoliday(dDate As Date) As String
    Dim cSqlStmt As String
    cSqlStmt = "select a.description from pa4329 a" & _
               " where (a.date=" & cQuote & Format(dDate, "yyyy-mm-dd") & cQuote & ") or" & _
               " (date_format(a.date,'%m %d')=" & cQuote & Format(dDate, "mmm dd") & cQuote & ")"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then ChkHoliday = objdbRs("description")
End Function

Private Sub Command10_Click()
    On Error GoTo ErrSave
    
    Dim nCtr As Integer, _
        cSqlStmt As String, _
        cSeries As String, _
        nAddEdit As Integer, _
        nRowPos As Integer, _
        aTimeInfo As Variant
    
    nRowPos = MSHFlexGrid2.RowSel
    
    Select Case MsgBox("Update DTR of " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & "?", vbYesNoCancel, "Update DTR")
    
        Case vbYes
            With MSHFlexGrid3
            
                ShowProgress 0
                
                cSqlStmt = "delete from pa84650 where logdate=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & _
                           " and empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                For nCtr = 1 To (.Rows - 1)
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    
                    If Trim(.TextMatrix(nCtr, 6)) <> "" Then
                    
                        If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                            cSeries = .TextMatrix(nCtr, 1)
                            nAddEdit = 2
                        Else
                            cSeries = GenerateSeries("bio")
                            While IfExists("pa84650", "pa84650.tran_no=" & cQuote & PadStr(cSeries, "0", 10) & cQuote)
                                cSeries = GenerateSeries("bio")
                            Wend
                            cSeries = PadStr(cSeries, "0", 10)
                            nAddEdit = 1
                        End If
                        
                        cSqlStmt = "insert into pa84650(tran_no,empid,logdate,transdate,trantime,trantype,shiftid)values(" & _
                                   cQuote & cSeries & cQuote & "," & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                   cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 2), "yyyy-mm-dd") & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 6), "HH:MM:SS") & cQuote & "," & _
                                   Val(.TextMatrix(nCtr, 4)) & "," & _
                                   cQuote & IIf((Trim(Text2.Text) <> "") And (Text2.Text <> MSHFlexGrid2.TextMatrix(nRowPos, 12)), Text2.Text, MSHFlexGrid2.TextMatrix(nRowPos, 12)) & cQuote & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                        
                        Log2Audit Name, IIf(nAddEdit = 1, "Add ", "Update ") & " Tran #" & cSeries & " to EmpID #" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2)
                    End If
                Next nCtr
                
                ' --> retrieve computed dtr here... 20060907
                aTimeInfo = ComputeDays(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2), _
                                        Array(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd"), Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd"), 0), _
                                        Array(Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 10)), Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 11))), Val(lblPClose.Caption) = 1)
                
                If IfExists("di36770", "(empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ") and (di36770.date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")") Then
                    cSqlStmt = "update di36770 set " & _
                               "  shiftid=" & cQuote & Text2.Text & cQuote & "," & _
                               "  time1=" & cQuote & Format(Label7.Caption, "HH:MM:SS") & cQuote & "," & _
                               "  time2=" & cQuote & Format(Label8.Caption, "HH:MM:SS") & cQuote & "," & _
                               "  reg_hr=" & aTimeInfo(0) * 8 & "," & _
                               "  reg_ot_hr=" & aTimeInfo(1) & "," & _
                               "  sa_reg_ot=" & aTimeInfo(2) & "," & _
                               "  nd_hr=" & aTimeInfo(3) * 8 & "," & _
                               "  nd_ot_hr=" & aTimeInfo(4) & "," & _
                               "  sa_nd_ot=" & aTimeInfo(12) & "," & _
                               "  sun_hr=" & aTimeInfo(5) & "," & _
                               "  sun_ot_hr=" & aTimeInfo(6) & "," & _
                               "  remark=" & cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), "No Entry or Absent", IIf(aTimeInfo(11) = 2, "On Leave", ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))), ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))) & cQuote & "," & _
                               "  tag=" & IIf(aTimeInfo(10) > 0, 3, IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), 1, IIf(aTimeInfo(11) = 2, 2, 0)), 0)) & _
                               " where (empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ") " & _
                               " and (di36770.date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")"
                Else
                    cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,time1,time2,`remark`,`tag`," & _
                               "reg_hr,reg_ot_hr,sa_reg_ot,nd_hr,nd_ot_hr,sa_nd_ot,sun_hr,sun_ot_hr)values(" & _
                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                               cQuote & Text1.Text & cQuote & "," & _
                               cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & "," & _
                               cQuote & Text2.Text & cQuote & "," & _
                               cQuote & Format(Label7.Caption, "HH:MM:SS") & cQuote & "," & _
                               cQuote & Format(Label8.Caption, "HH:MM:SS") & cQuote & "," & _
                               cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), "No Entry or Absent", IIf(aTimeInfo(11) = 2, "On Leave", ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))), ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))) & cQuote & "," & _
                               IIf(aTimeInfo(10) > 0, 3, IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), 1, IIf(aTimeInfo(11) = 2, 2, 0)), 0)) & "," & _
                               aTimeInfo(0) * 8 & "," & _
                               aTimeInfo(1) & "," & _
                               aTimeInfo(2) & "," & _
                               aTimeInfo(3) * 8 & "," & _
                               aTimeInfo(4) & "," & _
                               aTimeInfo(12) & "," & _
                               aTimeInfo(5) & "," & _
                               aTimeInfo(6) & ")"
                End If
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
    
    
                ' --> for security reason - 20070210
                Log2Audit Name, "Update DTR of EmpID#" & Text3(0).Text & " for " & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote
                
    
                ShowProgress 4
            End With
        
        Case vbNo
        
        Case vbCancel
            GoTo endsave
    End Select
    
    If nAddEdit > 0 Then CheckGrid MSHFlexGrid1.RowSel
   
    Label10.Visible = False
    Text2.Visible = False
    Command9.Visible = False
   
    Command6.Enabled = True
    MSHFlexGrid1.Enabled = True
    MSHFlexGrid2.Enabled = True
    MSHFlexGrid3.Enabled = True
    
    nAdd = 0
    CtrlPanel Me, nAdd
    BtnEnable 2
    
    MSHFlexGrid1_EnterCell
    
    MSHFlexGrid2.Row = nRowPos
    MSHFlexGrid2.ColSel = MSHFlexGrid2.Cols - MSHFlexGrid2.FixedCols
    MSHFlexGrid2_EnterCell
    MSHFlexGrid2.SetFocus
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command12_Click(Index As Integer)
    Select Case Index
        Case 0
            Select Case nTagSelect
                Case 1
                    GenerateReport "Daily Time Report ", "prv377.rpt"
                Case 2
                    GenerateReport "Daily Time Report ", "prv376.rpt"
                Case 3
                    GenerateReport "Daily Time Report (Summary)", "rpt387.rpt"
            End Select
            
        Case 1  ' --> Regular Report
            Select Case nTagSelect
                Case 1
                    GenerateReport "Daily Time Report ", "prv377A.rpt"
                Case 2
                    GenerateReport "Daily Time Report ", "prv376A.rpt"
                Case 3
                    GenerateReport "Daily Time Report (Summary)", "rpt387A.rpt"
            End Select
        
        Case 2  ' --> Extension Report
            Select Case nTagSelect
                Case 1
                    GenerateReport "Extension Daily Time Report ", "prv377E.rpt"
                Case 2
                    GenerateReport "Extension Daily Time Report ", "prv376E.rpt"
                Case 3
                    GenerateReport "Extension Daily Time Report (Summary)", "rpt387E.rpt"
            End Select
        
        Case 3
            XPPanel3.Visible = False
    End Select
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0  ' --> Summary
            If nTagSelect = 2 Then
                Command5.Enabled = False
                XPPanel2.Visible = True
                
'                Check1_Click
                
                OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINENAME", objdbRs, False
'                add2LstBox objdbRs, ListView1, Array("LINENAME", "LINEID")
            Else
            End If
        Case 2
            Frame1.Visible = False
    End Select
End Sub

Private Sub Command4_Click()
    Frame2.Visible = True
End Sub

Private Sub Command5_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo, App.Title) = vbYes Then
            Label10.Visible = False
            Text2.Visible = False
            Command9.Visible = False
           
            Command6.Enabled = True
            MSHFlexGrid1.Enabled = True
            MSHFlexGrid2.Enabled = True
            MSHFlexGrid3.Enabled = True
            
            nAdd = 0
            CtrlPanel Me, nAdd
            BtnEnable 2
            
            MSHFlexGrid1_EnterCell
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        cParam As String, _
        nCtr As Integer
    
    ShowProgress 0
    
    cSqlStmt = "select a.tcid, " & _
               "       a.empid, " & _
               "       concat(a.lastname,', ',a.firstname,' ',if(trim(a.mname)='',' ',concat(left(a.mname,1),'.'))) as fullname, " & _
               "       ifnull(b.posname,'') as position, " & _
               "       a.firstname, " & _
               "       a.lastname, " & _
               "       a.depid," & _
               "       ifnull(c.linename,'') as department, " & _
               "       a.active, " & _
               "       a.emp_stat, " & _
               "       a.wap, " & _
               "       round(ifnull(sum(d.reg_hr)/8,0),3) as reg_day, " & _
               "       round(ifnull(sum(d.reg_ot_hr),0),3) as reg_ot, " & _
               "       round(ifnull(sum(d.sa_reg_ot),0),3) as sa_reg_ot, " & _
               "       round(ifnull(sum(d.nd_hr)/8,0),3) as nd_day, " & _
               "       round(ifnull(sum(d.nd_ot_hr),0),3) as nd_ot, " & _
               "       round(ifnull(sum(d.sa_nd_ot),0),3) as sa_nd_ot, " & _
               "       round(ifnull(sum(d.sun_hr),0),3) as sun_hr, " & _
               "       round(ifnull(sum(d.sun_ot_hr),0),3) as sun_ot " & _
               "from di3670 a   left join " & IIf(Val(lblPClose.Caption) = 0, "di36770", "dih36770") & " d on a.empid=d.empid and d.date between " & cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & _
               " left join di7670 b on a.posid=b.posid " & _
               " left join di5463 c on a.depid=c.lineid " & _
               " where (((a.active=1) and ((a.date_res between " & cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or (a.date_res > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ") or (a.date_fin > " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and ((a.date_hire<=" & cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & ") or (a.date_hire between " & cQuote & Format(DTPicker5.Value, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(DTPicker1.Value, "yyyy-mm-dd") & cQuote & ")))"
    OpenQueryDNS cSqlStmt & " group by a.empid order by a.lastname,a.firstname", oTempADO, False
    If oTempADO.RecordCount > 0 Then
        QueryAttach oTempADO, MSHFlexGrid1, myArray
    
        nAdd = 0
        CtrlPanel Me, nAdd
    
        SetGridColumn myArray2, MSHFlexGrid2
        With MSHFlexGrid2
            .Redraw = False
            DoEvents
            For nCtr = 0 To DateDiff("d", DTPicker5.Value, DTPicker1.Value)
                .Rows = nCtr + 2
                .RowHeight(nCtr + 1) = 285
                
                .TextMatrix(nCtr + 1, 1) = DateAdd("d", nCtr, DTPicker5.Value)
                .TextMatrix(nCtr + 1, 2) = Format(DateAdd("d", nCtr, DTPicker5.Value), "ddd - mmm dd,yyyy")
            Next nCtr
            RefreshGrid MSHFlexGrid2, True
            .Redraw = True
        End With
    
        BtnEnable 2
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If

    ShowProgress 4
End Sub

Sub CreateTemp(ByVal nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTableName As String
        
    Select Case nMode
        Case 0
            cSqlStmt = " CREATE TABLE tmpDTR(   [EMPID] char(6)," & _
                       " [FULLNAME] char(100),  [POSITION] char(100)," & _
                       " [DEPID] char(3),       [DEPTNAME] char(100)," & _
                       " [EMP_STAT] integer,    [active] integer," & _
                       " [SDATE] date,          [EDATE] date," & _
                       " [REG_DAY] double,      [REG_OT_HR] double,     [SA_OT_HR] double,     " & _
                       " [ND_DAY] double,       [ND_OT_HR] double,      [SAND_OT_HR] double," & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [HOLIDAY] double)"
            cTableName = "tmpDTR"
        
        Case 1
            cSqlStmt = " CREATE TABLE tmp84650( [wap] integer," & _
                       " [EMPID] char(6),       [TRAN_NO] char(10)," & _
                       " [FULLNAME] char(100),  [DEPTNAME] char(100)," & _
                       " [DAY_DATE] date,       [DAY_NAME] char(20)," & _
                       " [RegHour] double,      [OTHour] double, " & _
                       " [SAOT] double,         [NDiff] double, " & _
                       " [NDiffOT] double,      [SANDOT] double, " & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [LOGDATE] date,        [TRANSDATE] date," & _
                       " [SHIFTDESC] char(100), [REMARK] char(100)," & _
                       " [TIME1] char(15),      [TIME2] char(15), " & _
                       " [SEQ_NO] integer,      [emp_stat] integer)"
            cTableName = "tmp84650"
        
        Case 2
            cSqlStmt = " CREATE TABLE tmpDTRD(  " & _
                       " [emp_stat] integer,    [wap] integer," & _
                       " [EMPID] char(6),       [TRAN_NO] char(10), " & _
                       " [FULLNAME] char(100),  [DEPTNAME] char(100), " & _
                       " [DAY_DATE] date,       [DAY_NAME] char(20), " & _
                       " [RegHour] double,      [OTHour] double, " & _
                       " [SAOT] double,         [NDiff] double, " & _
                       " [NDiffOT] double,      [SANDOT] double, " & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [LOGDATE] date,        [TRANSDATE] date," & _
                       " [outtrantime] char(15),[intrantime] char(15), " & _
                       " [SHIFTDESC] char(100), [REMARK] char(100)," & _
                       " [TIME1] char(15),      [TIME2] char(15)," & _
                       " [SEQ_NO] integer,      [tag] integer)"
            cTableName = "tmpDTRD"
        
    End Select
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM " & cTableName, oTempADO, True
End Sub

Private Sub Command7_Click(Index As Integer)
    Dim nCtr As Integer, _
        cSqlStmt As String
        
    Select Case Index
    
        Case 0      ' --> TMS Report
            ShowProgress 0
            
            CreateTemp 0
            
            With MSHFlexGrid1
                For nCtr = 1 To (.Rows - 1)
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    
                    cSqlStmt = "insert into tmpdtr(empid, fullname, [position], " & _
                               " deptname, [active], emp_stat, sdate, edate, " & _
                               " reg_day, reg_ot_hr, sa_ot_hr, nd_day, nd_ot_hr, sand_ot_hr,sun, sunot, holiday)values(" & _
                               cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 4) & cQuote & "," & _
                               cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & _
                               Val(.TextMatrix(nCtr, 9)) & "," & _
                               Val(.TextMatrix(nCtr, 10)) & "," & _
                               cQuote & Format(DTPicker5.Value, "mm/dd/yyyy") & cQuote & "," & _
                               cQuote & Format(DTPicker1.Value, "mm/dd/yyyy") & cQuote & "," & _
                               Val(.TextMatrix(nCtr, 12)) & "," & Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                               Val(.TextMatrix(nCtr, 15)) & "," & Val(.TextMatrix(nCtr, 16)) & "," & Val(.TextMatrix(nCtr, 17)) & "," & _
                               Val(.TextMatrix(nCtr, 18)) & "," & Val(.TextMatrix(nCtr, 19)) & ",0)"
                    QueryTemp cSqlStmt, objdbRs, True
                    
                Next nCtr
            End With
            
            ShowProgress 4
            
            Frame2.Visible = False
            nTagSelect = 3
            XPPanel3.Tag = 3
            XPPanel3.Visible = True
            
        Case 1, 2
            nTagSelect = Index
            Frame1.Visible = True
            
        Case 3
            Frame2.Visible = False
            
    End Select
End Sub

Private Sub Command8_Click()
    Dim cSqlStmt As String, _
        lProceed As Boolean
    
    
    ' --> added security as of 20070216
    If gUserGroup = 0 Then
        OpenQueryDNS "select isprocess, date_process from pa7730 where periodid=" & cQuote & Text1.Text & cQuote, objdbRs, False
        If objdbRs("isprocess") = 1 Then
            cSqlStmt = "This period had been processed for payroll as of " & Format(objdbRs("date_process"), "ddd - mmm dd, yyyy") & vbCrLf & _
                       "A representative from the Accounting Department is needed to allow you to modify this content."
            MsgBox cSqlStmt, vbCritical, "Warning!!!"
            
            frmManager.Tag = 1
            frmManager.Show 1
            If ModalResult = mrCancel Then Exit Sub
            lProceed = ModalResult = mrOk
        Else
            lProceed = True
        End If
    Else
        lProceed = True
    End If
    
    If Not lProceed Then
        MsgBox "You are not allowed to modify this content without affirmation from the Accounting Department!!!", vbExclamation, "System Advisory!!!"
        Exit Sub
    End If
    ' --> end of added security...
    
    
    If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 13)) = 2 Then
        cSqlStmt = "A leave is in effect for this date!!!" & vbCrLf & "You are not allowed to modify!!!"
        MsgBox cSqlStmt, vbCritical, "System Advisory!!!"
        Exit Sub
    End If
    
    Label10.Visible = True
    Text2.Visible = True
    Command9.Visible = True
    
    Command6.Enabled = False
    MSHFlexGrid1.Enabled = False
    MSHFlexGrid2.Enabled = False
    MSHFlexGrid3.Enabled = True
    
    nAdd = 2
    CtrlPanel Me, nAdd
    
    MSHFlexGrid3.Row = 1
    MSHFlexGrid3.SetFocus
End Sub

Private Sub Command9_Click()
    frmLookup.showPopup 9
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text2.Text = cResult
        OpenQueryDNS "SELECT * FROM PA74380 WHERE SHIFTID=" & cQuote & cResult & cQuote, objdbRs, False
        Label6.Caption = EncodeStr(objdbRs("description"))
        Label7.Caption = Format(objdbRs("time1"), "hh:mm AMPM")
        Label8.Caption = Format(objdbRs("time2"), "hh:mm AMPM")
    End If
End Sub

Private Sub dtFlex_DblClick()
    dtFlex_KeyDown vbKeyReturn, 0
End Sub

Private Sub dtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid3
        If KeyCode = vbKeyReturn Then
            dtFlex_LostFocus
            If .ColSel = 3 Then
                .TextMatrix(.Row, 2) = dtFlex.Value
                .TextMatrix(.Row, 3) = Format(dtFlex.Value, "ddd - mmm d, yyyy")
            Else
                .TextMatrix(.Row, 6) = Format(dtFlex.Value, "HH:MM:00")
            End If
            .SetFocus
        ElseIf KeyCode = vbKeyEscape Then
            dtFlex_LostFocus
            .SetFocus
        End If
    End With
End Sub

Private Sub dtFlex_LostFocus()
    dtFlex.Visible = False
    Command5.Cancel = True
End Sub

Private Sub Form_Load()
    Log2Audit Name, "Open"
    
    If lSuperUser Then
        lShow = True
    Else
        OpenQueryDNS "select userlevel from pa2360 where userid=" & cQuote & gUserID & cQuote, objdbRs, False
        lShow = objdbRs("userlevel") = 1
    End If
    
    aPeriodInfo = Array("", Now, Now)
    
    myArray = Array("TXT:[TCID]:6:True", _
                    "TXT:[Emp ID]:8:True", _
                    "TXT:[Fullname]:30:True", _
                    "TXT:[Position]:20:True", _
                    "TXT:[FName]:20:False", _
                    "TXT:[LName]:20:False", _
                    "NUM:[Dep ID]:3:False", _
                    "TXT:[Department]:20:True", _
                    "NUM:[Active]:1:False", _
                    "NUM:[emp stat]:1:False", _
                    "NUM:[WAP Status]:1:False", _
                    "NUM:[Days Work]:10:True", _
                    "NUM:[OT Hour]:9:True", _
                    "NUM:[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                    "NUM:[NDiff]:9:True", _
                    "NUM:[NDiff OT]:9:True", _
                    "NUM:[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                    "NUM:[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                    "NUM:[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"))
                    
    myArray2 = Array("DAT:[date]:10:False", _
                     "TXT:[Date]:20:True", _
                     "NUM:[Reg Hour]:9:True", _
                     "NUM:[OT Hour]:9:True", _
                     "NUM:[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                     "NUM:[NDiff]:9:True", _
                     "NUM:[NDiff OT]:9:True", _
                     "NUM:[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                     "NUM:[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "NUM:[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "TXT:[Remark]:25:True", _
                     "TXT:[Shift]:5:False", _
                     "NUM:[Leave Tag]:1:False")
                    
    myArray3 = Array("TXT:[Tran No]:12:True", _
                     "DAT:[date]:10:False", _
                     "TXT:[Date]:20:True", _
                     "NUM:[Type]:1:False", _
                     "TXT:[Type]:6:True", _
                     "TXT:[Time]:12:True")
                     
    Tag = nAccess_Tag
    nAdd = 0
    CtrlPanel Me, nAdd
    BtnEnable 0
    
    SetGridColumn myArray, MSHFlexGrid1
    SetGridColumn myArray2, MSHFlexGrid2
    SetGridColumn myArray3, MSHFlexGrid3
    
    Combo1.ListIndex = 1
    
    DTPicker1.Value = Now
    DTPicker5.Value = Now
    
    Label1.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label27.Caption = ""
    
    MSHFlexGrid1.Width = 14515
    
    XPFrame2.Visible = False
    
'    Command12(0).Visible = lShow
End Sub

Private Sub MSHFlexGrid1_DblClick()
    With MSHFlexGrid1
        If .Width = 4965 Then
            .Width = 14515
            XPFrame2.Visible = False
        Else
            .Width = 4965
            XPFrame2.Visible = True
            MSHFlexGrid1_EnterCell
        End If
    End With
End Sub

Private Sub MSHFlexGrid1_EnterCell()
    Dim nCtr As Integer, _
        aTimeInfo As Variant, _
        cSqlStmt As String

    If (MSHFlexGrid1.Width <> 4965) Then Exit Sub

    With MSHFlexGrid1
        Text3(0).Text = .TextMatrix(.RowSel, 2)
        Text3(1).Text = .TextMatrix(.RowSel, 1)
        Text3(2).Text = .TextMatrix(.RowSel, 3)
        Text3(3).Text = .TextMatrix(.RowSel, 8)
        Text3(4).Text = .TextMatrix(.RowSel, 4)
        Text3(5).Text = IIf(Trim(.TextMatrix(.RowSel, 10)) = "", "", IIf(Val(.TextMatrix(.RowSel, 10)) = 0, "WAP", IIf(Val(.TextMatrix(.RowSel, 10)) = 1, "Contractual", "Regular")))
        Label27.Caption = IIf(Val(.TextMatrix(.RowSel, 9)) = 0, "", IIf(Val(.TextMatrix(.RowSel, 9)) = 1, "Resigned", "Finished Contract"))
    End With

    DoEvents

    ShowProgress 0

    With MSHFlexGrid2
    
        .Redraw = False

        For nCtr = 0 To DateDiff("d", DTPicker5.Value, DTPicker1.Value)
            ShowProgress 2, (nCtr / DateDiff("d", DTPicker5.Value, DTPicker1.Value)) * 100

''    myArray2 = Array("DAT:1[date]:10:False", _
''                     "TXT:2[Date]:20:True", _
''                     "NUM:3[Reg Hour]:9:True", _
''                     "NUM:4[OT Hour]:9:True", _
''                     "NUM:5[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
''                     "NUM:6[NDiff]:9:True", _
''                     "NUM:7[NDiff OT]:9:True", _
''                     "NUM:8[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
''                     "NUM:9[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
''                     "NUM:0[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
''                     "TXT:1[Remark]:25:True", _
''                     "TXT:2[Shift]:5:False", _
''                     "NUM:3[Leave Tag]:1:False")
            ' --> retrieve assigned shift for the day...
            OpenQueryDNS "select a.shiftid, " & _
                         " a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, " & _
                         " a.nd_hr, a.nd_ot_hr, a.sa_nd_ot, " & _
                         " a.sun_hr, a.sun_ot_hr, " & _
                         " a.remark, a.tag " & _
                         "from " & IIf(Val(lblPClose.Caption) = 0, "di36770", "dih36770") & " a where a.empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                         " and a.date=" & cQuote & Format(DateAdd("d", nCtr, DTPicker5.Value), "yyyy-mm-dd") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                .TextMatrix(nCtr + 1, 3) = IIf(objdbRs("reg_hr") > 0, objdbRs("reg_hr"), "")
                .TextMatrix(nCtr + 1, 4) = IIf(objdbRs("reg_ot_hr") > 0, objdbRs("reg_ot_hr"), "")
                .TextMatrix(nCtr + 1, 5) = IIf(objdbRs("sa_reg_ot") > 0, objdbRs("sa_reg_ot"), "")
                .TextMatrix(nCtr + 1, 6) = IIf(objdbRs("nd_hr") > 0, objdbRs("nd_hr"), "")
                .TextMatrix(nCtr + 1, 7) = IIf(objdbRs("nd_ot_hr") > 0, objdbRs("nd_ot_hr"), "")
                .TextMatrix(nCtr + 1, 8) = IIf(objdbRs("sa_nd_ot") > 0, objdbRs("sa_nd_ot"), "")
                .TextMatrix(nCtr + 1, 9) = IIf(objdbRs("sun_hr") > 0, objdbRs("sun_hr"), "")
                .TextMatrix(nCtr + 1, 10) = IIf(objdbRs("sun_ot_hr") > 0, objdbRs("sun_ot_hr"), "")
                .TextMatrix(nCtr + 1, 11) = objdbRs("remark")
                .TextMatrix(nCtr + 1, 12) = objdbRs("shiftid")
                .TextMatrix(nCtr + 1, 13) = objdbRs("tag")
            Else
                .TextMatrix(nCtr + 1, 3) = ""
                .TextMatrix(nCtr + 1, 4) = ""
                .TextMatrix(nCtr + 1, 5) = ""
                .TextMatrix(nCtr + 1, 6) = ""
                .TextMatrix(nCtr + 1, 7) = ""
                .TextMatrix(nCtr + 1, 8) = ""
                .TextMatrix(nCtr + 1, 9) = ""
                .TextMatrix(nCtr + 1, 10) = ""
                .TextMatrix(nCtr + 1, 11) = ""
                .TextMatrix(nCtr + 1, 12) = ""
                .TextMatrix(nCtr + 1, 13) = ""
            End If

            HiLyt2 nCtr + 1, MSHFlexGrid2, IIf(Trim(.TextMatrix(nCtr + 1, 11)) = "", vbBlack, vbBlue)

            If Weekday(DateAdd("d", nCtr, DTPicker5.Value)) = vbSunday Then HiLyt2 nCtr + 1, MSHFlexGrid2, vbRed

        Next nCtr

        .Redraw = True
        
    End With
    
    ShowProgress 4

    MSHFlexGrid2_EnterCell
    
    MSHFlexGrid1.RowSel = MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid2_EnterCell()
    Dim cSqlStmt As String
    
    If MSHFlexGrid2.Cols > 2 Then
        Text2.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 12)
        cSqlStmt = "select * from PA74380 where shiftid=" & cQuote & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 12) & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Label6.Caption = EncodeStr(objdbRs("description"))
            Label7.Caption = Format(objdbRs("time1"), "hh:mm AMPM")
            Label8.Caption = Format(objdbRs("time2"), "hh:mm AMPM")
        Else
            Label6.Caption = ""
            Label7.Caption = ""
            Label8.Caption = ""
        End If
    Else
        Text2.Text = ""
        Label6.Caption = ""
        Label7.Caption = ""
        Label8.Caption = ""
    End If

    cSqlStmt = "select tran_no, " & _
           "       transdate, " & _
           "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
           "       trantype, " & _
           "       if(trantype=0,'In','Out') as trn_type, " & _
           "       trantime " & _
           " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
           " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
           "   and logdate=" & cQuote & Format(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1), "yyyy-mm-dd") & cQuote & _
           " order by transdate, trantime"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid3, myArray3, False, , , 1
    Else
        SetGridColumn myArray3, MSHFlexGrid3
    End If
End Sub

Private Sub MSHFlexGrid3_DblClick()
    MSHFlexGrid3_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid3_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid3
        Select Case KeyCode
            Case vbKeyDown
                If .Row = .Rows - 1 Then
                    If (Trim(.TextMatrix(.Rows - 1, 2)) <> "") And (Trim(.TextMatrix(.Rows - 1, 6)) <> "") Then
                        .AddItem "", .Rows
                        .RowHeight(.RowSel + 1) = 285
                        .Row = .RowSel + 1
                        .TopRow = .Row
                        
                        .TextMatrix(.Row, 2) = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)
                        .TextMatrix(.Row, 3) = Format(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1), "ddd - mmm d, yyyy")
                        
                        .LeftCol = 3
                        .Col = 3
                        .ColSel = 3
                    End If
                End If

            Case vbKeyUp
                If .Rows - 1 > 1 Then
                    If (Trim(.TextMatrix(.Rows - 1, 2)) = "") Or (Trim(.TextMatrix(.Rows - 1, 6)) = "") Then
                        .Rows = .Rows - 1
                    End If
                End If
                
            Case vbKeyInsert    ' --> 20050908
                If .TextMatrix(.RowSel, 2) <> "" Then
                    .AddItem "", .RowSel
                    .RowHeight(.RowSel) = 285
                    
                    RefreshGrid MSHFlexGrid3, True
                    
                    '.Row = .RowSel + 1
                    .SetFocus
                End If
        
            Case vbKeyReturn
                Select Case .ColSel
                    Case 3, 6
                        Command5.Cancel = False
                        If .ColSel = 6 Then
                            dtFlex.Format = dtpCustom
                            dtFlex.UpDown = True
                            dtFlex.CustomFormat = "hh:mm tt"
                            dtFlex.Width = 1400
                            dtFlex.Value = IIf(Trim(.Text) = "", Now, .Text)
                        Else
                            dtFlex.UpDown = False
                            dtFlex.Format = dtpLongDate
                            dtFlex.Width = 3495
                            dtFlex.Value = IIf(Trim(.TextMatrix(.Row, 2)) = "", MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1), .TextMatrix(.Row, 2))
                        End If
                        dtFlex.Visible = True
                        dtFlex.left = .CellLeft + .left - (dtFlex.Width - .CellWidth)
                        dtFlex.top = .CellTop + .top - 10
                        dtFlex.SetFocus
                    
                    Case 5
                        Command5.Cancel = False
                        cmbFlex.ZOrder 0
                        cmbFlex.Visible = True
                        cmbFlex.left = .CellLeft + .left - (cmbFlex.Width - .CellWidth)
                        cmbFlex.top = .CellTop + .top - 10
                        cmbFlex.ListIndex = IIf(Trim(.Text) = "", 0, Val(.TextMatrix(.Row, 4)))
                        cmbFlex.SetFocus
                        
                End Select
            
            Case vbKeyDelete
                If (.RowSel < .Rows) Then
                    If (Trim(.TextMatrix(.RowSel, 2)) <> "") And (Trim(.TextMatrix(.Rows - 1, 6)) <> "") Then
                        If MsgBox("Delete Record ?", vbYesNo, App.Title) = vbYes Then
                            If .Rows - 1 = 1 Then
                                .AddItem "", .Rows
                                .RowHeight(.RowSel + 1) = 285
                            End If
                            .RemoveItem .RowSel
                        End If
                    Else
                        If (Trim(.TextMatrix(.RowSel, 2)) = "") Or (Trim(.TextMatrix(.Rows - 1, 6)) = "") Then
                            .RemoveItem .RowSel
                        End If
                    End If
                    
                    .SetFocus
                End If
            
        End Select
    End With
End Sub

Private Sub MSHFlexGrid3_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then
        KeyPreview = (Screen.ActiveForm.ActiveControl.Name <> "cmbFlex") And _
                     (Screen.ActiveForm.ActiveControl.Name <> "dtFlex")
    End If
End Sub

Private Sub Text10_Change()
    Dim nPos As Integer, _
        nLenStr As Integer, _
        lFound As Boolean
        
    nLenStr = Len(Text10.Text)
    
    Select Case Combo1.ListIndex
        Case 0
            nPos = 2
        Case 1
            nPos = 1
        Case 2
            nPos = 3
        Case 3
            nPos = 5
        Case 4
            nPos = 6
        Case 5
            nPos = 8
    End Select
        
    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        Do While .Row < .Rows - 1 And _
                 UCase(left(.TextMatrix(.Row, nPos), Len(Trim(Text10.Text)))) <> UCase(Trim(Text10.Text))
            .Row = .Row + 1
        Loop
        If .Row <> .Rows - 1 Then
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols
            .TopRow = .Row
            .RowSel = .Row
            .Refresh
        End If
        .Redraw = True
        
        If (Trim(Text10.Text) <> "") And (UCase(left(.TextMatrix(.Row, nPos), Len(Trim(Text10.Text)))) = UCase(Trim(Text10.Text))) Then
            Text10.Text = .TextMatrix(.Row, nPos)
            Text10.SelStart = nLenStr
            Text10.SelLength = Len(Text10.Text) - nLenStr
        End If
    End With
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38, 40
            With MSHFlexGrid1
                .SetFocus
                .Col = 1
                .ColSel = .Cols - .FixedCols
                .RowSel = .Row
            End With
            
        Case vbKeyBack
            If (Trim(Text10.Text) <> "") And (Text10.SelStart > 0) Then Text10.Text = left(Text10.Text, Text10.SelStart - 1)
        
        Case vbKeyReturn
            If Trim(Text10.Text) <> "" Then
                If MSHFlexGrid1.Width = 4965 Then
                    MSHFlexGrid1_EnterCell
                Else
                    MSHFlexGrid1_DblClick
                End If
            End If
    End Select
End Sub
