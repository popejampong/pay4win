VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaXPPanel.ocx"
Object = "{30DA1A2F-A970-4238-AC17-5773BA9DC841}#1.1#0"; "CIAXPDatePicker.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaXPFrame.ocx"
Begin VB.Form frmTMS2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Management System"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14685
   Begin VB.CheckBox Check2 
      Caption         =   "Exclude Close Period"
      Height          =   255
      Left            =   3345
      TabIndex        =   67
      Top             =   90
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Eastcam"
      Height          =   660
      Left            =   6315
      Picture         =   "frmTMS2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      Tag             =   "15"
      Top             =   8415
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Preview"
      Height          =   660
      Left            =   5310
      Picture         =   "frmTMS2.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "15"
      Top             =   8430
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   10875
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   2220
      Begin VB.CommandButton Command7 
         Caption         =   "Ca&ncel"
         Height          =   450
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1125
         Width           =   1965
      End
      Begin VB.CommandButton Command7 
         Caption         =   "DTR (&Detail)"
         Height          =   450
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   690
         Width           =   1965
      End
      Begin VB.CommandButton Command7 
         Caption         =   "DTR (&Summary)"
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   1965
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Save"
      Height          =   660
      Left            =   12720
      Picture         =   "frmTMS2.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "20"
      Top             =   8415
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   660
      Left            =   13695
      Picture         =   "frmTMS2.frx":4C86
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "21"
      Top             =   8415
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Edit"
      Height          =   660
      Left            =   11865
      Picture         =   "frmTMS2.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "16"
      Top             =   8415
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Preview"
      Height          =   660
      Left            =   10890
      Picture         =   "frmTMS2.frx":7F8A
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "15"
      Top             =   8415
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Generate"
      Height          =   660
      Left            =   9915
      Picture         =   "frmTMS2.frx":990C
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "22"
      Top             =   8415
      Width           =   855
   End
   Begin VB.ComboBox cmbFlex 
      Height          =   315
      ItemData        =   "frmTMS2.frx":B28E
      Left            =   5370
      List            =   "frmTMS2.frx":B298
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   7845
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   495
      Left            =   3855
      TabIndex        =   0
      Top             =   8445
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ciaXPPanel.XPPanel XPPanel3 
      Height          =   1860
      Left            =   7020
      TabIndex        =   11
      Top             =   3285
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   3281
      LicValid        =   -1  'True
      Begin VB.CommandButton Command12 
         Caption         =   "Close"
         Height          =   435
         Index           =   4
         Left            =   90
         TabIndex        =   55
         Top             =   1335
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Extension Report"
         Height          =   435
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   915
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Regular Report"
         Height          =   435
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   495
         Width           =   1800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Complete Report"
         Height          =   435
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   75
         Width           =   1800
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel2 
      Height          =   5715
      Left            =   5190
      TabIndex        =   15
      Top             =   1680
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
         Left            =   90
         TabIndex        =   18
         Top             =   5430
         Width           =   1290
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4515
         Picture         =   "frmTMS2.frx":B2A5
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "15"
         Top             =   285
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Close"
         Height          =   660
         Left            =   4515
         Picture         =   "frmTMS2.frx":CC27
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "21"
         Top             =   1065
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5190
         Left            =   90
         TabIndex        =   19
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
         Height          =   5700
         Left            =   4290
         Top             =   15
         Width           =   1320
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   30
         Width           =   3600
      End
   End
   Begin MSComCtl2.DTPicker dtFlex 
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   7350
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122355712
      CurrentDate     =   38381
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7410
      Left            =   90
      TabIndex        =   22
      Top             =   900
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   13070
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   4860
      Left            =   5085
      TabIndex        =   23
      Top             =   2085
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   8573
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
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   1350
      Left            =   9750
      TabIndex        =   24
      Top             =   6960
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   2381
      LicValid        =   -1  'True
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   315
         Left            =   1695
         TabIndex        =   26
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
         Left            =   1065
         TabIndex        =   25
         Top             =   135
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label10 
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
         TabIndex        =   33
         Top             =   195
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         TabIndex        =   32
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label4 
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
         TabIndex        =   31
         Top             =   765
         Width           =   1110
      End
      Begin VB.Label Label5 
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1035
         Width           =   2865
         WordWrap        =   -1  'True
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel6 
      Height          =   810
      Left            =   75
      TabIndex        =   34
      Top             =   8340
      Width           =   3690
      _ExtentX        =   6509
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
         ItemData        =   "frmTMS2.frx":E5A9
         Left            =   720
         List            =   "frmTMS2.frx":E5BF
         Style           =   2  'Dropdown List
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   420
         Width           =   2880
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   465
         Width           =   1575
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   1230
      Left            =   5100
      TabIndex        =   39
      Top             =   870
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
         Index           =   6
         Left            =   5925
         TabIndex        =   68
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   810
         Width           =   1440
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
         Index           =   2
         Left            =   1080
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   510
         Width           =   3450
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Alt Stat"
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
         TabIndex        =   69
         Top             =   870
         Width           =   990
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         Left            =   7485
         TabIndex        =   46
         Top             =   840
         Width           =   1830
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   1335
      Left            =   5100
      TabIndex        =   53
      Top             =   6975
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   810
      Left            =   90
      TabIndex        =   58
      Top             =   60
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   1429
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "By Period"
      TabPicture(0)   =   "frmTMS2.frx":E60D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "lblPClose"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "By Date Range"
      TabPicture(1)   =   "frmTMS2.frx":E629
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "XPDatePicker2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "XPDatePicker1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   -73530
         TabIndex        =   60
         Top             =   420
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
         Left            =   -74250
         TabIndex        =   59
         Top             =   420
         Width           =   690
      End
      Begin ciaXPDatePicker.XPDatePicker XPDatePicker1 
         Height          =   315
         Left            =   765
         TabIndex        =   64
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_SCHED"
         Top             =   390
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         FormatString    =   "dddd - MMM dd, yyyy"
         MouseIcon       =   "frmTMS2.frx":E645
         CalendarDayBorder=   -1  'True
         CalendarDayBorderColor=   -2147483646
         CalendarMonthBorderColor=   8421504
         LicValid        =   -1  'True
      End
      Begin ciaXPDatePicker.XPDatePicker XPDatePicker2 
         Height          =   315
         Left            =   3315
         TabIndex        =   65
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_SCHED"
         Top             =   390
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         FormatString    =   "dddd - MMM dd, yyyy"
         MouseIcon       =   "frmTMS2.frx":E661
         CalendarDayBorder=   -1  'True
         CalendarDayBorderColor=   -2147483646
         CalendarMonthBorderColor=   8421504
         LicValid        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Range"
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
         TabIndex        =   66
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -73080
         TabIndex        =   63
         Top             =   480
         Width           =   3600
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
         Left            =   -74865
         TabIndex        =   62
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblPClose 
         BackStyle       =   0  'Transparent
         Caption         =   "Period Close"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -70125
         TabIndex        =   61
         Top             =   465
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   885
      Left            =   9750
      Top             =   8325
      Width           =   4950
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
      ForeColor       =   &H000000C0&
      Height          =   630
      Left            =   6225
      TabIndex        =   54
      Top             =   75
      Width           =   8370
   End
End
Attribute VB_Name = "frmTMS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmTMS2
' description   :   Time Management System (by Employee)
' programmer    :   _-=[ srm ]=-_
' date          :   2 May 2006

Option Explicit

Dim oTempADO As New ADODB.Recordset, _
    nTagSelect, nAdd As Integer, _
    myArray, myArray2, myArray3, aPeriodInfo As Variant, _
    lShow As Boolean

Sub add2LstBox(ByVal oRecordSet As ADODB.Recordset, ByVal oListBox As ListView, ByVal aField As Variant)
    On Error GoTo ErrFillLst
    Dim lstItem As ListItem
    
    oListBox.ListItems.Clear
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
            Set lstItem = oListBox.ListItems.Add()
            lstItem.Text = objdbRs(aField(1))
            lstItem.SubItems(1) = objdbRs(aField(0))
            If UBound(aField) >= 2 Then
                lstItem.SubItems(2) = objdbRs(aField(2))
            End If
            oRecordSet.MoveNext
        Wend
    End If

    Exit Sub
    
ErrFillLst:     ' --> in case may error
    oListBox.ListItems.Clear
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
            MSHFlexGrid1.Width = 14505
            
            For nCtr = 0 To 5
                Text3(nCtr).Text = ""
            Next nCtr
            
            SetGridColumn myArray, MSHFlexGrid1
            SetGridColumn myArray2, MSHFlexGrid2
            SetGridColumn myArray3, MSHFlexGrid3
    
        Case 2      ' --> disable generate button
            CtrlPanel Me, nAdd
            Command6.Enabled = False
            
            ' --> for super user only, debugging/correction purposes... 20070908
            If Not lSuperUser Then
                If Val(lblPClose.Caption) = 1 Then Command8.Enabled = False
            End If
    End Select
End Sub

Function ChkPeriod() As Boolean
    Dim cString
    If (Text1.Text <> aPeriodInfo(0)) Or _
       (XPDatePicker1.CurrentDate <> aPeriodInfo(1)) Or _
       (XPDatePicker2.CurrentDate <> aPeriodInfo(2)) Then
        aPeriodInfo = Array(Text1.Text, XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate)
        BtnEnable 1
        cString = "Warning!!!" & vbCrLf & _
                  "The selected period is already close.  Revision are not anymore allowed." & vbCrLf & _
                  "You can still generate and preview the TMS for reference/archival purposes."
        Label18.Caption = IIf(Val(lblPClose.Caption) = 0, "", cString)
    End If
End Function

Sub CheckGrid(Optional ByVal nRow As Integer = 0)
    Dim nCtr As Integer, _
        cSqlStmt As String, _
        aTimeInfo As Variant
    
    With MSHFlexGrid1
        If nRow > 0 Then
        
            aTimeInfo = ComputeDays(.TextMatrix(nRow, 2), _
                                    Array(XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate, 0), _
                                    Array(Val(.TextMatrix(nRow, 5)), Val(.TextMatrix(nRow, 22)), Val(.TextMatrix(nRow, 9))), _
                                    Val(lblPClose.Caption) = 1)
            .TextMatrix(nRow, 11) = aTimeInfo(0)
            .TextMatrix(nRow, 12) = aTimeInfo(1)
            .TextMatrix(nRow, 13) = aTimeInfo(2)
            .TextMatrix(nRow, 14) = aTimeInfo(1) + aTimeInfo(2)
            .TextMatrix(nRow, 15) = aTimeInfo(3)
            .TextMatrix(nRow, 16) = aTimeInfo(4)
            .TextMatrix(nRow, 17) = aTimeInfo(12)
            .TextMatrix(nRow, 18) = aTimeInfo(4) + aTimeInfo(12)
            .TextMatrix(nRow, 19) = aTimeInfo(5)
            .TextMatrix(nRow, 20) = aTimeInfo(6)
            .TextMatrix(nRow, 23) = aTimeInfo(13)
            .TextMatrix(nRow, 24) = aTimeInfo(14)
        Else
            ShowProgress 0
            DoEvents
            For nCtr = 1 To (.Rows - 1)
                ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , , "retrieving computed attendance of " & .TextMatrix(nCtr, 3)
'                If .TextMatrix(nCtr, 2) = "055904" Then MsgBox "Test"
                aTimeInfo = ComputeDays(.TextMatrix(nCtr, 2), _
                                        Array(XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate, 0), _
                                        Array(Val(.TextMatrix(nCtr, 5)), Val(.TextMatrix(nCtr, 22)), Val(.TextMatrix(nCtr, 9))), _
                                        Val(lblPClose.Caption) = 1)
                .TextMatrix(nRow, 11) = Format(aTimeInfo(0), "#0.00")
                .TextMatrix(nRow, 12) = Format(aTimeInfo(1), "#0.00")
                .TextMatrix(nRow, 13) = Format(aTimeInfo(2), "#0.00")
                '.TextMatrix(nRow, 14) = Format(aTimeInfo(1) + aTimeInfo(2), "#0.00")
                .TextMatrix(nRow, 14) = Format(aTimeInfo(1), "#0.00") + Format(aTimeInfo(2), "#0.00")
                .TextMatrix(nRow, 15) = Format(aTimeInfo(3), "#0.00")
                .TextMatrix(nRow, 16) = Format(aTimeInfo(4), "#0.00")
                .TextMatrix(nRow, 17) = Format(aTimeInfo(12), "#0.00")
'                .TextMatrix(nRow, 18) = Format(aTimeInfo(4) + aTimeInfo(12), "#0.00")
                .TextMatrix(nRow, 18) = Format(aTimeInfo(4), "#0.00") + Format(aTimeInfo(12), "#0.00")
                .TextMatrix(nRow, 19) = Format(aTimeInfo(5), "#0.00")
                .TextMatrix(nRow, 20) = Format(aTimeInfo(6), "#0.00")
                .TextMatrix(nRow, 23) = Format(aTimeInfo(13), "#0.00")
                .TextMatrix(nRow, 24) = Format(aTimeInfo(14), "#0.00")

                .TopRow = nCtr
            Next nCtr
            ShowProgress 4
        End If
    End With
End Sub

Private Sub Check1_Click()
    Dim nCtr As Integer
    
    ListView1.Enabled = Check1.Value <> 1
    For nCtr = 1 To ListView1.ListItems.Count
        ListView1.ListItems(nCtr).Checked = Check1.Value = vbChecked
    Next nCtr
End Sub

Private Sub Combo1_Click()
    With MSHFlexGrid1
        .Redraw = False
        Select Case Combo1.ListIndex
            Case 0
                .Col = 2
            Case 1
                .Col = 1
            Case 2
                .Col = 3
            Case 3
                .Col = 6
            Case 4
                .Col = 7
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
        XPDatePicker1.CurrentDate = objdbRs("date_start")
        XPDatePicker2.CurrentDate = objdbRs("date_end")
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
        
    Dim VarSample As String
    
    nRowPos = MSHFlexGrid2.RowSel
    
    Select Case MsgBox("Update DTR of " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & "?", vbYesNoCancel, "Update DTR")
    
        Case vbYes
            If Val(lblPClose.Caption) = 1 Then
                cSqlStmt = "update dih36770 set " & _
                           "  shiftid=" & cQuote & Text2.Text & cQuote & "," & _
                           "  time1=" & cQuote & Format(Label7.Caption, "HH:MM:SS") & cQuote & "," & _
                           "  time2=" & cQuote & Format(Label8.Caption, "HH:MM:SS") & cQuote & _
                           " where (empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ") " & _
                           " and (`date`=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")"
                OpenQueryDNS cSqlStmt, objdbRs, True
                
                cSqlStmt = "update pah84650 set shiftid=" & cQuote & Text2.Text & cQuote & _
                           " where (empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ")" & _
                           " and (logdate=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")"
                OpenQueryDNS cSqlStmt, objdbRs, True
            Else
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
                                       cQuote & Text2.Text & cQuote & ")"
                            OpenQueryDNS cSqlStmt, objdbRs, True
                            Script2File cSqlStmt
                            
                            Log2Audit Name, IIf(nAddEdit = 1, "Add ", "Update ") & " Tran #" & cSeries & " to EmpID #" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2)
                        Else
                            'for blank entry
                            nAddEdit = 1
                        End If
                    Next nCtr
                    
                    aTimeInfo = ComputeDays(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2), _
                                             Array(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd"), Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd"), 0), _
                                            Array(Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 5)), Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)), Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9))), _
                                            Val(lblPClose.Caption) = 1)
                                            
                    VarSample = aTimeInfo(0) * 85
                                
                    If IfExists("di36770", "(empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ") and (di36770.date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")") Then
                        cSqlStmt = "update di36770 set " & _
                                   "  shiftid=" & cQuote & Text2.Text & cQuote & "," & _
                                   "  time1=" & cQuote & Format(Label7.Caption, "HH:MM:SS") & cQuote & "," & _
                                   "  time2=" & cQuote & Format(Label8.Caption, "HH:MM:SS") & cQuote & "," & _
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
                                   "  remark=" & cQuote & IIf(aTimeInfo(10) > 0, "Incomplete entry", IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), "No Entry or Absent", IIf(aTimeInfo(11) = 2, "On Leave", ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))), ChkHoliday(Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd")))) & cQuote & "," & _
                                   "  tag=" & IIf(aTimeInfo(10) > 0, 3, IIf(aTimeInfo(11) > 0, IIf((aTimeInfo(11) = 1) And (Trim(Text2.Text) <> ""), 1, IIf(aTimeInfo(11) = 2, 2, 0)), 0)) & _
                                   " where (empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & ") " & _
                                   " and (di36770.date=" & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote & ")"
                    Else
                        cSqlStmt = "insert into di36770(empid,periodid,`date`,shiftid,time1,time2,`remark`,`tag`," & _
                                   "reg_hr,reg_ot_hr,sa_reg_ot,tot_ot,nd_hr,nd_ot_hr,sa_nd_ot,nd_tot_ot,sun_hr,sun_ot_hr,sun_nd,sun_nd_ot)values(" & _
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
                    OpenQueryDNS cSqlStmt, objdbRs, True
                    Script2File cSqlStmt
        
        
                    ' --> for security reason - 20070210
                    Log2Audit Name, "Update DTR of EmpID#" & Text3(0).Text & " for " & cQuote & Format(MSHFlexGrid2.TextMatrix(nRowPos, 1), "yyyy-mm-dd") & cQuote
                    
'                    ' --> 20071005
'                    cSqlStmt = "update di2340 set dtr_update=1"
'                    OpenQueryDNS cSqlStmt, objdbRs, True
'                    Script2File cSqlStmt
        
                    ShowProgress 4
                End With
            End If
            
        Case vbNo
        
        Case vbCancel
            GoTo endsave
            
    End Select
    
    If nAddEdit > 0 Then CheckGrid MSHFlexGrid1.RowSel
'    CheckGrid MSHFlexGrid1.RowSel
   
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
    
    With MSHFlexGrid2
        .Row = nRowPos
        .ColSel = .Cols - .FixedCols
        MSHFlexGrid2_EnterCell
        .SetFocus
    End With
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
    XPPanel2.Visible = False
    Command5.Enabled = True
End Sub

Private Sub Command12_Click(Index As Integer)
    Dim nCtr As Integer, _
        BDepGuard As Boolean, _
        aWant As String
        
    Select Case Index
        Case 0
            Select Case nTagSelect
                Case 1
                    GenerateReport "Daily Time Report ", "prv377.rpt"
                Case 2
                    If gCompanyID <> "0003" Then
                    
                        If (gCompanyID = "0006") Or (gCompanyID = "0001") Then
                        
                            aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                            
                            GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv376AR.rpt", "prv376.rpt")
                            
                        Else
                            GenerateReport "Daily Time Report ", "prv376.rpt"
                        End If
                    Else
                        For nCtr = 1 To ListView1.ListItems.Count
                            If ListView1.ListItems(nCtr).Checked Then
                                If ListView1.ListItems(nCtr).Text = gDepid Then
                                    BDepGuard = True
                                Else
                                    BDepGuard = False
                                End If
                            End If
                        Next nCtr
                    
                        If BDepGuard = True Then
                        
                            Select Case MsgBox("Customized DTR Report for Security dept?", vbYesNoCancel, App.Title)
                            
                                Case vbYes
                                    GenerateReport "Daily Time Report ", "prv376G.rpt"
                                Case vbNo
                                    GenerateReport "Daily Time Report ", "prv376.rpt"
                            End Select
                        Else
                            GenerateReport "Daily Time Report ", "prv376.rpt"
                        End If
                        
                    End If
                Case 3
                    GenerateReport "Daily Time Report (Summary)", "rpt387.rpt"
            End Select
            
        Case 1  ' --> Regular Report
                aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                
                Select Case nTagSelect
                    Case 1
                        GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv377AR.rpt", "prv377ARSUN.rpt")
                    Case 2
                        GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv376AR.rpt", "prv376AR_SUN.rpt")
                    Case 3
                        GenerateReport "Daily Time Report (Summary)", IIf(aWant = "6", "rpt387AR.rpt", "rpt387AR_SUN.rpt")
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
        
        Case 4
            XPPanel3.Visible = False
            
    End Select
End Sub

Private Sub Command13_Click()
    Dim cSqlStmt As String, _
        aTimeInfo As Variant
        
    With MSHFlexGrid1
        aTimeInfo = ComputeDays(.TextMatrix(.RowSel, 2), _
                                Array(Format(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1), "yyyy-mm-dd"), Format(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1), "yyyy-mm-dd"), 0), _
                                Array(Val(.TextMatrix(.RowSel, 5)), Val(.TextMatrix(.RowSel, 20)), Val(.TextMatrix(.RowSel, 9))), _
                                IIf(Val(lblPClose.Caption) > 0, True, False))
                                
        MsgBox aTimeInfo(0) & vbCrLf & _
               aTimeInfo(1) & vbCrLf & _
               aTimeInfo(2) & vbCrLf & _
               aTimeInfo(3) & vbCrLf
    End With
End Sub

Sub CreateTemp2()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String, _
        cTableName As String
        
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
                   " [TOT_OT] double,       [ND_TOT_OT] double )"

        cTableName = "tmpDTRD"
    
    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
    
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM " & cTableName, oTempADO, True
End Sub


Private Sub Command14_Click()
    Dim cSqlStmt As String, _
        aTimeInfo As Variant, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        dLogDate As Date, _
        cYear As String, _
        oRecordSet As New Recordset, _
        oRSet As New Recordset, _
        oRSet2 As New Recordset
        
    CreateTemp 2
    
    
    OpenQueryDNS " select lastname,firstname,mname from di3670 " & _
                 " where empid =" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2) & cQuote, objdbRs, False
    

    aShiftInfo = Array("", "", "", "")
    aTrantype = Array("", "", "", "")
    
    
    cSqlStmt = " select a.EMPID, concat(a.LASTNAME,', ', a.FIRSTNAME,' ', left(a.MNAME,1),'.') as fullname, a.POSID, " & _
               " a.EMP_STAT, a.ACTIVE, a.PAYSTATUS, a.DATE_RES, a.WAP, b.PERIODID, b.DATE, c.date_start,c.date_end,duration, " & _
               " a.DEPID,ifnull(d.linename,'') as linename " & _
               " from di3670 a " & _
               " left join dih36770 b on a.empid=b.empid " & _
               " left join pa7730 c on b.periodid=c.periodid " & _
               " left join di5463 d on a.depid=d.lineid " & _
               " where a.lastname=" & cQuote & objdbRs("lastname") & cQuote & _
               " and a.firstname=" & cQuote & objdbRs("Firstname") & cQuote & _
               " and a.emp_stat <> 0 and a.wap=0 and a.paystatus<>2 " & _
               " and year(c.date_start) = " & cQuote & nAssessAmt & cQuote & _
               " group by b.periodid " & _
               " order by c.date_start desc, b.periodid, b.date "
    OpenQueryDNS cSqlStmt, oRecordSet, False
    If oRecordSet.RecordCount > 0 Then
        While Not oRecordSet.EOF
            
            cSqlStmt = " select EMPID, PERIODID, DATE, SHIFTID, DESCRIPTION, TIME1, TIME2, reg_hr, reg_ot_hr, " & _
                       " sa_reg_ot, nd_hr, nd_ot_hr, sa_nd_ot, sun_hr, sun_ot_hr, REMARK, CMPID, allowance, " & _
                       " TAG, sun_nd, sun_nd_ot, Inc_hr, tot_ot, nd_tot_ot from dih36770 " & _
                       " where empid = " & cQuote & oRecordSet("empid") & cQuote & _
                       " and periodid = " & cQuote & oRecordSet("periodid") & cQuote & _
                       " order by date desc "
                       
            OpenQueryDNS cSqlStmt, oRSet, False
            If oRSet.RecordCount > 0 Then
                While Not oRSet.EOF
                    
                    aShiftInfo = Array("", "", "", "")
                    aTrantype = Array("", "", "", "")

                    cSqlStmt = " SELECT a.empid, a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2, " & _
                               " a.tran_no,a.transdate,date_format(a.transdate,'%a - %b %e, %Y') as `day`,trantype,if(a.trantype=0,'In','Out') as trn_type,a.trantime " & _
                               " FROM pah84650 a " & _
                               " left join pa74380 b on a.shiftid = b.shiftid " & _
                               " Where a.empid = " & cQuote & oRecordSet("empid") & cQuote & _
                               " And a.logdate = " & cQuote & Format(oRSet("date"), "yyyy-mm-dd)") & cQuote & _
                               " order by a.logdate,a.transdate, a.trantime"
                    OpenQueryDNS cSqlStmt, oRSet2, False
                    If oRSet2.RecordCount > 0 Then
                        While Not oRSet2.EOF
                            
                            
                            aTrantype(3) = oRSet2("TRANSDATE")
                            If oRSet2("trantype") = 0 Then
                                If Trim(aTrantype(1)) = "" Then
                                    aTrantype(0) = oRSet2("trantype")
                                    aTrantype(1) = oRSet2("trantime")
                                    dLogDate = oRSet2("logdate")
                                End If
                            Else
                                aTrantype(0) = oRSet2("trantype")
                                aTrantype(2) = oRSet2("trantime")
                                
                                aShiftInfo(0) = oRSet("description")
                                
                                aShiftInfo(1) = Format(oRSet("time1"), "hh:mm AMPM")
                                aShiftInfo(2) = Format(oRSet("time2"), "hh:mm AMPM")
                                dLogDate = oRSet2("logdate")
                            End If
                                                        
                            oRSet2.MoveNext
                            
                            If Not oRSet2.EOF Then
                                If dLogDate = oRSet2("logdate") Then
                                    If (oRSet2("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                    
                                        'oRset("reg_hr")
                                        If aTrantype(2) <> "" Then
                                            If aTrantype(1) <> "" Then
                                                If oRSet("nd_hr") <> 0 Then
                                                    'for nDIFF
                                                    If aShiftInfo(2) <= Format("06:00 AM", "hh:mm AMPM") Then
                                                        If aShiftInfo(1) = Format("10:00 PM", "hh:mm AMPM") Then
                                                            aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            aTrantype(1) = DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))
                                                            aTimeInfo(11) = Format(DateAdd("h", 0, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 0, aShiftInfo(2)), "hh:mm AMPM")
                                                        Else
                                                            aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                            aTimeInfo(11) = ""
                                                            aTimeInfo(11) = Format(DateAdd("h", 2, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 2, aShiftInfo(2)), "hh:mm AMPM")
                                                        End If
                                                    Else
                                                        aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                        aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                    End If
                                                Else
                                                    If oRSet("reg_hr") >= 8 Then
                                                        'for regular
                                                        If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                            If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
        '                                                        MsgBox Hour(Format(aTrantype(1), "hh:mm AMPM"))
                                                                If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                End If
                                                            Else
                                                                If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                End If
                                                            End If
                                                        Else
                                                            If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                End If
                                                            Else
                                                                If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                    aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                            aTrantype(2) = ""
                                                            aTrantype(1) = ""
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If

                                        cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                   " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                   " LOGDATE,TRANSDATE, " & _
                                                   " intrantime,outtrantime," & _
                                                   " SHIFTDESC,REMARK," & _
                                                   " TIME1,TIME2," & _
                                                   " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                                   cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                                   cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                                   cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                                   oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                                   oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                                   oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                                   oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                                   cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                   cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                                   cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                                   oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                                   cQuote & oRecordSet("periodid") & cQuote & "," & _
                                                   cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
'                                        MsgBox cSqlStmt
                                        QueryTemp cSqlStmt, objdbRs, True
                                            
                                        aTrantype = Array("", "", "", "")
                                    End If
                                Else
                                    If aTrantype(2) <> "" Then
                                        If aTrantype(1) <> "" Then
                                            If oRSet("nd_hr") <> 0 Then
                                                'for nDIFF
                                                If aShiftInfo(2) <= Format("06:00 AM", "hh:mm AMPM") Then
                                                    If aShiftInfo(1) = Format("10:00 PM", "hh:mm AMPM") Then
                                                        aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                        aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        aTrantype(1) = DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))
                                                        aTimeInfo(11) = Format(DateAdd("h", 0, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 0, aShiftInfo(2)), "hh:mm AMPM")
                                                    Else
                                                        aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                        aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                        aTimeInfo(11) = ""
                                                        aTimeInfo(11) = Format(DateAdd("h", 2, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 2, aShiftInfo(2)), "hh:mm AMPM")
                                                    End If
                                                Else
                                                    aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                End If
                                            Else
                                                If oRSet("reg_hr") >= 8 Then
                                                    'for regular
                                                    If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                        If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
    '                                                        MsgBox Hour(Format(aTrantype(1), "hh:mm AMPM"))
                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            End If
                                                        Else
                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            End If
                                                        End If
                                                    Else
                                                        If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            End If
                                                        Else
                                                            If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                                aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                        aTrantype(2) = ""
                                                        aTrantype(1) = ""
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                               " LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime," & _
                                               " SHIFTDESC,REMARK," & _
                                               " TIME1,TIME2," & _
                                               " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                               cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                               cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                               oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                               oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                               oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                               oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                               cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                               oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                               cQuote & oRecordSet("periodid") & cQuote & "," & _
                                               cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
                                    QueryTemp cSqlStmt, objdbRs, True
                                    
                                    aTrantype = Array("", "", "", "")
                                End If
                            Else
                                If aTrantype(2) <> "" Then
                                    If aTrantype(1) <> "" Then
                                        If oRSet("nd_hr") <> 0 Then
                                            'for nDIFF
                                            If aShiftInfo(2) <= Format("06:00 AM", "hh:mm AMPM") Then
                                                If aShiftInfo(1) = Format("10:00 PM", "hh:mm AMPM") Then
                                                    aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                    aTrantype(1) = DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))
                                                    aTimeInfo(11) = Format(DateAdd("h", 0, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 0, aShiftInfo(2)), "hh:mm AMPM")
                                                Else
                                                    aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                    aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                    aTrantype(1) = DateAdd("h", 2, Format(aTrantype(1), "hh:mm AMPM"))
                                                    aTimeInfo(11) = ""
                                                    aTimeInfo(11) = Format(DateAdd("h", 2, aShiftInfo(1)), "hh:mm AMPM") & " - " & Format(DateAdd("h", 2, aShiftInfo(2)), "hh:mm AMPM")
                                                End If
                                            Else
                                                aTrantype(2) = Hour(DateAdd("h", aTimeInfo(3) + aTimeInfo(4) + 1, Hour(DateAdd("h", 0, Format(aTrantype(1), "hh:mm AMPM"))) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                            End If
                                        Else
                                            If oRSet("reg_hr") >= 8 Then
                                                'for regular
                                                If Minute(Format(aTrantype(1), "hh:mm AMPM")) > 5 Then
                                                    If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
'                                                        MsgBox Hour(Format(aTrantype(1), "hh:mm AMPM"))
                                                        If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        End If
                                                    Else
                                                        If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) + 1 & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(2), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        End If
                                                    End If
                                                Else
                                                    If 12 > Hour(Format(aTrantype(1), "hh:mm AMPM")) Then
                                                        If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Format(aTrantype(1), "hh:mm AMPM")) Then
                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr") + 1, Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        End If
                                                    Else
                                                        If aTrantype(2) > DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Format(aTrantype(1), "hh:mm AMPM")) Then
                                                            aTrantype(2) = Hour(DateAdd("h", oRSet("reg_hr") + oRSet("reg_ot_hr"), Hour(Format(aTrantype(1), "hh:mm AMPM")) & ":" & Minute(Format(aTrantype(2), "hh:mm AMPM")) & ":" & Second(Format(aTrantype(1), "hh:mm AMPM")))) & ":" & "0" & left(Minute(Format(aTrantype(2), "hh:mm AMPM")), 1)
                                                            aTrantype(2) = Format(aTrantype(2), "hh:mm AMPM")
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If oRSet("reg_hr") = 0 Or oRSet("reg_hr") = "" Then
                                                    aTrantype(2) = ""
                                                    aTrantype(1) = ""
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            
                                cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                           " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                           " LOGDATE,TRANSDATE, " & _
                                           " intrantime,outtrantime," & _
                                           " SHIFTDESC,REMARK," & _
                                           " TIME1,TIME2," & _
                                           " tag,SEQ_NO,TOT_OT,ND_TOT_OT,periodid,Duration)values(" & _
                                           cQuote & oRecordSet("empid") & cQuote & "," & cQuote & oRecordSet("Fullname") & cQuote & "," & _
                                           cQuote & oRecordSet("Linename") & cQuote & "," & oRecordSet("paystatus") & "," & oRecordSet("emp_stat") & "," & oRecordSet("wap") & "," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                           oRSet("Reg_Hr") & "," & oRSet("reg_ot_hr") & "," & oRSet("sa_reg_ot") & "," & _
                                           oRSet("nd_hr") & "," & oRSet("nd_ot_hr") & "," & oRSet("sa_nd_ot") & "," & _
                                           oRSet("sun_hr") & "," & oRSet("sun_nd") & "," & _
                                           oRSet("sun_ot_hr") & "," & oRSet("nd_tot_ot") & "," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                           cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                           cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                           cQuote & EncodeStr2(oRSet("remark")) & cQuote & "," & _
                                           cQuote & aShiftInfo(1) & cQuote & "," & cQuote & aShiftInfo(2) & cQuote & "," & _
                                           oRSet("tag") & "," & oRSet.AbsolutePosition & "," & oRSet("tot_ot") & "," & oRSet("nd_tot_ot") & "," & _
                                           cQuote & oRecordSet("periodid") & cQuote & "," & _
                                           cQuote & EncodeStr2(oRecordSet("duration")) & cQuote & ")"
'                                MsgBox cSqlStmt
                                QueryTemp cSqlStmt, objdbRs, True
                                aTrantype = Array("", "", "", "")
                            End If
                            
                        Wend
                    End If
                    
                    oRSet.MoveNext
                Wend
            End If
            oRecordSet.MoveNext
        Wend
        
    End If


    ShowProgress 3

    GenerateReport "Daily Time Report", "prv376a2.rpt"

    ShowProgress 4
End Sub

Private Sub Command15_Click()
    Dim aTimeInfo As Variant, _
        cShiftid As String, _
        cSqlStmt As String, _
        nCtr As Integer, _
        oRSet As New ADODB.Recordset

    ShowProgress 0

    cSqlStmt = "select a.tcid, " & _
               "       a.empid, " & _
               "       concat(a.lastname,', ',a.firstname,' ',if(trim(a.mname)='',' ',concat(left(a.mname,1),'.'))) as fullname, " & _
               "       ifnull(b.posname,'') as position, a.emp_stat, " & _
               "       a.firstname, a.lastname, c.linename, " & _
               "       a.paystatus, a.active, d.shiftid, " & _
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
               "from di3670 a   left join " & IIf(Val(lblPClose.Caption) = 0, "di36770", "dih36770") & " d on a.empid=d.empid and d.date between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & _
               " left join di7670 b on a.posid=b.posid " & _
               " left join di5463 c on a.depid=c.lineid " & _
               " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ")))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))"
    OpenQueryDNS cSqlStmt & " group by a.empid order by a.lastname,a.firstname", oTempADO, False

    If oTempADO.RecordCount > 0 Then
        While Not oTempADO.EOF
            ShowProgress 2, (oTempADO.AbsolutePosition / oTempADO.RecordCount) * 100

            For nCtr = 0 To DateDiff("d", XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate)
                cSqlStmt = "select tran_no, " & _
                       "       transdate, " & _
                       "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                       "       trantype, " & _
                       "       if(trantype=0,'In','Out') as trn_type, " & _
                       "       trantime,shiftid " & _
                       " from pa84650 " & _
                       " where empid=" & cQuote & oTempADO("empid") & cQuote & _
                       "   and logdate=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & _
                       " order by transdate, trantime"
            '    Script2File cSqlStmt
                OpenQueryDNS cSqlStmt, oRSet, False
                If oRSet.RecordCount > 0 Then
                    While Not oRSet.EOF
                        If oRSet("trantime") <> "" Then

                        cShiftid = oRSet("shiftid")
                            If IfExists("pa84650", "(tran_no=" & cQuote & oRSet("tran_no") & cQuote & ") and (pa84650.logdate=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & ") and (pa84650.trantype=" & cQuote & oRSet("trantype") & cQuote & ")") Then
                                cSqlStmt = "update pa84650 set " & _
                                           "  shiftid=" & cQuote & cShiftid & cQuote & "," & _
                                           "  trantime=" & cQuote & Format(oRSet("trantime"), "HH:MM:SS") & cQuote & _
                                           " where (empid=" & cQuote & oTempADO("empid") & cQuote & ")" & _
                                           " And (logdate=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & ")" & _
                                           " and (trantype = " & cQuote & oRSet("trantype") & cQuote & ")"

                            Else
                                cSqlStmt = "insert into pa84650(tran_no,empid,logdate,transdate,trantime,trantype,shiftid)values(" & _
                                           cQuote & oRSet("tran_no") & cQuote & "," & _
                                           cQuote & oTempADO("empid") & cQuote & "," & _
                                           cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & "," & _
                                           cQuote & Format(oRSet("trantime"), "HH:MM:SS") & cQuote & "," & _
                                           oRSet("trantype") & "," & _
                                           cQuote & cShiftid & cQuote & ")"
                            End If
                            OpenQueryDNS cSqlStmt, objdbRs, True
'                            Script2File cSqlStmt
                        End If
                        oRSet.MoveNext
                    Wend
                End If
                
                aTimeInfo = ComputeDays(oTempADO("empid"), _
                                        Array(Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd"), Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd"), 0), _
                                        Array(oTempADO("emp_stat"), oTempADO("wap"), oTempADO("paystatus")), _
                                        Val(lblPClose.Caption) = 1)

                If IfExists("di36770", "(empid=" & cQuote & oTempADO("empid") & cQuote & ") and (di36770.date=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & ")") Then
                    cSqlStmt = "update di36770 set " & _
                               "  shiftid=" & cQuote & cShiftid & cQuote & "," & _
                               "  time1=" & cQuote & Format(Label7.Caption, "HH:MM:SS") & cQuote & "," & _
                               "  time2=" & cQuote & Format(Label8.Caption, "HH:MM:SS") & cQuote & "," & _
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
                               "  sun_nd_ot=" & aTimeInfo(14) & _
                               " where (empid=" & cQuote & oTempADO("empid") & cQuote & ") " & _
                               " and (di36770.date=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & ")"
                End If
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt

            Next nCtr

            oTempADO.MoveNext
        Wend
    End If

    ShowProgress 4

End Sub

Private Sub Command2_Click(Index As Integer)
    Dim cSqlStmt As String, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        aWant As String, _
        oRecordSet As New ADODB.Recordset, _
        oRset1 As New ADODB.Recordset
        
    Dim nCtr As Integer
    nCtr = 0
    Select Case Index
        
        Case 1  ' --> by Employee
            If nTagSelect = 1 Then
                CreateTemp 1
                aShiftInfo = Array("", "", "")
    
                OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380", oRecordSet, False
    
                ShowProgress 0
    
                With MSHFlexGrid2
                    For nCtr = 1 To (.Rows - 1)
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
    
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.Requery adAsyncFetch
                            oRecordSet.Find "shiftid='" & .TextMatrix(nCtr, 14) & "'"
                            If Not oRecordSet.EOF Then
                                aShiftInfo(0) = EncodeStr(oRecordSet("description"))
                                aShiftInfo(1) = Format(objdbRs("time1"), "hh:mm AMPM")
                                aShiftInfo(2) = Format(objdbRs("time2"), "hh:mm AMPM")
                            Else
                                aShiftInfo = Array("", "", "")
                            End If
                        End If
    
                        cSqlStmt = " insert into tmp84650(EMPID,FULLNAME,DEPTNAME,paystatus,DAY_DATE,DAY_NAME,RegHour," & _
                                   " OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,SUN,SUNOT,SUN_ND,SUN_ND_OT,SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no)values(" & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                   Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & _
                                   Val(.TextMatrix(nCtr, 6)) & "," & Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                   Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                   Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                   Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 15) & cQuote & "," & _
                                   cQuote & aShiftInfo(1) & cQuote & "," & _
                                   cQuote & aShiftInfo(2) & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                   nCtr & ")"
                                   
                        QueryTemp cSqlStmt, objdbRs, True
                        
                    Next nCtr
                End With
    
                ShowProgress 3
                
                aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv377AR.rpt", "prv377ARSUN.rpt")
                
                ShowProgress 4
            Else
                ShowProgress 0
                CreateTemp 2
                aShiftInfo = Array("", "", "")
                aTrantype = Array("", "", "", "")
    
                With MSHFlexGrid2
    
                    OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380", oRecordSet, False
    
                    For nCtr = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
    
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.Requery adAsyncFetch
                            oRecordSet.Find "shiftid='" & .TextMatrix(nCtr, 16) & "'"
                            If Not oRecordSet.EOF Then
                                aShiftInfo(0) = EncodeStr(oRecordSet("description"))
                                
                                aShiftInfo(1) = Format(oRecordSet("time1"), "hh:mm AMPM")
                                aShiftInfo(2) = Format(oRecordSet("time2"), "hh:mm AMPM")
                            Else
                                aShiftInfo = Array("", "", "")
                            End If
                        End If
                        
                        cSqlStmt = "select tran_no, " & _
                                   "       transdate, " & _
                                   "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                                   "       trantype, " & _
                                   "       if(trantype=0,'In','Out') as trn_type, " & _
                                   "       trantime " & _
                                   " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
                                   " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                                   "   and logdate=" & cQuote & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd") & cQuote & _
                                   " order by transdate, trantime "
                        OpenQueryDNS cSqlStmt, objdbRs, False
                        
                        cSqlStmt = "select tran_no, " & _
                                   "       transdate, " & _
                                   "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                                   "       trantype, " & _
                                   "       if(trantype=0,'In','Out') as trn_type, " & _
                                   "       trantime " & _
                                   " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
                                   " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                                   "   and logdate=" & cQuote & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd") & cQuote & _
                                   " order by transdate, trantime "
                                   
                        OpenQueryDNS cSqlStmt, oRset1, False
                        If oRset1.RecordCount > 0 Then
                            aTrantype = Array("", "", "", "")
                            While Not oRset1.EOF
                                aTrantype(3) = oRset1("TRANSDATE")
                                If oRset1("trantype") = 0 Then
                                    If Trim(aTrantype(1)) = "" Then
                                        aTrantype(0) = oRset1("trantype")
                                        aTrantype(1) = oRset1("trantime")
                                    End If
                                Else
                                    aTrantype(0) = oRset1("trantype")
                                    aTrantype(2) = oRset1("trantime")
                                End If
    
                                oRset1.MoveNext
    
                                If Not oRset1.EOF Then
                                    If (oRset1("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                    
                                        cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME, " & _
                                                   " RegHour,OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,sun,sunot,sun_nd,sun_nd_ot,LOGDATE,TRANSDATE, " & _
                                                   " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                                   Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & _
                                                   Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 6)) & "," & _
                                                   Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                                   Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                                   Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                                   Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                   cQuote & EncodeStr2(.TextMatrix(nCtr, 13)) & cQuote & "," & _
                                                   cQuote & aShiftInfo(1) & cQuote & "," & _
                                                   cQuote & aShiftInfo(2) & cQuote & "," & _
                                                   nCtr & ")"
                                                   
                                        QueryTemp cSqlStmt, objdbRs, True
    
                                        aTrantype = Array("", "", "", "")
                                    End If
                                Else
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME, " & _
                                               " RegHour,OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,sun,sunot,sun_nd,sun_nd_ot,LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                               Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & _
                                               Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 6)) & "," & _
                                               Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                               Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                               Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                               Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(.TextMatrix(nCtr, 13)) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & _
                                               cQuote & aShiftInfo(2) & cQuote & "," & _
                                               nCtr & ")"
                                    QueryTemp cSqlStmt, objdbRs, True
                                    
                                    aTrantype = Array("", "", "", "")
                                End If
                            Wend
                                    
                        Else
                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME," & _
                                       " SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                       cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                       cQuote & EncodeStr2(.TextMatrix(nCtr, 15)) & cQuote & "," & _
                                       cQuote & aShiftInfo(1) & cQuote & "," & _
                                       cQuote & aShiftInfo(2) & cQuote & "," & _
                                       nCtr & ")"
                                       
                            QueryTemp cSqlStmt, objdbRs, True
                        End If
                        
                        aTrantype = Array("", "", "", "")
                        
                    Next nCtr
    
                End With
    
                ShowProgress 3
                
                aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv376AR.rpt", "prv376AR_SUN.rpt")

                ShowProgress 4
                
            End If
            
            
        Case 2
            Command5.Enabled = False
            XPPanel2.Visible = True
            
            Check1_Click
            
            OpenQueryDNS "SELECT LINENAME, LINEID FROM DI5463 ORDER BY LINENAME", objdbRs, False
            add2LstBox objdbRs, ListView1, Array("LINENAME", "LINEID")
    
        
    End Select
    
'    Frame1.Visible = False
    Frame2.Visible = False
    
    Set oRecordSet = Nothing
    Set oRset1 = Nothing
    
End Sub


Private Function CustomDtr(ByVal nTime1 As String, ByVal nTime2 As String, ByVal nTimeVal As String) As Variant
    Dim nTimeDTRVal As Variant
    
    nTimeDTRVal = Array(0#, 0#)
    
    If (nTime1 = "") Or (nTime2 = "") Then
        nTimeDTRVal(0) = 0
        nTimeDTRVal(1) = 0
    Else
        If Val(nTimeVal) <> 0 Then
            nTimeDTRVal(0) = 0
            nTimeDTRVal(1) = DateDiff("h", nTime1, nTime2)
        Else
            nTimeDTRVal(0) = DateDiff("h", nTime1, nTime2)
            nTimeDTRVal(1) = 0
        End If
    End If
    
    CustomDtr = nTimeDTRVal
End Function


Private Sub Command3_Click()
    Dim cParam, _
        cSqlStmt As String, _
        cDepid As String, _
        nCtr As Integer, _
        aTInfo As Variant, _
        aTimeInfo As Variant, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        aTimeDtrVal As Variant, _
        dLogDate As Date, _
        oRecordSet As New ADODB.Recordset, _
        oRSet As New ADODB.Recordset, _
        lWap As Boolean

    ' --> for Department
    If Check1.Value = vbUnchecked Then
        For nCtr = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(nCtr).Checked Then
                cParam = cParam & cQuote & ListView1.ListItems(nCtr).Text & cQuote & ","
            End If
        Next nCtr
        
        If Trim(cParam) <> "" Then
            cParam = "(" & left(cParam, Len(cParam) - 1) & ")"
        Else
            MsgBox "Please specify an item to preview!", vbInformation, "TMS - " & App.Title
            ListView1.SetFocus
            Exit Sub
        End If
    Else
        cParam = ""
    End If

    CreateTemp nTagSelect

    ShowProgress 0
    
    With MSHFlexGrid1
    
        For nCtr = 1 To .Rows - 1
        
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
            If (InStr(1, cParam, MSHFlexGrid1.TextMatrix(nCtr, 21), vbTextCompare)) Or (Trim(cParam) = "") Then
            
                aShiftInfo = Array("", "", "", "")
                aTrantype = Array("", "", "", "")
                
                aTimeDtrVal = Array(0#, 0#)
'
'                If .TextMatrix(nCtr, 2) = "017837" Then
'                    MsgBox "stop"
'                End If
                
                If nTagSelect = 1 Then
                    cSqlStmt = " select distinct a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2 from " & IIf(Val(lblPClose.Caption) = 0, "pa", "pah") & "84650 a " & _
                               " left join pa74380 b on a.shiftid = b.shiftid " & _
                               " where (a.empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & _
                               ") and (a.logdate between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ")"
                Else
                    cSqlStmt = " select a.empid, a.logdate, a.shiftid,ifnull(b.description,'') as description,b.time1,b.time2," & _
                               " a.tran_no,a.transdate,date_format(a.transdate,'%a - %b %e, %Y') as `day`,trantype,if(a.trantype=0,'In','Out') as trn_type,a.trantime " & _
                               " from " & IIf(Val(lblPClose.Caption) = 0, "pa", "pah") & "84650 a left join pa74380 b on a.shiftid = b.shiftid " & _
                               " where (a.empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & _
                               ") and (a.logdate between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ")" & _
                               " order by a.logdate,a.transdate, a.trantime"
                End If
                OpenQueryDNS cSqlStmt, oRecordSet, False
                
                
                If oRecordSet.RecordCount > 0 Then
                
                    While Not oRecordSet.EOF
                        If gCompanyID = "0001" Or gCompanyID = "0006" Then
                            cSqlStmt = "select a.reg_hr, a.reg_ot_hr, a.sa_reg_ot,(a.reg_ot_hr+ a.sa_reg_ot) as tot_ot, a.nd_hr, a.nd_ot_hr, (a.nd_ot_hr+a.sa_nd_ot) as nd_tot_ot, a.sun_hr, a.sun_ot_hr, " & _
                                       " 0, 0, 0, a.tag, a.sa_nd_ot, a.sun_nd, a.sun_nd_ot, a.remark,b.description,b.time1,b.time2 " & _
                                       "from " & IIf(Val(lblPClose.Caption) = 0, "di", "dih") & "36770  a " & _
                                       " left join pa74380 b on a.shiftid = b.shiftid " & _
                                       "where (a.empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & ")" & _
                                       " and (a.date=" & cQuote & Format(oRecordSet("logdate"), "yyyy-mm-dd") & cQuote & ")"
                        Else
                            cSqlStmt = "select a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.tot_ot, a.nd_hr, a.nd_ot_hr, a.nd_tot_ot, a.sun_hr, a.sun_ot_hr, " & _
                                       " 0, 0, 0, a.tag, a.sa_nd_ot, a.sun_nd, a.sun_nd_ot, a.remark,b.description,b.time1,b.time2 " & _
                                       "from " & IIf(Val(lblPClose.Caption) = 0, "di", "dih") & "36770  a " & _
                                       " left join pa74380 b on a.shiftid = b.shiftid " & _
                                       "where (a.empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & ")" & _
                                       " and (a.date=" & cQuote & Format(oRecordSet("logdate"), "yyyy-mm-dd") & cQuote & ")"
                        End If
'                        cSqlStmt = "select a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.tot_ot, a.nd_hr, a.nd_ot_hr, a.nd_tot_ot, a.sun_hr, a.sun_ot_hr, " & _
'                                   " 0, 0, 0, a.tag, a.tag, a.sa_nd_ot, a.sun_nd, a.sun_nd_ot, a.remark,b.description,b.time1,b.time2 " & _
'                                   "from " & IIf(Val(lblPClose.Caption) = 0, "di", "dih") & "36770  a " & _
'                                   " left join pa74380 b on a.shiftid = b.shiftid " & _
'                                   "where (a.empid=" & cQuote & .TextMatrix(nCtr, 2) & cQuote & ")" & _
'                                   " and (a.date=" & cQuote & Format(oRecordSet("logdate"), "yyyy-mm-dd") & cQuote & ")"
                                   
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
                                          
                        
                                          
'                        If Format(oRecordSet("logdate"), "yyyy-mm-dd") = "2009-03-25" Then MsgBox "stop"
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
                                       "SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no)values(" & _
                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                       cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
                                       cQuote & Format(oRecordSet("logdate"), "mm/dd/yyyy") & cQuote & "," & cQuote & Format(oRecordSet("logdate"), "dddd") & cQuote & "," & _
                                       aTimeInfo(0) & "," & aTimeInfo(1) & "," & aTimeInfo(2) & "," & _
                                       aTimeInfo(3) & "," & aTimeInfo(4) & "," & aTimeInfo(12) & "," & _
                                       aTimeInfo(5) & "," & aTimeInfo(6) & "," & _
                                       aTimeInfo(13) & "," & aTimeInfo(14) & "," & _
                                       aTimeInfo(15) & "," & aTimeInfo(16) & "," & _
                                       cQuote & EncodeStr2(oRecordSet("description")) & cQuote & "," & _
                                       cQuote & EncodeStr2(aTimeInfo(11)) & cQuote & "," & _
                                       cQuote & objdbRs("time1") & cQuote & "," & cQuote & objdbRs("time2") & cQuote & "," & _
                                       cQuote & Format(oRecordSet("logdate"), "mm/dd/yyyy") & cQuote & "," & _
                                       nCtr & ")"
                            QueryTemp cSqlStmt, objdbRs, True
                            
                        Else
                            
                            aTrantype(3) = oRecordSet("TRANSDATE")
                            If oRecordSet("trantype") = 0 Then

                                If Trim(aTrantype(1)) = "" Then
                                    aTrantype(0) = oRecordSet("trantype")
                                    aTrantype(1) = oRecordSet("trantime")
                                    dLogDate = oRecordSet("logdate")
                                End If

                            Else
                                aTrantype(0) = oRecordSet("trantype")
                                aTrantype(2) = oRecordSet("trantime")
'                                If gCompanyID <> "0003" Then
                                    aShiftInfo(0) = oRecordSet("description")
                                    aShiftInfo(1) = oRecordSet("time1")
                                    aShiftInfo(2) = oRecordSet("time2")
'                                Else
'                                    aShiftInfo(0) = objdbRs("description")
'                                    aShiftInfo(1) = objdbRs("time1")
'                                    aShiftInfo(2) = objdbRs("time2")
'                                End If
                                dLogDate = oRecordSet("logdate")
                            End If
                        End If
                        
                        oRecordSet.MoveNext
                        
                        If nTagSelect <> 1 Then
                        
                            If Not oRecordSet.EOF Then
                                If dLogDate = oRecordSet("logdate") Then
                                    If (oRecordSet("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                        If gDepid <> MSHFlexGrid1.TextMatrix(nCtr, 21) Then
                                           
                                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                       " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                       " LOGDATE,TRANSDATE, " & _
                                                       " intrantime,outtrantime," & _
                                                       " SHIFTDESC,REMARK," & _
                                                       " TIME1,TIME2," & _
                                                       " tag,SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                                       cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
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
                                                       aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & ")"

                                            QueryTemp cSqlStmt, objdbRs, True
                                            aTrantype = Array("", "", "", "")
                                        Else
                                            cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                                       " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                                       " LOGDATE,TRANSDATE, " & _
                                                       " intrantime,outtrantime," & _
                                                       " SHIFTDESC,REMARK," & _
                                                       " TIME1,TIME2," & _
                                                       " tag,SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                                       cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                                       cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
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
                                                       aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & ")"
                                                       
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
                                               " tag, SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                               cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                               cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
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
                                               aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & ")"
                                    QueryTemp cSqlStmt, objdbRs, True
                                    aTrantype = Array("", "", "", "")

                                End If
                            Else
                                If gDepid <> MSHFlexGrid1.TextMatrix(nCtr, 21) Then
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME, " & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                               cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                               cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
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
                                               nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & ")"
                                    QueryTemp cSqlStmt, objdbRs, True
                                    aTrantype = Array("", "", "", "")
                                Else
                                    cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME," & _
                                               " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT," & _
                                               " LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime," & _
                                               " SHIFTDESC,REMARK," & _
                                               " TIME1,TIME2," & _
                                               " tag,SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                               cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                               cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
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
                                               aTimeInfo(10) & "," & nCtr & "," & aTimeInfo(15) & "," & aTimeInfo(16) & ")"
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
                                   " OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,TOT_OT,ND_TOT_OT,SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no)values(" & _
                                   cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & _
                                   Val(.TextMatrix(nCtr, 5)) & "," & _
                                   Val(.TextMatrix(nCtr, 22)) & "," & _
                                   cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & _
                                   "'',0,0,0,0,0,0,0,0,0,0,0,0,'','','',''," & _
                                   cQuote & Format(Now, "yyyy-mm-dd") & cQuote & "," & nCtr & ")"
                        QueryTemp cSqlStmt, objdbRs, True
                    Else
                        If gCompanyID <> "0003" Then '20080328 custom setting for mico only
                            If .TextMatrix(nCtr, 5) <> 0 And .TextMatrix(nCtr, 9) = 0 Then
                                cSqlStmt = " insert into tmpDTRD(EMPID,FULLNAME,DEPTNAME,paystatus,emp_stat,wap,DAY_DATE,DAY_NAME, " & _
                                           " RegHour,OTHour,SAOT,NDiff,NDiffOT,SANDOT,SUN,SUNOT,SUN_ND,SUN_ND_OT,LOGDATE,TRANSDATE, " & _
                                           " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO,TOT_OT,ND_TOT_OT)values(" & _
                                           cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & cQuote & .TextMatrix(nCtr, 3) & cQuote & "," & _
                                           cQuote & .TextMatrix(nCtr, 8) & cQuote & "," & Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 22)) & "," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "dddd") & cQuote & "," & _
                                           "0,0,0,0,0,0,0,0,0,0," & _
                                           cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & cQuote & Format(dLogDate, "mm/dd/yyyy") & cQuote & "," & _
                                           "'','','','','',''," & _
                                           nCtr & ",0,0 )"
                                QueryTemp cSqlStmt, objdbRs, True
                                aTrantype = Array("", "", "", "")
                            End If
                        End If
                    End If
                End If
            End If
jump:
        Next nCtr
    End With
    
    ShowProgress 3
    
    QueryTemp "select * from " & IIf(nTagSelect = 1, "tmp84650", "tmpDTRD"), objdbRs, False
    If objdbRs.RecordCount > 0 Then
        '2010-04-22
        If nTagSelect <> 1 Then
            QueryTemp "select * from tmpDTRD where intrantime=''", objdbRs, False
            If objdbRs.RecordCount > 0 Then
                QueryTemp "delete from tmpDTRD where intrantime=''", objdbRs, True
            End If
            
            QueryTemp "select * from tmpDTRD where outtrantime=''", objdbRs, False
            If objdbRs.RecordCount > 0 Then
                QueryTemp "delete from tmpDTRD where outtrantime=''", objdbRs, True
            End If
        End If
    
        If lExtension Then
            XPPanel1.Tag = nTagSelect
            XPPanel3.Visible = True
        Else
            Command12_Click 0
        End If
    Else
        MsgBox "No Report to Generate!", vbInformation, "TMS - " & App.Title
    End If
    
    ShowProgress 4

    Set oRecordSet = Nothing



End Sub

Private Sub Command4_Click()
    Command7(1).Enabled = MSHFlexGrid1.Width = 4965
    Command7(2).Enabled = MSHFlexGrid1.Width = 4965
    Frame2.Visible = True
End Sub

Private Sub Command5_Click()
    Dim cString As String
    
    If nAdd = 0 Then
        Unload Me
        
        OpenQueryDNS "update di2340 set dtr_update=1", objdbRs, True
        Script2File "update di2340 set dtr_update=1"
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Log2Audit Name, "Cancel Transation to EmpID #" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2)
            
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
            
            OpenQueryDNS "update di2340 set dtr_update=1", objdbRs, True
            Script2File "update di2340 set dtr_update=1"
                        
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
               "from di3670 a   left join " & IIf(Val(lblPClose.Caption) = 0, "di36770", "dih36770") & " d on a.empid=d.empid and d.date between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & _
               " left join di7670 b on a.posid=b.posid " & _
               " left join di5463 c on a.depid=c.lineid " & _
               " where (((a.active=1) or (a.active=3)) and ((a.date_res between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_res > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ")))) or " & _
               "       ((a.active=2) and ((a.date_fin between " & cQuote & Format(XPDatePicker1.CurrentDate, "yyyy-mm-dd") & cQuote & " and " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") or ((a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & ") and (a.date_fin > " & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))))" & _
               " or ((a.ACTIVE=0) and (a.date_hire<=" & cQuote & Format(XPDatePicker2.CurrentDate, "yyyy-mm-dd") & cQuote & "))"
    
    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt & " group by a.empid order by a.lastname,a.firstname", oTempADO, False

    If oTempADO.RecordCount > 0 Then
        QueryAttach oTempADO, MSHFlexGrid1, myArray, False, , , 1

        nAdd = 0
        CtrlPanel Me, nAdd

        SetGridColumn myArray2, MSHFlexGrid2
        With MSHFlexGrid2
            .Redraw = False
            DoEvents
            For nCtr = 0 To DateDiff("d", XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate)
                .Rows = nCtr + 2
                .RowHeight(nCtr + 1) = 285

                .TextMatrix(nCtr + 1, 1) = DateAdd("d", nCtr, XPDatePicker1.CurrentDate)
                .TextMatrix(nCtr + 1, 2) = Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "ddd - mmm dd,yyyy")
            Next nCtr
            RefreshGrid MSHFlexGrid2, True
            .Redraw = True
        End With

        BtnEnable 2
        Command8.Enabled = False
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
                       
            cSqlStmt = " CREATE TABLE tmpDTR( " & _
                       " [EMPID] char(6),       [paystatus] integer, " & _
                       " [FULLNAME] char(100),  [POSITION] char(100)," & _
                       " [DEPID] char(3),       [DEPTNAME] char(100)," & _
                       " [EMP_STAT] integer,    [active] integer," & _
                       " [SDATE] date,          [EDATE] date," & _
                       " [REG_DAY] double,      [REG_OT_HR] double,     [SA_OT_HR] double,          [TOT_OT] double," & _
                       " [ND_DAY] double,       [ND_OT_HR] double,      [ND_TOT_OT] double,         [SAND_OT_HR] double," & _
                       " [SUN] double,          [SUNOT] double, " & _
                       " [SUN_ND] double,       [SUN_ND_OT] double, " & _
                       " [HOLIDAY] double)"
                       
            cTableName = "tmpDTR"
        Case 1
            cSqlStmt = " CREATE TABLE tmp84650D( " & _
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
                       " [TOT_OT] double,       [ND_TOT_OT] double)"
                       
            cTableName = "tmp84650D"
        
        Case 2
            cSqlStmt = " CREATE TABLE tmpDTRDD(  [paystatus] integer, " & _
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
                       " [TOT_OT] double,       [ND_TOT_OT] double )"

            cTableName = "tmpDTRDD"
        
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
    Dim cSqlStmt As String, _
        nCtr As Integer, _
        nCtr2 As Integer, _
        aTrantype As Variant, _
        aShiftInfo As Variant, _
        oRecordSet As New ADODB.Recordset
    
        
    Dim aWant As Integer
    Dim oRset1 As New ADODB.Recordset
    
    Select Case Index
        Case 0
            CreateTemp 0
            
            ShowProgress 0
                
            With MSHFlexGrid1
                For nCtr = 1 To .Rows - 1
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100
                    
                    cSqlStmt = "insert into tmpdtr(empid, fullname, [position], " & _
                               " deptname, paystatus, emp_stat, [active], sdate, edate, " & _
                               " reg_day, reg_ot_hr, sa_ot_hr, tot_ot, nd_day, nd_ot_hr, sand_ot_hr, nd_tot_ot, sun, sunot, sun_nd, sun_nd_ot, holiday)values(" & _
                               cQuote & .TextMatrix(nCtr, 2) & cQuote & "," & _
                               cQuote & EncodeStr2(.TextMatrix(nCtr, 3)) & cQuote & "," & _
                               cQuote & EncodeStr2(.TextMatrix(nCtr, 4)) & cQuote & "," & _
                               cQuote & EncodeStr2(.TextMatrix(nCtr, 8)) & cQuote & "," & _
                               Val(.TextMatrix(nCtr, 9)) & "," & _
                               Val(.TextMatrix(nCtr, 5)) & "," & _
                               Val(.TextMatrix(nCtr, 10)) & "," & _
                               cQuote & Format(XPDatePicker1.CurrentDate, "mm/dd/yyyy") & cQuote & "," & _
                               cQuote & Format(XPDatePicker2.CurrentDate, "mm/dd/yyyy") & cQuote & "," & _
                               Val(.TextMatrix(nCtr, 11)) & "," & _
                               Val(.TextMatrix(nCtr, 12)) & "," & _
                               Val(.TextMatrix(nCtr, 13)) & "," & _
                               Val(.TextMatrix(nCtr, 14)) & "," & _
                               Val(.TextMatrix(nCtr, 15)) & "," & _
                               Val(.TextMatrix(nCtr, 16)) & "," & _
                               Val(.TextMatrix(nCtr, 17)) & "," & _
                               Val(.TextMatrix(nCtr, 18)) & "," & _
                               Val(.TextMatrix(nCtr, 19)) & "," & _
                               Val(.TextMatrix(nCtr, 20)) & "," & _
                               Val(.TextMatrix(nCtr, 23)) & "," & _
                               Val(.TextMatrix(nCtr, 24)) & ",0)"
                    QueryTemp cSqlStmt, objdbRs, True
                Next nCtr
            End With
            
            ShowProgress 4
            
            Frame2.Visible = False
            nTagSelect = 3
            XPPanel3.Tag = 3
            XPPanel3.Visible = True
            
        Case 1
            'Frame1.Visible = True
            nTagSelect = 1
            
            CreateTemp 1
                aShiftInfo = Array("", "", "")
    
                OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380", oRecordSet, False
    
                ShowProgress 0
    
                With MSHFlexGrid2
                    For nCtr = 1 To (.Rows - 1)
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
    
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.Requery adAsyncFetch
                            oRecordSet.Find "shiftid='" & .TextMatrix(nCtr, 14) & "'"
                            If Not oRecordSet.EOF Then
                                aShiftInfo(0) = EncodeStr(oRecordSet("description"))
                                aShiftInfo(1) = Format(objdbRs("time1"), "hh:mm AMPM")
                                aShiftInfo(2) = Format(objdbRs("time2"), "hh:mm AMPM")
                            Else
                                aShiftInfo = Array("", "", "")
                            End If
                        End If
    
                        cSqlStmt = " insert into tmp84650D(EMPID,FULLNAME,DEPTNAME,paystatus,DAY_DATE,DAY_NAME,RegHour," & _
                                   " OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,SUN,SUNOT,SUN_ND,SUN_ND_OT,SHIFTDESC,[REMARK],TIME1,TIME2,logdate,seq_no)values(" & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                   Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & Val(.TextMatrix(nCtr, 5)) & "," & _
                                   Val(.TextMatrix(nCtr, 6)) & "," & Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                   Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                   Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                   Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 15) & cQuote & "," & _
                                   cQuote & aShiftInfo(1) & cQuote & "," & _
                                   cQuote & aShiftInfo(2) & cQuote & "," & _
                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                   nCtr & ")"
                                   
                        QueryTemp cSqlStmt, objdbRs, True
                        
                    Next nCtr
                End With
    
                ShowProgress 3
                
                aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv377ARD.rpt", "prv377ARSUND.rpt")
                
                ShowProgress 4
            
        Case 2
            'Frame1.Visible = True
            nTagSelect = 2
            
            ShowProgress 0
                CreateTemp 2
                aShiftInfo = Array("", "", "")
                aTrantype = Array("", "", "", "")
    
                With MSHFlexGrid2
    
                    OpenQueryDNS "select shiftid, `description`, time1, time2 from pa74380", oRecordSet, False
    
                    For nCtr = 1 To (.Rows - 1)
                    
                        ShowProgress 2, (nCtr / (.Rows - 1)) * 100
    
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.Requery adAsyncFetch
                            oRecordSet.Find "shiftid='" & .TextMatrix(nCtr, 16) & "'"
                            If Not oRecordSet.EOF Then
                                aShiftInfo(0) = EncodeStr(oRecordSet("description"))
                                
                                aShiftInfo(1) = Format(oRecordSet("time1"), "hh:mm AMPM")
                                aShiftInfo(2) = Format(oRecordSet("time2"), "hh:mm AMPM")
                            Else
                                aShiftInfo = Array("", "", "")
                            End If
                        End If
                        
                        cSqlStmt = "select tran_no, " & _
                                   "       transdate, " & _
                                   "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                                   "       trantype, " & _
                                   "       if(trantype=0,'In','Out') as trn_type, " & _
                                   "       trantime " & _
                                   " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
                                   " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                                   "   and logdate=" & cQuote & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd") & cQuote & _
                                   " order by transdate, trantime "
                        OpenQueryDNS cSqlStmt, objdbRs, False
                        
                        cSqlStmt = "select tran_no, " & _
                                   "       transdate, " & _
                                   "       date_format(transdate,'%a - %b %e, %Y') as `day`, " & _
                                   "       trantype, " & _
                                   "       if(trantype=0,'In','Out') as trn_type, " & _
                                   "       trantime " & _
                                   " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
                                   " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                                   "   and logdate=" & cQuote & Format(.TextMatrix(nCtr, 1), "yyyy-mm-dd") & cQuote & _
                                   " order by transdate, trantime "
                                   
                        OpenQueryDNS cSqlStmt, oRset1, False
                        If oRset1.RecordCount > 0 Then
                            aTrantype = Array("", "", "", "")
                            While Not oRset1.EOF
                                aTrantype(3) = oRset1("TRANSDATE")
                                If oRset1("trantype") = 0 Then
                                    If Trim(aTrantype(1)) = "" Then
                                        aTrantype(0) = oRset1("trantype")
                                        aTrantype(1) = oRset1("trantime")
                                    End If
                                Else
                                    aTrantype(0) = oRset1("trantype")
                                    aTrantype(2) = oRset1("trantime")
                                End If
    
                                oRset1.MoveNext
    
                                If Not oRset1.EOF Then
                                    If (oRset1("trantype") = 0) And (Trim(aTrantype(2)) <> "") Then
                                    
                                        cSqlStmt = " insert into tmpDTRDD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME, " & _
                                                   " RegHour,OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,sun,sunot,sun_nd,sun_nd_ot,LOGDATE,TRANSDATE, " & _
                                                   " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                                   cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                                   Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & _
                                                   Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 6)) & "," & _
                                                   Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                                   Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                                   Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                                   Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                                   cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                                   cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                                   cQuote & EncodeStr2(.TextMatrix(nCtr, 13)) & cQuote & "," & _
                                                   cQuote & aShiftInfo(1) & cQuote & "," & _
                                                   cQuote & aShiftInfo(2) & cQuote & "," & _
                                                   nCtr & ")"
                                                   
                                        QueryTemp cSqlStmt, objdbRs, True
    
                                        aTrantype = Array("", "", "", "")
                                    End If
                                Else
                                    cSqlStmt = " insert into tmpDTRDD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME, " & _
                                               " RegHour,OTHour,SAOT,TOT_OT,NDiff,NDiffOT,SANDOT,ND_TOT_OT,sun,sunot,sun_nd,sun_nd_ot,LOGDATE,TRANSDATE, " & _
                                               " intrantime,outtrantime,SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                               cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                               Val(.TextMatrix(nCtr, 3)) & "," & Val(.TextMatrix(nCtr, 4)) & "," & _
                                               Val(.TextMatrix(nCtr, 5)) & "," & Val(.TextMatrix(nCtr, 6)) & "," & _
                                               Val(.TextMatrix(nCtr, 7)) & "," & Val(.TextMatrix(nCtr, 8)) & "," & _
                                               Val(.TextMatrix(nCtr, 9)) & "," & Val(.TextMatrix(nCtr, 10)) & "," & _
                                               Val(.TextMatrix(nCtr, 11)) & "," & Val(.TextMatrix(nCtr, 12)) & "," & _
                                               Val(.TextMatrix(nCtr, 13)) & "," & Val(.TextMatrix(nCtr, 14)) & "," & _
                                               cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(3), "mm/dd/yyyy") & cQuote & "," & _
                                               cQuote & Format(aTrantype(1), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & Format(aTrantype(2), "hh:mm AMPM") & cQuote & "," & _
                                               cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                               cQuote & EncodeStr2(.TextMatrix(nCtr, 13)) & cQuote & "," & _
                                               cQuote & aShiftInfo(1) & cQuote & "," & _
                                               cQuote & aShiftInfo(2) & cQuote & "," & _
                                               nCtr & ")"
                                    QueryTemp cSqlStmt, objdbRs, True
                                    
                                    aTrantype = Array("", "", "", "")
                                End If
                            Wend
                                    
                        Else
                            cSqlStmt = " insert into tmpDTRDD(EMPID,FULLNAME,DEPTNAME,paystatus,wap,DAY_DATE,DAY_NAME," & _
                                       " SHIFTDESC,REMARK,TIME1,TIME2,SEQ_NO)values(" & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & "," & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3) & cQuote & "," & _
                                       cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 8) & cQuote & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 9)) & "," & Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 22)) & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 1), "mm/dd/yyyy") & cQuote & "," & _
                                       cQuote & Format(.TextMatrix(nCtr, 1), "dddd") & cQuote & "," & _
                                       cQuote & EncodeStr2(aShiftInfo(0)) & cQuote & "," & _
                                       cQuote & EncodeStr2(.TextMatrix(nCtr, 15)) & cQuote & "," & _
                                       cQuote & aShiftInfo(1) & cQuote & "," & _
                                       cQuote & aShiftInfo(2) & cQuote & "," & _
                                       nCtr & ")"
                                       
                            QueryTemp cSqlStmt, objdbRs, True
                        End If
                        
                        aTrantype = Array("", "", "", "")
                        
                    Next nCtr
    
                End With
    
                ShowProgress 3
                
                aWant = MsgBox("Audit Report?", vbYesNo + vbInformation, App.Title)
                GenerateReport "Daily Time Report ", IIf(aWant = "6", "prv376ARD.rpt", "prv376AR_SUND.rpt")

                ShowProgress 4
            
        Case 3
            Frame2.Visible = False
            
    End Select
    
    Set oRecordSet = Nothing
End Sub

Private Sub Command8_Click()
    Dim cSqlStmt As String, _
        lProceed As Boolean
    
    
    ' --> added security as of 20070216
    If (gUserGroup = 0) And Not lSuperUser Then
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
    
    
    If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 15)) > 0 Then
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
    
    Log2Audit Name, "Edit Transation to EmpID #" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2)
    
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
                    "NUM:[emp stat]:1:False", _
                    "TXT:[FName]:20:False", _
                    "TXT:[LName]:20:False", _
                    "TXT:[Department]:20:True", _
                    "NUM:[pay status]:1:False", _
                    "NUM:[Active]:1:False", _
                    "NUM:[Days Work]:10:True", _
                    "NUM:[OT Hour]:9:" & IIf(lShow, "True", "False"), _
                    "NUM:[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                    "NUM:[Tot OT]:9:False", _
                    "NUM:[NDiff]:9:True", _
                    "NUM:[NDiff OT]:9:" & IIf(lShow, "True", "False"), _
                    "NUM:[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                    "NUM:[TOT SA ND OT]:15:False", _
                    "NUM:[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                    "NUM:[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                    "NUM:[Dep ID]:3:False", _
                    "NUM:[WAP Status]:1:False", _
                    "NUM:[Sun ND]:10:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                    "NUM:[Sun ND OT]:10:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                    "NUM:[Inc Hour]:10:True")
    
    myArray2 = Array("DAT:[date]:10:False", _
                     "TXT:[Date]:20:True", _
                     "NUM:[Reg Hour]:9:True", _
                     "NUM:[OT Hour]:9:" & IIf(lShow, "True", "False"), _
                     "NUM:[SA OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                     "NUM:[Tot OT]:9:False", _
                     "NUM:[NDiff]:9:True", _
                     "NUM:[NDiff OT]:9:" & IIf(lShow, "True", "False"), _
                     "NUM:[SA ND OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "False"), _
                     "NUM:[TOT SA ND OT]:False", _
                     "NUM:[Sunday]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "NUM:[Sun OT]:9:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "NUM:[Sun ND]:10:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "NUM:[Sun ND OT]:10:" & IIf(lExtension, IIf(lShow, "True", "False"), "True"), _
                     "TXT:[Remark]:25:True", _
                     "TXT:[Shift]:5:False", _
                     "NUM:[Leave Tag]:1:False", _
                     "NUM:[Inc Hour]:10:True")
                    
    myArray3 = Array("TXT:[Tran No]:10:False", _
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
    Label1.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label27.Caption = ""
    MSHFlexGrid1.Width = 14505
    SSTab1.Tab = 0
    XPDatePicker1.CurrentDate = Now
    XPDatePicker2.CurrentDate = Now
    
    Command8.Enabled = False
    Command12(0).Visible = lShow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log2Audit Name, "CLOSE"

End Sub

Private Sub Form_Terminate()
    Set oTempADO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub MSHFlexGrid1_DblClick()
    Dim cSqlStmt As String
    With MSHFlexGrid1
        If .Width = 4965 Then
            .Width = 14505
            Command8.Enabled = False
        Else
            .Width = 4965
            Command8.Enabled = True
        End If
            MSHFlexGrid1_EnterCell
    End With
End Sub

Private Sub MSHFlexGrid1_EnterCell()
    Dim nCtr As Integer, _
    aTimeInfo As Variant, _
    oRecordSet As New ADODB.Recordset
    

'    If (MSHFlexGrid1.Width <> 3675) Or (Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)) = "") Then Exit Sub
    If (MSHFlexGrid1.Width <> 4965) Then Exit Sub

    With MSHFlexGrid1
        Text3(0).Text = .TextMatrix(.RowSel, 2)
        Text3(1).Text = .TextMatrix(.RowSel, 1)
        Text3(2).Text = .TextMatrix(.RowSel, 3)
        Text3(3).Text = .TextMatrix(.RowSel, 8)
        Text3(4).Text = .TextMatrix(.RowSel, 4)
        Text3(5).Text = IIf(Trim(.TextMatrix(.RowSel, 5)) = "", "", IIf(Val(.TextMatrix(.RowSel, 5)) = 0, "WAP", IIf(Val(.TextMatrix(.RowSel, 5)) = 1, "Contractual", IIf(Val(.TextMatrix(.RowSel, 5)) = 2, "Regular", "Tesda"))))
        Text3(6).Text = IIf(Trim(.TextMatrix(.RowSel, 22)) = "", "", IIf(Val(.TextMatrix(.RowSel, 22)) = 0, "", "WAP-C"))
        Label27.Caption = IIf(Val(.TextMatrix(.RowSel, 10)) = 0, "", IIf(Val(.TextMatrix(.RowSel, 10)) = 1, "Resigned", IIf(Val(.TextMatrix(.RowSel, 10)) = 2, "Finished Contract", "Terminated")))
    End With
    
    With MSHFlexGrid2
    
        ShowProgress 0

        .Redraw = False
        
        DoEvents
'--------------------- (2017-03-02) Updated
        For nCtr = 0 To DateDiff("d", XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate)
            If XPDatePicker1.CurrentDate = XPDatePicker2.CurrentDate Then
                ShowProgress 2, nCtr = 1 * 100
            Else
            
            ShowProgress 2, (nCtr / DateDiff("d", XPDatePicker1.CurrentDate, XPDatePicker2.CurrentDate)) * 100
            End If
'---------------------
            ' --> retrieve assigned shift for the day...
            OpenQueryDNS "select a.shiftid, " & _
                         " a.reg_hr, a.reg_ot_hr, a.sa_reg_ot, a.tot_ot, " & _
                         " a.nd_hr, a.nd_ot_hr, a.sa_nd_ot, a.nd_tot_ot, " & _
                         " a.sun_hr, a.sun_ot_hr, " & _
                         " a.sun_nd, a.sun_nd_ot, " & _
                         " a.remark, a.inc_hr " & _
                         "from " & IIf(Val(lblPClose.Caption) = 0, "di36770", "dih36770") & " a where a.empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                         " and a.date=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote, objdbRs, False
'            MsgBox "select * from di36770 where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & " and di36770.date=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote

            If objdbRs.RecordCount > 0 Then
            
                OpenQueryDNS " select shiftid " & _
                             " from " & IIf(Val(lblPClose.Caption) = 0, "pa84650 ", "pah84650 ") & _
                             " where empid=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2) & cQuote & _
                             " and logdate=" & cQuote & Format(DateAdd("d", nCtr, XPDatePicker1.CurrentDate), "yyyy-mm-dd") & cQuote & _
                             " and trantype = 0 group by empid", oRecordSet, False
                If oRecordSet.RecordCount > 0 Then
                    If oRecordSet("shiftid") <> objdbRs("shiftid") Then
                        .TextMatrix(nCtr + 1, 16) = oRecordSet("shiftid")
                    Else
                        .TextMatrix(nCtr + 1, 16) = objdbRs("shiftid")
                    End If
                Else
                    .TextMatrix(nCtr + 1, 16) = objdbRs("shiftid")
                End If
                
                .TextMatrix(nCtr + 1, 3) = IIf(objdbRs("reg_hr") > 0, objdbRs("reg_hr"), "")
                .TextMatrix(nCtr + 1, 4) = IIf(objdbRs("reg_ot_hr") > 0, objdbRs("reg_ot_hr"), "")
                .TextMatrix(nCtr + 1, 5) = IIf(objdbRs("sa_reg_ot") > 0, objdbRs("sa_reg_ot"), "")
                .TextMatrix(nCtr + 1, 6) = IIf(objdbRs("tot_ot") > 0, objdbRs("tot_ot"), "")
                .TextMatrix(nCtr + 1, 7) = IIf(objdbRs("nd_hr") > 0, objdbRs("nd_hr"), "")
                .TextMatrix(nCtr + 1, 8) = IIf(objdbRs("nd_ot_hr") > 0, objdbRs("nd_ot_hr"), "")
                .TextMatrix(nCtr + 1, 9) = IIf(objdbRs("sa_nd_ot") > 0, objdbRs("sa_nd_ot"), "")
                .TextMatrix(nCtr + 1, 10) = IIf(objdbRs("nd_tot_ot") > 0, objdbRs("nd_tot_ot"), "")
                .TextMatrix(nCtr + 1, 11) = IIf(objdbRs("sun_hr") > 0, objdbRs("sun_hr"), "")
                .TextMatrix(nCtr + 1, 12) = IIf(objdbRs("sun_ot_hr") > 0, objdbRs("sun_ot_hr"), "")
                .TextMatrix(nCtr + 1, 13) = IIf(objdbRs("sun_nd") > 0, objdbRs("sun_nd"), "")
                .TextMatrix(nCtr + 1, 14) = IIf(objdbRs("sun_nd_ot") > 0, objdbRs("sun_nd_ot"), "")
                .TextMatrix(nCtr + 1, 15) = objdbRs("remark")
                .TextMatrix(nCtr + 1, 18) = IIf(objdbRs("inc_hr") > 0, objdbRs("inc_hr"), "")
                
            Else
                .TextMatrix(nCtr + 1, 16) = ""
                
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
                .TextMatrix(nCtr + 1, 14) = ""
                .TextMatrix(nCtr + 1, 15) = ""
                .TextMatrix(nCtr + 1, 18) = ""
                
            End If
            
            HiLyt2 nCtr + 1, MSHFlexGrid2, IIf(Trim(.TextMatrix(nCtr + 1, 15)) = "", vbBlack, vbBlue)
                
            If .TextMatrix(nCtr + 1, 18) <> "" Then
                HiLyt2 nCtr + 1, MSHFlexGrid2, vbMagenta
            End If
            
            If Weekday(DateAdd("d", nCtr, XPDatePicker1.CurrentDate)) = vbSunday Then
                HiLyt2 nCtr + 1, MSHFlexGrid2, vbRed
            End If
            
        Next nCtr
        
        .Redraw = True
        ShowProgress 4
        
    End With
    
    MSHFlexGrid2_EnterCell
    
    MSHFlexGrid1.RowSel = MSHFlexGrid1.Row
    
    Set oRecordSet = Nothing
End Sub

Private Sub MSHFlexGrid2_EnterCell()
    Dim cSqlStmt As String
    
    If MSHFlexGrid2.Cols > 2 Then
        Text2.Text = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 16)
        cSqlStmt = "select * from PA74380 where shiftid=" & cQuote & MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 16) & cQuote
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
'    Script2File cSqlStmt
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid3, myArray3, False, , True, 1
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
                        
                        RefreshGrid MSHFlexGrid3, True
                        
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
                            RefreshGrid MSHFlexGrid3, True
                        End If
                    Else
                        If (Trim(.TextMatrix(.RowSel, 2)) = "") Or (Trim(.TextMatrix(.Rows - 1, 6)) = "") Then
                            .RemoveItem .RowSel
                            RefreshGrid MSHFlexGrid3, True
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text1.Text) = "" Then
            Command1_Click
        Else
            OpenQueryDNS "SELECT * FROM PA7730 where periodid=" & cQuote & Text1.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label1.Caption = objdbRs("duration")
                lblPClose.Caption = objdbRs("pclose")
                XPDatePicker1.CurrentDate = objdbRs("date_start")
                XPDatePicker2.CurrentDate = objdbRs("date_end")
            Else
                Label1.Caption = ""
                lblPClose.Caption = ""
                XPDatePicker1.CurrentDate = Now
                XPDatePicker2.CurrentDate = Now
            End If
            ChkPeriod
        End If
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
            nPos = 6
        Case 4
            nPos = 7
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

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Text2.Text) = "" Then
            Command9_Click
        Else
            OpenQueryDNS "SELECT * FROM PA74380 WHERE SHIFTID=" & cQuote & Text2.Text & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label6.Caption = EncodeStr(objdbRs("description"))
                Label7.Caption = Format(objdbRs("time1"), "hh:mm AMPM")
                Label8.Caption = Format(objdbRs("time2"), "hh:mm AMPM")
            Else
                Label6.Caption = ""
                Label7.Caption = ""
                Label8.Caption = ""
                MsgBox "Shift ID not found!", vbCritical, App.Title
                Text2.SetFocus
            End If
        End If
    End If
End Sub

Private Sub XPDatePicker2_Validate(Cancel As Boolean)
    If XPDatePicker2.CurrentDate < XPDatePicker1.CurrentDate Then
        MsgBox "Start Date should be greater than or equal to End Date", vbCritical, "System Advisory"
        Cancel = True
    End If
End Sub
