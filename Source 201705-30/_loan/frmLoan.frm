VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaxppanel.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaxpframe.ocx"
Begin VB.Form frmLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Loan Entry"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12720
   Begin VB.CheckBox Check3 
      Caption         =   "Only active Employee"
      Height          =   270
      Left            =   3840
      TabIndex        =   61
      Top             =   90
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.TextBox Text14 
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
      Left            =   1740
      TabIndex        =   58
      Top             =   690
      Width           =   4725
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&New"
      Height          =   330
      Left            =   2970
      TabIndex        =   46
      Top             =   60
      Width           =   690
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   330
      Left            =   2490
      TabIndex        =   45
      Top             =   60
      Width           =   450
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   5730
      Left            =   0
      TabIndex        =   28
      Top             =   1320
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10107
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmLoan.frx":0000
         Left            =   1740
         List            =   "frmLoan.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "1"
         ToolTipText     =   "NUM:STATUS"
         Top             =   5355
         Width           =   1395
      End
      Begin ciaXPFrame.XPFrame XPFrame1 
         Height          =   885
         Left            =   5865
         TabIndex        =   56
         Top             =   3135
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1561
         Caption         =   " Period Covered "
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
         Radius          =   20
         LicValid        =   -1  'True
         Begin VB.CheckBox Check2 
            Caption         =   "Period 2 (16-end)"
            Height          =   345
            Left            =   135
            TabIndex        =   5
            Tag             =   "1"
            ToolTipText     =   "NUM:PERIOD2"
            Top             =   450
            Width           =   1620
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Period 1 (1-15)"
            Height          =   345
            Left            =   135
            TabIndex        =   4
            Tag             =   "1"
            ToolTipText     =   "NUM:PERIOD1"
            Top             =   180
            Width           =   1620
         End
      End
      Begin VB.TextBox Text13 
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
         Left            =   3135
         TabIndex        =   54
         Top             =   5355
         Width           =   4575
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10815
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   4950
         Width           =   1515
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
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
         Left            =   10815
         TabIndex        =   50
         Tag             =   "1"
         ToolTipText     =   "NUM:ACC_AMT"
         Top             =   4650
         Width           =   1515
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   4305
         TabIndex        =   49
         Tag             =   "19"
         Top             =   2790
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   3540
         TabIndex        =   48
         Tag             =   "18"
         Top             =   2790
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   2280
         TabIndex        =   44
         Top             =   3120
         Width           =   450
      End
      Begin VB.TextBox Text10 
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
         Left            =   1740
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "TXT:REMARK"
         Top             =   4695
         Width           =   5970
      End
      Begin VB.TextBox Text9 
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
         Left            =   1740
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "TXT:REF_NO"
         Top             =   4395
         Width           =   2115
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   1740
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "NUM:DEF_AMT"
         Top             =   3735
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   1740
         TabIndex        =   8
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_GRANT"
         Top             =   4095
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   503
         _Version        =   393216
         Format          =   245104640
         CurrentDate     =   38623
      End
      Begin ciaXPPanel.XPPanel XPPanel2 
         Height          =   990
         Left            =   1755
         TabIndex        =   37
         Top             =   4035
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   1746
         HasLeftBorder   =   0   'False
         HasRightBorder  =   0   'False
         LicValid        =   -1  'True
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   1740
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "NUM:CUT_OFF_AMT"
         Top             =   3435
         Width           =   1185
      End
      Begin VB.TextBox Text6 
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
         Left            =   1740
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "TXT:DEDID"
         Top             =   3135
         Width           =   510
      End
      Begin VB.TextBox Text5 
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
         Left            =   1740
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "TXT:CTRL_NO"
         Top             =   2835
         Width           =   990
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2715
         Left            =   75
         TabIndex        =   29
         Top             =   60
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   4789
         _Version        =   393216
         GridColor       =   12640511
         GridColorUnpopulated=   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
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
         Height          =   4260
         Left            =   7800
         TabIndex        =   31
         Top             =   375
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   7514
         _Version        =   393216
         GridColor       =   12640511
         GridColorUnpopulated=   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1740
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_START"
         Top             =   5055
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Format          =   245104641
         CurrentDate     =   38623
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3645
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_END"
         Top             =   5055
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Format          =   245104641
         CurrentDate     =   38623
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Add"
         Height          =   345
         Left            =   2775
         TabIndex        =   47
         Tag             =   "17"
         Top             =   2790
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   5340
         TabIndex        =   60
         Tag             =   "1"
         ToolTipText     =   "DAT:DATE_FIN"
         Top             =   5055
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Format          =   245104641
         CurrentDate     =   38623
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE: Cell in BLUE foreground color are Active Payroll and are not yet included in the Total Deduction."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   7815
         TabIndex        =   57
         Top             =   5235
         Width           =   4785
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   180
         TabIndex        =   55
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Left            =   9285
         TabIndex        =   53
         Top             =   4980
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deduction"
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
         Left            =   9285
         TabIndex        =   51
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3465
         TabIndex        =   43
         Top             =   5085
         Width           =   90
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
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
         Left            =   195
         TabIndex        =   42
         Top             =   5070
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remark"
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
         Left            =   180
         TabIndex        =   41
         Top             =   4725
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Number"
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
         Left            =   180
         TabIndex        =   40
         Top             =   4425
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amortization"
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
         Left            =   180
         TabIndex        =   39
         Top             =   3795
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Granted"
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
         Left            =   180
         TabIndex        =   38
         Top             =   4110
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount"
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
         Left            =   180
         TabIndex        =   36
         Top             =   3495
         Width           =   1455
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2835
         TabIndex        =   35
         Top             =   3180
         Width           =   4530
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction"
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
         Left            =   180
         TabIndex        =   34
         Top             =   3195
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Control Number"
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
         Left            =   180
         TabIndex        =   33
         Top             =   2895
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   2940
         Left            =   15
         Top             =   2790
         Width           =   1710
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Access Rights"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7800
         TabIndex        =   32
         Top             =   105
         Width           =   4845
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7800
         Top             =   60
         Width           =   4845
      End
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   5160
      TabIndex        =   30
      Top             =   6990
      Width           =   7500
      Begin VB.CommandButton Command7 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   4560
         Picture         =   "frmLoan.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   5550
         Picture         =   "frmLoan.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bottom"
         Height          =   660
         Index           =   3
         Left            =   2625
         Picture         =   "frmLoan.frx":3324
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "12"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   660
         Index           =   2
         Left            =   1785
         Picture         =   "frmLoan.frx":4CA6
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "14"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   660
         Index           =   1
         Left            =   945
         Picture         =   "frmLoan.frx":6628
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "13"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   6510
         Picture         =   "frmLoan.frx":7FAA
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Searc&h"
         Height          =   660
         Left            =   3585
         Picture         =   "frmLoan.frx":992C
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Top"
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "frmLoan.frx":B2AE
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "11"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
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
      Left            =   8115
      TabIndex        =   26
      Top             =   390
      Visible         =   0   'False
      Width           =   1515
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
      Left            =   1740
      TabIndex        =   24
      Top             =   990
      Width           =   4725
   End
   Begin VB.TextBox Text2 
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
      Left            =   1740
      TabIndex        =   0
      Top             =   390
      Width           =   4725
   End
   Begin VB.TextBox Text1 
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
      Left            =   1740
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "TXT:EMPID"
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label19 
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
      Left            =   180
      TabIndex        =   59
      Top             =   735
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   6555
      TabIndex        =   27
      Top             =   435
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   180
      TabIndex        =   23
      Top             =   435
      Width           =   1455
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
      Left            =   180
      TabIndex        =   22
      Top             =   135
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   8355
      Left            =   0
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'               $`````$
'             $( o  o )$
'    >------oOO--(_)--OOo------------------------------------------------------------------------------<
'    "Intelligent people can be bored. They just know a lot. But Smart people
'    are never bored, because they're always looking for something to engage
'    their minds."
'    >------oooo(O) (0)oooo----------------------------------------------------------------------------<

' project name  :   Dong-in Payroll & Time Management System
' module        :   frmLoan
' description   :   Employee Loan Entry/Monitoring Module
' programmer    :   _-=[ srm ]=-_
' date started  :   1 aug 2006

Option Explicit
    Dim lNewEmp As Boolean, _
        nAdd As Integer, _
        cSeries As String, _
        oTempADO As New ADODB.Recordset, _
        oRecordSet As New ADODB.Recordset, _
        myArray As Variant, _
        myArray2 As Variant
        
Sub ShowData(ByVal cString As String, oLabel As Label, nMode As Integer)
    Dim cSqlStmt As String
    If nMode = 1 Then
        cSqlStmt = "select dedname, period1, period2 from pa3330 where dedid=" & cQuote & cString & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            oLabel.Caption = objdbRs("dedname")
            Label5.Caption = "Amortization for " & oLabel.Caption
            Check1.Value = IIf(objdbRs("period1") = 1, vbChecked, vbUnchecked)
            Check2.Value = IIf(objdbRs("period2") = 1, vbChecked, vbUnchecked)
        Else
            oLabel.Caption = ""
            Label5.Caption = ""
            Check1.Value = vbUnchecked
            Check2.Value = vbUnchecked
        End If
    Else
        cSqlStmt = "select a.empid, " & _
                   " concat(a.firstname,' ',if(trim(a.mname)='','',concat(left(a.mname,1),'. ')),a.lastname) as fullname, " & _
                   " ifnull(b.linename,'') as department " & _
                   "from di3670 a left join di5463 b on a.depid=b.lineid " & _
                   "where empid=" & cQuote & cString & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        Text2.Text = IIf(objdbRs.RecordCount > 0, objdbRs("fullname"), "")
        Text3.Text = IIf(objdbRs.RecordCount > 0, objdbRs("department"), "")
    End If

End Sub

Sub ShowRecords()
    Dim cSqlStmt As String
    
    If oTempADO.RecordCount > 0 Then
        If oTempADO.EOF Then oTempADO.MoveFirst
        
        Text1.Text = oTempADO("empid")
        Text2.Text = oTempADO("fullname")
        Text3.Text = oTempADO("linename")

        cSqlStmt = "select a.ctrl_no, " & _
                   " a.dedid, " & _
                   " ifnull(b.dedname,'') as dedname, " & _
                   " a.status, " & _
                   " if(a.status=1,concat('Loan settled as of ',date_format(a.date_fin,'%M %d, %Y')),'') as remark " & _
                   "from di3673 a left join pa3330 b on a.dedid=b.dedid " & _
                   "where a.empid=" & cQuote & Text1.Text & cQuote
                   
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            QueryAttach objdbRs, MSHFlexGrid1, myArray
        Else
            SetGridColumn myArray, MSHFlexGrid1
        End If
        
        MSHFlexGrid1_EnterCell
    Else
        SetGridColumn myArray, MSHFlexGrid1
        SetGridColumn myArray2, MSHFlexGrid2
    End If
End Sub

Private Sub Check3_Click()
    Dim cSqlStmt As String
    cSqlStmt = "select distinct a.empid, " & _
               "  ifnull(concat(b.firstname,' ',if(trim(b.mname)='','',concat(left(b.mname,1),'. ')),b.lastname),'') as fullname, " & _
               "  ifnull(c.linename,'') as linename, " & _
               "  ifnull(if(b.active>0,concat(if(b.active=1,'Resigned ','Finish Contract'),' as of ',date_format(if(b.active=1,b.date_res,b.date_fin),'%b %d, %Y')),if(b.emp_stat=0,'WAP',if(b.emp_stat=1,'Contractual','Regular'))),'') as status " & _
               "from di3673 a left join di3670 b on a.empid=b.empid " & _
               "  left join di5463 c on b.depid=c.lineid " & _
               IIf(Check3.Value = vbChecked, " Where b.active = 0 ", "") & _
               "order by fullname "
               
    OpenQueryDNS cSqlStmt, oTempADO, False
End Sub

Private Sub Combo4_Click()
    If nAdd = 0 Then Exit Sub
    If Combo4.ListIndex = 1 Then DTPicker4.Value = Now
End Sub

Private Sub Command1_Click(Index As Integer)
    If TypeOf Screen.ActiveControl Is CommandButton Then
'        dbNavigator Screen.ActiveControl, Me, oTempADO
        If oTempADO.RecordCount > 0 Then
            Select Case Command1(Index).Tag
                Case 11     ' --> top
                    oTempADO.MoveFirst
                Case 12     ' --> bottom
                    oTempADO.MoveLast
                Case 13     ' --> previous
                    oTempADO.MovePrevious
                    If oTempADO.BOF Then oTempADO.MoveFirst
                Case 14     ' --> next
                    oTempADO.MoveNext
                    If oTempADO.EOF Then oTempADO.MoveLast
            End Select
        End If
        ShowRecords
    End If
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrSaveLoan
    Dim cString, _
        cSqlStmt, _
        cCondition As String, _
        nCtr As Integer
    
    If nAdd = 0 Then
        MsgBox "Please insert a loan entry first!", vbCritical, "System Advisory!!!"
        GoTo endsave
    End If
    
    cString = Text1.Text
    
    Select Case MsgBox(IIf(nAdd = 1, "Save", "Update") & " Loan entry?", vbYesNoCancel, "Loan Monitoring Entry...")
        Case vbYes
            If nAdd = 1 Then
                If IfExists("di3673", "CTRL_NO=" & cQuote & Text5.Text & cQuote) Then
                    MsgBox "Loan Control Number is already existing!", vbCritical, "System Advisory!"
                    Text5.SetFocus
                    GoTo endsave
                Else
                    OpenQueryDNS InsertFields(Me, "DI3673"), objdbRs, True
                    Script2File InsertFields(Me, "DI3673")
                    Log2Audit Name, "Add Loan Control #" & Text5.Text & " to Employee ID#" & Text1.Text
                End If
            Else
                cCondition = "(EMPID=" & cQuote & Text1.Text & cQuote & ") AND " & _
                             "(DEDID=" & cQuote & Text6.Text & cQuote & ") AND " & _
                             "(CTRL_NO=" & cQuote & IIf(oRecordSet("CTRL_NO") <> Text5.Text, "", Text5.Text) & cQuote & ")"
                OpenQueryDNS EditField(Me, "di3673", cCondition), objdbRs, True
                Script2File EditField(Me, "di3673", cCondition)
                
                Log2Audit Name, "Update Loan Control #" & Text5.Text & " to Employee ID#" & Text1.Text
                
                If oRecordSet("CTRL_NO") <> Text5.Text Then
                    With MSHFlexGrid2
                        For nCtr = 1 To .Rows - 1
                            If Val(.TextMatrix(nCtr, 4)) = 1 Then
                                cSqlStmt = "update " & IIf(Val(.TextMatrix(nCtr, 5)) = 1, "pah87263", "pa87263") & _
                                           " set ctrl_no=" & cQuote & Text5.Text & cQuote & _
                                           " where " & cCondition & " and periodid=" & cQuote & .TextMatrix(nCtr, 3) & cQuote
                                OpenQueryDNS cSqlStmt, objdbRs, True
                                Script2File cSqlStmt
                            End If
                        Next nCtr
                    End With
                End If
            End If
            
        Case vbNo
            cString = ""
            
        Case vbCancel
            GoTo endsave
    End Select

    If lNewEmp Then lNewEmp = False
    
    Lock2User Me.Name, Text5.ToolTipText, Text5.Text, False     ' --> 20050321

    If Text5.Text <> cSeries Then ResetSeries "LOAN", cSeries

    nAdd = 0
    CtrlPanel Me, nAdd
    ClearAll Me, False, True

    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True

    MSHFlexGrid1.Enabled = True

    oTempADO.Requery adAsyncFetch
    If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
    ShowRecords

endsave:
    Exit Sub
    
ErrSaveLoan:
    ErrorMsg Err.Number, Err.Description, "Save Button", Name
End Sub

Private Sub Command11_Click()
    Dim cString As String
    
    cString = Text1.Text
    If nAdd = 0 Then
        If lNewEmp Then
            lNewEmp = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = True
            
            MSHFlexGrid1.Enabled = True
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
        
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            ShowRecords
        Else
            Unload Me
        End If
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, Text5.ToolTipText, Text5.Text, False     ' --> 20050321
            
            If Text5.Text <> cSeries Then ResetSeries "LOAN", cSeries
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = True
            
            MSHFlexGrid1.Enabled = True
            
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
'            GetFields Me, oTempADO
            ShowRecords
            
            If lNewEmp Then
                Command4_Click
                Text1.Text = cString
                ShowData cString, Label27, 2
            End If
        End If
    End If
End Sub

Private Sub Command12_Click()
    If Not isDataLock(Me.Name, Text5.ToolTipText, Text5.Text) Then
        
        If Trim(Text5.Text) = "" Then
            cSeries = GenerateSeries("LOAN")
            Text5.Text = PadStr(cSeries, "0", Text5.MaxLength)
            While IfExists("DI3673", "DI3673.CTRL_NO=" & cQuote & PadStr(cSeries, "0", Text5.MaxLength) & cQuote)
                cSeries = GenerateSeries("LOAN")
                Text5.Text = PadStr(cSeries, "0", Text5.MaxLength)
            Wend
        End If
        
        Lock2User Me.Name, Text5.ToolTipText, Text5.Text, True
        nAdd = 2
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        
        Text1.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
    
        MSHFlexGrid1.Enabled = False
        
        Text7.SetFocus
    End If
End Sub

' --> activated 20070118
Private Sub Command13_Click()
    Dim cSqlStmt As String, cString As String
    
    cString = Text1.Text
    
    If (MSHFlexGrid2.Rows - 1) = 1 Then
        If MsgBox("Delete loan entry for " & Trim(Text2.Text) & "?", vbYesNo, "System Advisory!!!") = vbYes Then
            cSqlStmt = "delete from di3673 where ctrl_no=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1) & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, True
            Script2File cSqlStmt
        
            oTempADO.Requery adAsyncFetch
            If Trim(cString) <> "" Then oTempADO.Find "EMPID='" & PadStr(cString, " ", Text1.MaxLength, PadRight) & "'"
            ShowRecords
        End If
    Else
        MsgBox "You are not allowed to delete active loan!!!", vbCritical, "System Advisory!!!"
    End If
End Sub

Private Sub Command2_Click()
    Dim cParam As String, _
        nCtr As Integer
    
    With MSHFlexGrid1
        For nCtr = 1 To (.Rows - 1)
            If (Trim(.TextMatrix(nCtr, 2)) <> "") And (Val(.TextMatrix(nCtr, 4)) = 0) Then
                cParam = cParam & cQuote & .TextMatrix(nCtr, 2) & cQuote & ","
            End If
        Next nCtr
    End With
    
    If Trim(cParam) <> "" Then cParam = " AND DEDID NOT IN (" & left(cParam, Len(cParam) - 1) & ")"
    
    frmLookup.showPopup 7, "WHERE DEDTYPE=1 " & cParam
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text6.Text = cResult
        ShowData cResult, Label27, 1
        Text6.SetFocus
    End If
    
End Sub

Private Sub Command3_Click()
    frmLookup.showPopup 3, "where (a.empid not in (select distinct empid from di3673))" & IIf(Check3.Value = vbChecked, " and a.active = 0 ", "")
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        Text1.Text = cResult
        ShowData cResult, Label27, 2
        Text1.SetFocus
    End If
End Sub
Private Sub Command4_Click()
    lNewEmp = True
    
    ClearAll Me, True, True
    CtrlPanel Me, 1
    
    Command3.Enabled = True
    Command6.Enabled = Mid(Tag, 2, 1) = "1"
    Command12.Enabled = Mid(Tag, 3, 1) = "1"
    Command13.Enabled = Mid(Tag, 4, 1) = "1"
    
    Label5.Caption = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    
    ShowData Text6.Text, Label27, 1
    
    MSHFlexGrid1.Enabled = False
    SetGridColumn myArray, MSHFlexGrid1
    SetGridColumn myArray2, MSHFlexGrid2
    
    Text1.SetFocus
End Sub

Private Sub Command5_Click()
    Log2Audit Name, "SEARCH"
    frmLookup.showPopup 13, IIf(Check3.Value = vbChecked, " where active = 0 ", "")
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        oTempADO.Requery adAsyncFetch
        oTempADO.Find "EMPID='" & PadStr(cResult, " ", Text1.MaxLength, PadRight) & "'"
        If Not oTempADO.EOF Then
'            GetFields Me, oTempADO
            ShowRecords
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim cString As String
    
    cString = Text1.Text
    
    nAdd = 1
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Text1.Text = cString
    Text1.Enabled = False
    Text12.Text = ""
    Text13.Text = ""
    
    Command2.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = False
    
    Label5.Caption = ""
    Label27.Caption = ""
    
    MSHFlexGrid1.Enabled = False
    SetGridColumn myArray2, MSHFlexGrid2
    
    cSeries = GenerateSeries("LOAN")
    Text5.Text = PadStr(cSeries, "0", Text5.MaxLength)
    While IfExists("DI3673", "DI3673.CTRL_NO=" & cQuote & PadStr(cSeries, "0", Text5.MaxLength) & cQuote)
        cSeries = GenerateSeries("LOAN")
        Text5.Text = PadStr(cSeries, "0", Text5.MaxLength)
    Wend
    Text5.SetFocus
    
End Sub

Sub CreateTemp()
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    cSqlStmt = " CREATE TABLE tmpLoanDet( " & _
               " [EMPID] char(6),            [FULLNAME] char(100), " & _
               " [POSNAME] char(100),        [LineName] char(100), " & _
               " [DEDNAME] char(100),        [DEF_AMT] double, " & _
               " [CUT_OFF_AMT] double,       [PERIOD1] char(10), " & _
               " [PERIOD2] char(10),         [CTRL_NO] char(10), " & _
               " [DURATION] char(100), " & _
               " [DATE_GRANT] date,          [deduction] char(100), " & _
               " [REMARKH] char(100),        [REMARKD] char(100), " & _
               " [StatusName] char(100),     [PayName] char(100), " & _
               " [Amount] double,            [CMPName] char(100), " & _
               " [REF_NO] char(100),         [SEQ_NO] integer )"

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM tmpLoanDet", oTempADO, True
End Sub

Private Sub Command7_Click()
    Dim cSqlStmt As String, _
        cCmpname As String, _
        cDuration As String, _
        nCtr As Integer, _
        oRecordSet As New ADODB.Recordset

    CreateTemp
    
    ShowProgress 0
    
    OpenQueryDNS " select cmpname from di2660 where cmpid = " & gCompanyID, objdbRs, False
    cCmpname = IIf(objdbRs.RecordCount > 0, objdbRs("cmpname"), "")
    
    With MSHFlexGrid2
        
        For nCtr = 1 To .Rows - 1
            
            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
            
                ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            
                'cDuration = IIf(Combo4.ListIndex <> 0, Format(DTPicker1, "yyyy-mm-dd") & " - " & Format(DTPicker4, "yyyy-mm-dd"), Format(DTPicker1, "yyyy-mm-dd") & " - " & Format(DTPicker2, "yyyy-mm-dd"))
                cDuration = ""
                cDuration = IIf(Combo4.ListIndex <> 0, MonthName(Month(DTPicker1.Value)) & " " & Day(DTPicker1.Value) & ", " & Year(DTPicker1.Value) & " - " & _
                            MonthName(Month(DTPicker4.Value)) & " " & Day(DTPicker4.Value) & ", " & Year(DTPicker4.Value), MonthName(Month(DTPicker1.Value)) & " " & Day(DTPicker1.Value) & ", " & Year(DTPicker1.Value) & " - " & _
                            MonthName(Month(DTPicker2.Value)) & " " & Day(DTPicker2.Value) & ", " & Year(DTPicker2.Value))
            
                cSqlStmt = " INSERT INTO tmpLoanDet(EMPID, FULLNAME, POSNAME, LineName, DEDNAME, DEF_AMT, CUT_OFF_AMT, PERIOD1, PERIOD2," & _
                           " CTRL_NO, DATE_GRANT, DURATION, deduction, REMARKH, REMARKD, StatusName, PayName, Amount,REF_NO,CMPName,SEQ_NO)VALUES(" & _
                           cQuote & Text1.Text & cQuote & "," & _
                           cQuote & Text2.Text & cQuote & "," & _
                           cQuote & Text14.Text & cQuote & "," & _
                           cQuote & Text3.Text & cQuote & "," & _
                           cQuote & Label27.Caption & cQuote & "," & _
                           cQuote & Text8.Text & cQuote & "," & _
                           Val(Format(Text7.Text, "###0.00")) & "," & _
                           cQuote & IIf(Check1.Value = vbChecked, "Yes", "No") & cQuote & "," & _
                           cQuote & IIf(Check2.Value = vbChecked, "Yes", "No") & cQuote & "," & _
                           cQuote & Text5.Text & cQuote & "," & _
                           cQuote & Format(DTPicker3.Value, "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & cDuration & cQuote & "," & _
                           cQuote & Text13.Text & cQuote & "," & _
                           cQuote & Text10.Text & cQuote & "," & _
                           cQuote & MSHFlexGrid1.TextMatrix(.RowSel, 3) & cQuote & "," & _
                           cQuote & Combo4.Text & cQuote & "," & _
                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                           .TextMatrix(nCtr, 2) & "," & _
                           cQuote & Text9.Text & cQuote & "," & _
                           cQuote & cCmpname & cQuote & "," & _
                           nCtr & ")"

                MsgBox cSqlStmt
                QueryTemp cSqlStmt, oRecordSet, True
            End If
        Next
        
        ShowProgress 3

        GenerateReport "Employee Loan Preview", "prv3673.RPT", , True

        ShowProgress 4
       
    End With
    
    Set oRecordSet = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim cSqlStmt As String
    
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Ctrl No]:13:True", _
                    "TXT:[Ded ID]:3:False", _
                    "TXT:[Description]:30:True", _
                    "NUM:[Status]:1:False", _
                    "TXT:[Remark]:40:True")
    myArray2 = Array("TXT:[Payroll Period]:30:True", _
                     "NUM:[Amount]:18.4:True", _
                     "TXT:[PERIODID]:5:False", _
                     "NUM:[Tag]:1:False", _
                     "NUM:[Old]:1:False", _
                     "TXT:[Date]:10:False")

    Tag = nAccess_Tag
    nAdd = 0
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
            
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True
        
    cSqlStmt = "select distinct a.empid, " & _
               "  ifnull(concat(b.firstname,' ',if(trim(b.mname)='','',concat(left(b.mname,1),'. ')),b.lastname),'') as fullname, " & _
               "  ifnull(c.linename,'') as linename, " & _
               "  ifnull(if(b.active>0,concat(if(b.active=1,'Resigned ','Finish Contract'),' as of ',date_format(if(b.active=1,b.date_res,b.date_fin),'%b %d, %Y')),if(b.emp_stat=0,'WAP',if(b.emp_stat=1,'Contractual','Regular'))),'') as status " & _
               "from di3673 a left join di3670 b on a.empid=b.empid " & _
               "  left join di5463 c on b.depid=c.lineid " & _
               IIf(Check3.Value = vbChecked, " Where b.active = 0 ", "") & _
               "order by fullname "
               
    OpenQueryDNS cSqlStmt, oTempADO, False
'    GetFields Me, oTempADO
    
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

Private Sub MSHFlexGrid1_EnterCell()
    Dim cSqlStmt As String, _
        nCtr As Integer
    
    With MSHFlexGrid1
        cSqlStmt = "select * from di3673 " & _
                   "where ctrl_no=" & cQuote & .TextMatrix(.RowSel, 1) & cQuote & _
                   " and empid=" & cQuote & Text1.Text & cQuote & _
                   " and dedid=" & cQuote & .TextMatrix(.RowSel, 2) & cQuote
        OpenQueryDNS cSqlStmt, oRecordSet, False
        GetFields Me, oRecordSet
        
        CtrlPanel Me, nAdd, oRecordSet("status") <> 1
        
        Text12.Text = Format(Val(Format(Text7.Text, "###0.00")) - Val(Format(Text11.Text, "###0.00")), "#,##0.00")
        Text13.Text = IIf(.TextMatrix(.RowSel, 4) = 1, "Deduction finished as of " & Format(oRecordSet("date_fin"), "mmm dd, yyyy"), "")
        ShowData Text6.Text, Label27, 1
        
'        Label5.Caption = "Amortization for " & Trim(.TextMatrix(.RowSel, 3))
        
        cSqlStmt = "select concat(date_format(b.date_start,'%b %d'),' - ',date_format(b.date_end,'%d, %Y')) as date, " & _
                   " a.ded_amt, a.periodid, 0, 1, b.date_start " & _
                   "from pah87263 a left join pa7730 b on a.periodid=b.periodid " & _
                   "where a.ctrl_no=" & cQuote & .TextMatrix(.RowSel, 1) & cQuote & _
                   " and a.empid=" & cQuote & Text1.Text & cQuote & _
                   " and a.dedid=" & cQuote & .TextMatrix(.RowSel, 2) & cQuote & _
                   " union all " & _
                   "select concat(date_format(b.date_start,'%b %d'),' - ',date_format(b.date_end,'%d, %Y')) as date, " & _
                   " a.ded_amt, a.periodid, 0, 0, b.date_start " & _
                   "from pa87263 a left join pa7730 b on a.periodid=b.periodid " & _
                   "where a.ctrl_no=" & cQuote & .TextMatrix(.RowSel, 1) & cQuote & _
                   " and a.empid=" & cQuote & Text1.Text & cQuote & _
                   " and a.dedid=" & cQuote & .TextMatrix(.RowSel, 2) & cQuote
'        Script2File cSqlStmt
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            QueryAttach objdbRs, MSHFlexGrid2, myArray2, False
            
            DoEvents
            With MSHFlexGrid2
                For nCtr = 1 To (.Rows - 1)
                    If Val(.TextMatrix(nCtr, 5)) = 0 Then HiLyt2 nCtr, MSHFlexGrid2, vbBlue
                Next nCtr
            End With
        Else
            SetGridColumn myArray2, MSHFlexGrid2
        End If
    End With
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nCtr As Integer, _
        nTotal As Double
    
    If nAdd = 0 Then Exit Sub
    
    With MSHFlexGrid2
        If (Button = vbLeftButton) And ((Shift And vbCtrlMask) > 0) Then
            .TextMatrix(.RowSel, 4) = IIf(Val(.TextMatrix(.RowSel, 4)) = 0, 1, 0)
            HiLyt2 .RowSel, MSHFlexGrid2, IIf(Val(.TextMatrix(.RowSel, 4)) = 0, vbBlack, vbRed)
            
            DoEvents
            For nCtr = 1 To (.Rows - 1)
                If .TextMatrix(nCtr, 5) = 1 Then
                    nTotal = nTotal + IIf(Val(.TextMatrix(nCtr, 4)) = 1, Val(.TextMatrix(nCtr, 2)), 0)
                End If
            Next nCtr
            Text11.Text = Format(nTotal, "##00.##00")
        End If
    End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If nAdd = 0 Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Trim(Text1.Text) = "" Then
            Command3_Click
        Else
            ShowData Text1.Text, Label1, 2
            Text1.SetFocus
        End If
    End If
End Sub


Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cParam As String, _
        cSqlStmt As String, _
        nCtr As Integer
    
    If KeyCode = vbKeyReturn Then
        If Trim(Text6.Text) = "" Then
            Command2_Click
        Else
            With MSHFlexGrid1
                For nCtr = 1 To (.Rows - 1)
                    If (Trim(.TextMatrix(nCtr, 2)) <> "") And (Val(.TextMatrix(nCtr, 4)) = 0) Then
                        cParam = cParam & cQuote & .TextMatrix(nCtr, 2) & cQuote & ","
                    End If
                Next nCtr
            End With
        
            If Trim(cParam) <> "" Then cParam = " AND DEDID NOT IN (" & left(cParam, Len(cParam) - 1) & ")"
    
'            cParam = "WHERE DEDTYPE=1 " & cParam & " and dedid = " & cQuote & Text6.Text & cQuote
            cParam = "WHERE DEDTYPE=1 " & cParam
        
            cSqlStmt = "SELECT DEDID, DEDNAME FROM PA3330 " & cParam
            OpenQueryDNS cSqlStmt, objdbRs, False
            If objdbRs.RecordCount > 0 Then
'                Text6.Text = objdbRs("DEDID")
'                Label27.Caption = objdbRs("DEDNAME")
'                Text7.SetFocus
                
                ShowData Text6.Text, Label27, 1
                Text6.SetFocus
            Else
                Text6.Text = ""
                Label27.Caption = ""
                Text6.SetFocus
            End If
        End If
    End If
    
End Sub
