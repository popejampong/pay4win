VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{11931057-9334-4856-BDAF-C62B6B94B551}#1.1#0"; "ciaxppanel.ocx"
Object = "{DF5E40D4-CC15-4039-861D-5D824D450C09}#1.1#0"; "ciaxpframe.ocx"
Begin VB.Form frmTransaction 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9450
   ClientLeft      =   495
   ClientTop       =   525
   ClientWidth     =   12240
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "No Deduction"
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
      Height          =   255
      Left            =   8310
      TabIndex        =   99
      Top             =   7020
      Width           =   2670
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Add"
      Height          =   660
      Left            =   5355
      Picture         =   "frmTransaction.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "17"
      Top             =   8760
      Width           =   855
   End
   Begin ciaXPPanel.XPPanel XPPanel6 
      Height          =   765
      Left            =   90
      TabIndex        =   86
      Top             =   7905
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1349
      LicValid        =   -1  'True
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
         Left            =   675
         TabIndex        =   88
         Top             =   405
         Width           =   3315
      End
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
         ItemData        =   "frmTransaction.frx":1982
         Left            =   660
         List            =   "frmTransaction.frx":1992
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   75
         Width           =   2820
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Searc&h"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   90
         Top             =   450
         Width           =   1575
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
         Height          =   375
         Left            =   90
         TabIndex        =   89
         Top             =   75
         Width           =   1935
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel4 
      Height          =   1455
      Left            =   8325
      TabIndex        =   71
      Top             =   7275
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2566
      HasLeftBorder   =   0   'False
      HasRightBorder  =   0   'False
      LicValid        =   -1  'True
      Begin VB.TextBox Text22 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2100
         TabIndex        =   21
         Tag             =   "1"
         Text            =   "99,999.99"
         ToolTipText     =   "NUM:GROSS_PAY"
         Top             =   435
         Width           =   1590
      End
      Begin VB.TextBox Text3 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1140
         TabIndex        =   96
         Tag             =   "1"
         Text            =   "99,999.99"
         ToolTipText     =   "NUM:BASICPAY"
         Top             =   435
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox Text7 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2100
         TabIndex        =   23
         Tag             =   "1"
         Text            =   "99,999.99"
         ToolTipText     =   "NUM:SA_NET_PAY"
         Top             =   1095
         Width           =   1590
      End
      Begin VB.TextBox Text6 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2100
         TabIndex        =   20
         Tag             =   "1"
         Text            =   "99,999.99"
         ToolTipText     =   "NUM:DED_AMT"
         Top             =   75
         Width           =   1590
      End
      Begin VB.TextBox Text23 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2100
         TabIndex        =   22
         Tag             =   "1"
         Text            =   "99,999.99"
         ToolTipText     =   "NUM:NET_PAY"
         Top             =   735
         Width           =   1590
      End
      Begin VB.Line Line2 
         X1              =   135
         X2              =   3690
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "NET PAY (SA)"
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
         TabIndex        =   75
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   3690
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deductions"
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
         TabIndex        =   74
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "NET PAY"
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
         TabIndex        =   73
         Top             =   780
         Width           =   1800
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "GROSS PAY"
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
         TabIndex        =   72
         Top             =   480
         Width           =   1800
      End
   End
   Begin ciaXPPanel.XPPanel XPPanel1 
      Height          =   480
      Left            =   4245
      TabIndex        =   47
      Top             =   60
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   847
      LicValid        =   -1  'True
      Begin VB.Label lblDuration 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   3195
         TabIndex        =   50
         Top             =   105
         Width           =   3855
      End
      Begin VB.Label lblPeriod 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   465
         Left            =   900
         TabIndex        =   49
         Top             =   105
         Width           =   600
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Period ID"
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   120
         TabIndex        =   48
         Top             =   135
         Width           =   1155
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame1 
      Height          =   1530
      Left            =   4245
      TabIndex        =   30
      Top             =   540
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   2699
      Alignment       =   2
      Caption         =   " Employee Information "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
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
         Index           =   10
         Left            =   3870
         TabIndex        =   106
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Index           =   9
         Left            =   7065
         TabIndex        =   98
         Tag             =   "1"
         ToolTipText     =   "NUM:SUN_COLA"
         Top             =   1215
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox Text2 
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
         Index           =   8
         Left            =   6555
         TabIndex        =   97
         Tag             =   "1"
         ToolTipText     =   "NUM:COLA"
         Top             =   1155
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   300
         Left            =   2100
         TabIndex        =   95
         Top             =   195
         Width           =   420
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
         Index           =   7
         Left            =   4530
         TabIndex        =   93
         Tag             =   "1"
         ToolTipText     =   "NUM:PAYSTATUS"
         Top             =   375
         Visible         =   0   'False
         Width           =   990
      End
      Begin ciaXPPanel.XPPanel XPPanel3 
         Height          =   1170
         Left            =   5520
         TabIndex        =   64
         Top             =   240
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2064
         HasRightBorder  =   0   'False
         HasTopBorder    =   0   'False
         HasBottomBorder =   0   'False
         LicValid        =   -1  'True
         Begin VB.TextBox Text2 
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
            Index           =   5
            Left            =   1215
            TabIndex        =   70
            Tag             =   "1"
            ToolTipText     =   "NUM:POS_ALLOW"
            Top             =   435
            Width           =   990
         End
         Begin VB.TextBox Text2 
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
            Index           =   6
            Left            =   1215
            TabIndex        =   67
            Tag             =   "1"
            ToolTipText     =   "NUM:COLA_AMT"
            Top             =   735
            Width           =   990
         End
         Begin VB.TextBox Text2 
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
            Index           =   4
            Left            =   1215
            TabIndex        =   65
            Tag             =   "1"
            ToolTipText     =   "NUM:RATE_AMT"
            Top             =   135
            Width           =   990
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Pos Allowance"
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
            TabIndex        =   69
            Top             =   495
            Width           =   1035
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "C.O.L.A."
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
            TabIndex        =   68
            Top             =   795
            Width           =   915
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
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
            TabIndex        =   66
            Top             =   135
            Width           =   915
         End
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
         Index           =   3
         Left            =   1080
         TabIndex        =   62
         Top             =   1110
         Width           =   2385
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
         Index           =   2
         Left            =   1080
         TabIndex        =   36
         Top             =   810
         Width           =   3600
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
         Index           =   1
         Left            =   1080
         TabIndex        =   34
         Tag             =   "1"
         ToolTipText     =   "TXT:FULLNAME"
         Top             =   510
         Width           =   4365
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
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "TXT:EMPID"
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No"
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
         Left            =   2595
         TabIndex        =   107
         Top             =   270
         Width           =   1335
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
         Left            =   3570
         TabIndex        =   91
         Top             =   1170
         Width           =   1830
      End
      Begin VB.Label Label19 
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
         Left            =   120
         TabIndex        =   63
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   37
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label3 
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
         TabIndex        =   35
         Top             =   570
         Width           =   915
      End
      Begin VB.Label Label2 
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
         TabIndex        =   33
         Top             =   270
         Width           =   915
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame6 
      Height          =   2175
      Left            =   4245
      TabIndex        =   52
      Top             =   5760
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   3836
      Caption         =   " Others "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   2685
         TabIndex        =   105
         Tag             =   "1"
         ToolTipText     =   "NUM:INC_PAY"
         Top             =   1710
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Computed"
         Height          =   255
         Left            =   1530
         TabIndex        =   94
         Top             =   1440
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   2685
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "NUM:M13PAY"
         Top             =   1410
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   12
         Left            =   2685
         TabIndex        =   16
         Tag             =   "1"
         ToolTipText     =   "NUM:LEAVE_PAY"
         Top             =   1110
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   2685
         TabIndex        =   15
         Tag             =   "1"
         ToolTipText     =   "NUM:OTHER_PAY"
         Top             =   810
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   2685
         TabIndex        =   14
         Tag             =   "1"
         ToolTipText     =   "NUM:SA_ADJ_PAY"
         Top             =   510
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   2685
         TabIndex        =   13
         Tag             =   "1"
         ToolTipText     =   "NUM:ADJ_PAY"
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive Pay"
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
         TabIndex        =   104
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "13th Month Pay"
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
         TabIndex        =   92
         Top             =   1455
         Width           =   1800
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive Leave"
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
         TabIndex        =   56
         Top             =   1155
         Width           =   1800
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         TabIndex        =   55
         Top             =   855
         Width           =   1800
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "SA Adjustment"
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
         TabIndex        =   54
         Top             =   555
         Width           =   2520
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment"
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
         TabIndex        =   53
         Top             =   255
         Width           =   1800
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame5 
      Height          =   7200
      Left            =   105
      TabIndex        =   51
      Top             =   720
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   12700
      Alignment       =   2
      Caption         =   " Employee List "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6795
         Left            =   90
         TabIndex        =   28
         Top             =   225
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   11986
         _Version        =   393216
         RowHeightMin    =   285
         ForeColorFixed  =   8388608
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         GridColor       =   -2147483632
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
   Begin ciaXPFrame.XPFrame XPFrame4 
      Height          =   720
      Left            =   105
      TabIndex        =   44
      Top             =   15
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1270
      Alignment       =   2
      Caption         =   " Department Info "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   345
         Left            =   645
         TabIndex        =   45
         Top             =   210
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   1155
         TabIndex        =   46
         Top             =   285
         Width           =   2790
      End
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   4260
      TabIndex        =   43
      Top             =   7845
      Width           =   3990
      Begin VB.CommandButton Command11 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   660
         Left            =   2985
         Picture         =   "frmTransaction.frx":19C5
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "21"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Save"
         Height          =   660
         Left            =   2040
         Picture         =   "frmTransaction.frx":3347
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "20"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Edit"
         Height          =   660
         Left            =   1095
         Picture         =   "frmTransaction.frx":4CC9
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "18"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Preview"
         Height          =   660
         Left            =   150
         Picture         =   "frmTransaction.frx":664B
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "15"
         Top             =   150
         Width           =   855
      End
   End
   Begin ciaXPFrame.XPFrame XPFrame3 
      Height          =   4980
      Left            =   8310
      TabIndex        =   32
      Top             =   2040
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8784
      Alignment       =   2
      Caption         =   " Deductions "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Radius          =   20
      LicValid        =   -1  'True
      Begin VB.TextBox txtFlex 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   720
         TabIndex        =   19
         Text            =   "Text3"
         Top             =   1410
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4575
         Left            =   105
         TabIndex        =   18
         Top             =   225
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   8070
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
   Begin ciaXPFrame.XPFrame XPFrame2 
      Height          =   3765
      Left            =   4245
      TabIndex        =   31
      Top             =   2040
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   6641
      Caption         =   " Days Worked "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      BackStyle       =   1
      Radius          =   20
      LicValid        =   -1  'True
      Begin ciaXPPanel.XPPanel XPPanel5 
         Height          =   3525
         Left            =   2580
         TabIndex        =   76
         Top             =   150
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   6218
         HasRightBorder  =   0   'False
         HasTopBorder    =   0   'False
         HasBottomBorder =   0   'False
         LicValid        =   -1  'True
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   10
            Left            =   105
            TabIndex        =   103
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SUN_ND_OT_PAY"
            Top             =   3180
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   9
            Left            =   105
            TabIndex        =   102
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SUN_ND_PAY"
            Top             =   2880
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   8
            Left            =   105
            TabIndex        =   85
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SUN_OT_PAY"
            Top             =   2580
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   7
            Left            =   105
            TabIndex        =   84
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SUN_PAY"
            Top             =   2280
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   6
            Left            =   105
            TabIndex        =   83
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SA_NDIFF_PAY"
            Top             =   1980
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   5
            Left            =   105
            TabIndex        =   82
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:SA_REG_PAY"
            Top             =   1680
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   4
            Left            =   105
            TabIndex        =   81
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:NDIFF_OT_PAY"
            Top             =   1260
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   3
            Left            =   105
            TabIndex        =   80
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:REG_OT_PAY"
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   2
            Left            =   105
            TabIndex        =   79
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:HOL_PAY"
            Top             =   660
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   1
            Left            =   105
            TabIndex        =   78
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:NDIFF_PAY"
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox Text9 
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
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   0
            Left            =   105
            TabIndex        =   77
            Tag             =   "1"
            Text            =   "99,999.99"
            ToolTipText     =   "NUM:REG_PAY"
            Top             =   60
            Width           =   1155
         End
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1590
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "NUM:REG_OT_HR"
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   1590
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "NUM:NDIFF_OT_HR"
         Top             =   1410
         Width           =   900
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   1590
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "NUM:HOLIDAY"
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   1590
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "NUM:NDIFF_DAY"
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   1590
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "NUM:REG_DAY"
         Top             =   210
         Width           =   900
      End
      Begin ciaXPPanel.XPPanel XPPanel2 
         Height          =   1920
         Left            =   60
         TabIndex        =   57
         Top             =   1740
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   3387
         HasLeftBorder   =   0   'False
         HasRightBorder  =   0   'False
         HasBottomBorder =   0   'False
         LicValid        =   -1  'True
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   15
            Left            =   1530
            TabIndex        =   12
            Tag             =   "1"
            ToolTipText     =   "NUM:SUN_ND_OT"
            Top             =   1575
            Width           =   900
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   14
            Left            =   1530
            TabIndex        =   11
            Tag             =   "1"
            ToolTipText     =   "NUM:SUN_ND"
            Top             =   1275
            Width           =   900
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   1530
            TabIndex        =   7
            Tag             =   "1"
            ToolTipText     =   "NUM:SA_REG_OT"
            Top             =   75
            Width           =   900
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   1530
            TabIndex        =   9
            Tag             =   "1"
            ToolTipText     =   "NUM:SUN_HR"
            Top             =   675
            Width           =   900
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   1530
            TabIndex        =   8
            Tag             =   "1"
            ToolTipText     =   "NUM:SA_NDIFF_OT"
            Top             =   375
            Width           =   900
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            Index           =   8
            Left            =   1530
            TabIndex        =   10
            Tag             =   "1"
            ToolTipText     =   "NUM:SUN_OT"
            Top             =   975
            Width           =   900
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday ND OT"
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
            Left            =   60
            TabIndex        =   101
            Top             =   1650
            Width           =   1800
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday Night Diff"
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
            Left            =   60
            TabIndex        =   100
            Top             =   1350
            Width           =   1800
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SA Regular OT"
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
            Left            =   60
            TabIndex        =   61
            Top             =   150
            Width           =   1800
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday Hours"
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
            Left            =   60
            TabIndex        =   60
            Top             =   750
            Width           =   1800
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "SA Night Diff"
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
            Left            =   60
            TabIndex        =   59
            Top             =   450
            Width           =   1800
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday OT"
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
            Left            =   60
            TabIndex        =   58
            Top             =   1050
            Width           =   1800
         End
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Regular OT"
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
         TabIndex        =   42
         Top             =   570
         Width           =   1800
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Night Diff OT"
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
         TabIndex        =   41
         Top             =   1470
         Width           =   1800
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays"
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
         TabIndex        =   40
         Top             =   870
         Width           =   1800
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Night Diff Days"
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
         TabIndex        =   39
         Top             =   1170
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. Days Worked"
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
         Top             =   270
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project name  :   Dong-in Payroll System
' module        :   frmTransaction
' description   :   module for Days Worked/Transactions to Process in Payroll
' programmer    :   _-=[ srm ]=-_
' date          :   23 Oct 2005

Option Explicit
    Dim nAdd As Integer, _
        nNoofDays As Double, _
        myArray As Variant, _
        myArray2 As Variant, _
        oTempADO As New ADODB.Recordset
Sub Compute2()
    Dim nCtr As Integer, _
        nTotDedAmt As Double, _
        nGrossAmt As Double, _
        nsunday As Double
    
    Text2(8).Text = (Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Val(Text5(1).Text) + Val(Text5(0).Text), "###0.000")))
    
    nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
                Val(Format(Text9(3).Text, "###0.00")) + _
                Val(Format(Text9(1).Text, "###0.00")) + _
                Val(Format(Text9(4).Text, "###0.00")) + _
                Val(Format(Text9(2).Text, "###0.00")) + _
                Val(Format(Text2(5).Text, "###0.00")) + _
                Val(Format(Text2(8).Text, "###0.00")) + _
                Val(Format(Text5(12).Text, "###0.00")) + _
                Val(Format(Text5(9).Text, "###0.00")) + _
                Val(Format(Text5(11).Text, "###0.00"))
    
    ' ===================>>> edited code for ectc
    If gCompanyID = "0001" Then
        OpenQueryDNS "SELECT * FROM DI3670 WHERE PAYSTATUS=2 AND EMPID=" & cQuote & Text2(0).Text & cQuote, objdbRs, False
        
        If objdbRs.RecordCount > 0 Then
                
            If Not lExtension Then
                nsunday = Val(Format(Text9(7).Text, "###0.00")) + _
                            Val(Format(Text9(8).Text, "###0.00")) + _
                            Val(Format(Text9(9).Text, "###0.00")) + _
                            Val(Format(Text9(10).Text, "###0.00")) + _
                            Val(Format(Text2(9).Text, "###0.00"))
            End If
        Else
             If lExtension Then
            
                nsunday = Val(Format(Text9(7).Text, "###0.00")) + _
                            Val(Format(Text9(8).Text, "###0.00")) + _
                            Val(Format(Text9(9).Text, "###0.00")) + _
                            Val(Format(Text9(10).Text, "###0.00")) + _
                            Val(Format(Text2(9).Text, "###0.00"))
            End If
        End If
    Else
        If lExtension Then
            
            nsunday = Val(Format(Text9(7).Text, "###0.00")) + _
                        Val(Format(Text9(8).Text, "###0.00")) + _
                        Val(Format(Text9(9).Text, "###0.00")) + _
                        Val(Format(Text9(10).Text, "###0.00")) + _
                        Val(Format(Text2(9).Text, "###0.00"))
        Else
            If gCompanyID = "0005" Then
                nsunday = Val(Format(Text9(7).Text, "###0.00")) + _
                            Val(Format(Text9(8).Text, "###0.00")) + _
                            Val(Format(Text9(9).Text, "###0.00")) + _
                            Val(Format(Text9(10).Text, "###0.00")) + _
                            Val(Format(Text2(9).Text, "###0.00"))
            
                nGrossAmt = nGrossAmt + nsunday
                
                
            End If
        End If
    End If
    
    If gCompanyID = "0001" Then
        If lExtension Then
            nGrossAmt = nGrossAmt + nsunday
        End If
    End If
    ' ===================>>> edited code for ectc
    
    With MSHFlexGrid2
        For nCtr = 1 To (.Rows - 1)
            nTotDedAmt = nTotDedAmt + Val(.TextMatrix(nCtr, 3))
        Next nCtr
    End With
    
    Text6.Text = Format(nTotDedAmt, "##,##0.00")
    Text23.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")) - nTotDedAmt, "##,##0.00")       ' --> Net

End Sub
Sub Compute()
    Dim nGrossAmt, _
        nDedAmt, _
        nTotAmt, _
        nTotExempt, _
        nNetAmt As Double, _
        nCtr2, _
        nCtr As Integer, _
        cDedID, _
        cSqlStmt As String, _
        aDedAmt As Variant, _
        oRecordSet As New ADODB.Recordset, _
        lWithTax As Boolean, _
        lAllDed As Boolean, _
        aTmpTax As Variant, _
        nholiday As Integer, _
        nholWPay As Integer, _
        nHolRegDay As Integer, _
        nHolWPND As Integer
        
    Dim dStartDate As String, _
        dEndDate
    
        
    If oTempADO.RecordCount = 0 Then Exit Sub
    
    ' --> 20061005
    For nCtr = 0 To UBound(aTaxExempt)
        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
        cDedID = cDedID & aTaxExempt(nCtr) & ","
    Next nCtr
    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
        
    nTotAmt = 0
    nNetAmt = 0
    nGrossAmt = 0
    
    aDedAmt = Array(0#, 0#, 0#, 0#, 0#)
    
    If Val(Text2(7).Text) <> 1 Then
        Text9(0).Text = Format(Round((Val(Text5(0).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                                                                    ' --> Reg Day Pay
        Text9(1).Text = Format(Round((Val(Text5(1).Text) * Val(Format(Text2(4).Text, "###0.000")) * 1.1), 2), "##,##0.00")                                                              ' --> NDiff Days Pay
        
        'period for date range
        OpenQueryDNS " SELECT PERIODID, DATE_START, DATE_END, DURATION FROM pa7730 Where periodid = " & cQuote & lblPeriod & cQuote, objdbRs, False
        dStartDate = IIf(objdbRs.RecordCount > 0, objdbRs("DATE_START"), "")
        dEndDate = IIf(objdbRs.RecordCount > 0, objdbRs("DATE_END"), "")
        
        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
        
        
        Text9(3).Text = Format(Round((Val(Text5(3).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25)), 2), "##,##0.00")                                                     ' --> Reg OT Pay
        Text9(4).Text = Format(Round((Val(Text5(4).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.1 * 1.25)), 2), "##,##0.00")                                               ' --> NDiff OT Pay
        Text9(5).Text = Format(Round((Val(Text5(5).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25)), 2), "##,##0.00")                                                     ' --> SA Reg OT Pay
        Text9(6).Text = Format(Round((Val(Text5(6).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25 * 1.1)), 2), "##,##0.00")                                               ' --> SA NDiff OT Pay
        Text9(7).Text = Format(Round((Val(Text5(7).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3)), 2), "##,##0.00")                                                      ' --> Sunday Hours Pay
        Text9(8).Text = Format(Round((Val(Text5(8).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.3)), 2), "##,##0.00")                                                ' --> Sunday OT Hours Pay
        Text9(9).Text = Format(Round((Val(Text5(14).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.1)), 2), "##,##0.00")                                               ' --> Sunday NDiff Hours Pay
        Text9(10).Text = Format(Round((Val(Text5(15).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.1 * 1.3)), 2), "##,##0.00")                                        ' --> Sunday NDiff OT Hours Pay
    Else
'        MsgBox Format(Text2(4).Text, "###0.00") & vbCrLf &
'               nNoofDays & vbCrLf & _
'               Val(Text5(0).Text) & vbCrLf & _
'               (Val(Format(Text2(4).Text, "###0.00")) / 26.08)
        Text9(0).Text = Round((Val(Format(Text2(4).Text, "###0.00")) / 2) - ((Val(Format(Text2(4).Text, "###0.00")) / 26.08) * (nNoofDays - Val(Text5(0).Text))), 2)
        For nCtr = 1 To 8
            Text9(nCtr).Text = 0
        Next nCtr
    End If
    
'    nGrossAmt = RegPay +
'                RegOTPay +
'                NDiffPay +
'                NDiffOTPay +
'                HolPay +
'                PosAllow +
'                COLA +
'                Incentive Leave +
'                13th Month Pay +       ' --> remove first for deduction purposes...
'                Adjustment

'    nNetAmt = SARegOTPay +
'              SANDiffOTPay +
'              SunCola +
'              SunPay +
'              SunOTPay +
'              SunNDPay +
'              SunNDOTPay +
'              SAAdjPay

    ' --> cola
    Text2(8).Text = Round(Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Val(Text5(1).Text) + Val(Text5(0).Text), "###0.000")), 2)
    
    ' --> sunday cola
    Text2(9).Text = Round(Val(Format(Text2(6).Text, "###0.00")) * ((Val(Format(Text5(7).Text, "###0.00")) + Val(Format(Text5(14).Text, "###0.00"))) / 8), 2)

    ' --> enhanced 20071009
    If lExtension Then
        nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
                    Val(Format(Text9(3).Text, "###0.00")) + _
                    Val(Format(Text9(1).Text, "###0.00")) + _
                    Val(Format(Text9(4).Text, "###0.00")) + _
                    Val(Format(Text9(2).Text, "###0.00")) + _
                    Val(Format(Text2(5).Text, "###0.00")) + _
                    Val(Format(Text2(8).Text, "###0.00")) + _
                    Val(Format(Text5(12).Text, "###0.00")) + _
                    Val(Format(Text5(9).Text, "###0.00")) + _
                    Val(Format(Text5(11).Text, "###0.00"))
    
        nNetAmt = Val(Format(Text9(5).Text, "###0.00")) + _
                  Val(Format(Text9(6).Text, "###0.00")) + _
                  Val(Format(Text2(9).Text, "###0.00")) + _
                  Val(Format(Text9(7).Text, "###0.00")) + _
                  Val(Format(Text9(8).Text, "###0.00")) + _
                  Val(Format(Text9(9).Text, "###0.00")) + _
                  Val(Format(Text9(10).Text, "###0.00")) + _
                  Val(Format(Text5(10).Text, "###0.00"))
'        MsgBox Text2(7).Text
        If Trim(Text2(7).Text) = 2 Then
            nGrossAmt = nGrossAmt + nNetAmt
            nNetAmt = 0
        End If
    Else
        If (gCompanyID = "0001") Or (gCompanyID = "0006") Then
            '20100119
'            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
'                        Val(Format(Text9(3).Text, "###0.00")) + _
'                        Val(Format(Text9(1).Text, "###0.00")) + _
'                        Val(Format(Text9(4).Text, "###0.00")) + _
'                        Val(Format(Text9(2).Text, "###0.00")) + _
'                        Val(Format(Text2(5).Text, "###0.00")) + _
'                        Val(Format(Text2(8).Text, "###0.00")) + _
'                        Val(Format(Text5(12).Text, "###0.00")) + _
'                        Val(Format(Text2(9).Text, "###0.00")) + _
'                        Val(Format(Text9(7).Text, "###0.00")) + _
'                        Val(Format(Text9(8).Text, "###0.00")) + _
'                        Val(Format(Text9(9).Text, "###0.00")) + _
'                        Val(Format(Text9(10).Text, "###0.00")) + _
'                        Val(Format(Text5(9).Text, "###0.00")) + _
'                        Val(Format(Text5(11).Text, "###0.00"))
                        
            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
                        Val(Format(Text9(3).Text, "###0.00")) + _
                        Val(Format(Text9(1).Text, "###0.00")) + _
                        Val(Format(Text9(4).Text, "###0.00")) + _
                        Val(Format(Text9(2).Text, "###0.00")) + _
                        Val(Format(Text2(5).Text, "###0.00")) + _
                        Val(Format(Text2(8).Text, "###0.00")) + _
                        Val(Format(Text5(12).Text, "###0.00")) + _
                        Val(Format(Text9(9).Text, "###0.00")) + _
                        Val(Format(Text9(10).Text, "###0.00")) + _
                        Val(Format(Text5(9).Text, "###0.00")) + _
                        Val(Format(Text5(11).Text, "###0.00"))
        
            nNetAmt = Val(Format(Text9(7).Text, "###0.00")) + _
                      Val(Format(Text9(8).Text, "###0.00")) + _
                      Val(Format(Text9(9).Text, "###0.00")) + _
                      Val(Format(Text9(10).Text, "###0.00")) + _
                      Val(Format(Text2(9).Text, "###0.00")) + _
                      Val(Format(Text5(10).Text, "###0.00"))
                      
            If Trim(Text2(7).Text) = 2 Then
                nGrossAmt = nGrossAmt + nNetAmt
                nNetAmt = 0
            End If
                      
        Else
            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
                        Val(Format(Text9(3).Text, "###0.00")) + _
                        Val(Format(Text9(1).Text, "###0.00")) + _
                        Val(Format(Text9(4).Text, "###0.00")) + _
                        Val(Format(Text9(2).Text, "###0.00")) + _
                        Val(Format(Text2(5).Text, "###0.00")) + _
                        Val(Format(Text2(8).Text, "###0.00")) + _
                        Val(Format(Text5(12).Text, "###0.00")) + _
                        Val(Format(Text2(9).Text, "###0.00")) + _
                        Val(Format(Text9(7).Text, "###0.00")) + _
                        Val(Format(Text9(8).Text, "###0.00")) + _
                        Val(Format(Text9(9).Text, "###0.00")) + _
                        Val(Format(Text9(10).Text, "###0.00")) + _
                        Val(Format(Text5(9).Text, "###0.00")) + _
                        Val(Format(Text5(11).Text, "###0.00"))
        
            nNetAmt = Val(Format(Text9(5).Text, "###0.00")) + _
                      Val(Format(Text9(6).Text, "###0.00")) + _
                      Val(Format(Text5(10).Text, "###0.00"))
        End If
    End If
                
'    MsgBox Val(Format(Text9(0).Text, "###0.00")) & vbCrLf & Val(Format(Text9(3).Text, "###0.00")) & vbCrLf & Val(Format(Text9(1).Text, "###0.00")) & vbCrLf & Val(Format(Text9(2).Text, "###0.00")) & vbCrLf & Val(Format(Text2(5).Text, "###0.00")) & vbCrLf & (Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Val(Text5(1).Text) + Val(Text5(0).Text), "###0.000")))

    If Check1.Value = vbChecked Then
        If Val(Text2(7).Text) = 2 Then
            Text5(13).Text = 0
        Else
            If (oTempADO("active") > 0) And (oTempADO("wap") = 0) Then
                If Val(Text5(13).Text) = 0 Then
                    cSqlStmt = "select ytd_basic, ytd_cola, ytd_gross, ytd_gross_sa from di3670 where empid=" & cQuote & Text2(0).Text & cQuote
                    OpenQueryDNS cSqlStmt, objdbRs, False
                    If (oTempADO("EMP_STAT") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)) Then
                        ' --> 20060724 Text5(13).Text = Format((objdbRs("YTD_BASIC") - objdbRs("YTD_COLA") + (Val(Text9(0).Text) + Val(Text9(1).Text)) - (Val(Text2(6).Text) * Val(Text5(0).Text))) / 12, "###0.00")
                        Text5(13).Text = Format((objdbRs("YTD_BASIC") + (Val(Format(Text9(0).Text, "###0.00")) + Val(Format(Text9(1).Text, "###0.00")))) / 12, "###0.00")
                    ElseIf oTempADO("EMP_STAT") = 2 Then
                        Text5(13).Text = Format((objdbRs("YTD_GROSS") + objdbRs("YTD_GROSS_SA") - objdbRs("YTD_COLA") + nGrossAmt + nNetAmt - (Text2(8).Text)) / 12, "###0.00")
                    Else
                        Text5(13).Text = 0
                    End If
                End If
            End If
        End If
    End If
    
    ' --> Basic Pay
    Text3.Text = Format(Val(Format(Text9(0).Text, "###0.00")) + Val(Format(Text9(1).Text, "###0.00")), "##,##0.00")
    
    ' --> Gross
    Text22.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")), "##,##0.00")
    
    ' --> Net (SA)
    Text7.Text = Format(nNetAmt, "##,##0.00")
    
    If Check2.Value <> vbChecked Then
    '    If (nGrossAmt > 0) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
    
        If (Val(Text2(7).Text) <> 2) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
            lAllDed = True
        Else
            lAllDed = False
        End If
    
        'If (Val(Text2(7).Text) <> 2) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
            With MSHFlexGrid2
                nTotExempt = 0
                nTotAmt = 0
                For nCtr = 1 To (.Rows - 1)
                    aDedAmt = Array(0#, 0#, 0#, 0#, 0#)
                    Select Case .TextMatrix(nCtr, 1)
                        Case "001"      ' --> SSS Premium
                            If lAllDed Then
                                'cSqlStmt = "select ER_SS, EE_SS, ER_EC from pa7770 where " & (oTempADO("GROSS16231") + nGrossAmt) & " between range1 and range2"
                                '20090616
                                cSqlStmt = "select ER_SS, EE_SS, ER_EC from pa7770 where " & (oTempADO("GROSS16231") + nGrossAmt + Val(Format(Text5(13).Text, "###0.00"))) & " between range1 and range2"
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                .TextMatrix(nCtr, 3) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("EE_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSPREM1215"), 0), 2)
                                .TextMatrix(nCtr, 4) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("ER_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSER1215"), 0), 2)
        '                        aDedAmt(0) = IIf(objdbRs.RecordCount > 0, objdbRs("EE_TOT"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSPREM1215"), 0)
        '                        aDedAmt(1) = IIf(objdbRs.RecordCount > 0, objdbRs("ER_TOT"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSER1215"), 0)
                                If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
                                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
                            End If
                            
                        Case "003"      ' --> Pag-Ibig Premium, revised 20120502
                            If lAllDed Then
                                cSqlStmt = "select DEF_AMT from pa3330 where dedid='003'"
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                If objdbRs.RecordCount > 0 Then
                                                                        
                                    nDedAmt = (oTempADO("GROSS16231") + nGrossAmt + Val(Format(Text5(13).Text, "###0.00")))
                                    
                                    If (nDedAmt * IIf((oTempADO("GROSS16231") + nGrossAmt + Val(Format(Text5(13).Text, "###0.00"))) > 1500, 0.02, 0.01)) >= 100 Then
                                        .TextMatrix(nCtr, 3) = objdbRs("def_amt")
                                    Else
                                        .TextMatrix(nCtr, 3) = Round(nDedAmt * IIf((oTempADO("GROSS16231") + nGrossAmt + Val(Format(Text5(13).Text, "###0.00"))) > 1500, 0.02, 0.01), 2)
                                    End If
                                    
                                    ' --> employer's contribution
                                    If nDedAmt > 5000 Then
                                        .TextMatrix(nCtr, 4) = objdbRs("def_amt")
                                    Else
                                        .TextMatrix(nCtr, 4) = Round(nDedAmt * 0.02, 2)
                                    End If
                                    
                                    If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
                                    nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
                                End If
                            End If
                            
                            ' --> Pag-Ibig Premium, revised 20070705
'                            If lAllDed Then
'                                cSqlStmt = "select DEF_AMT from pa3330 where dedid='003'"
'                                OpenQueryDNS cSqlStmt, objdbRs, False
'                                If objdbRs.RecordCount > 0 Then
'                                    nDedAmt = (oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00")))
'                                    If (nDedAmt * IIf((oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00"))) > 1500, 0.02, 0.01)) >= 100 Then
'                                        .TextMatrix(nCtr, 3) = objdbRs("def_amt")
'                                    Else
'                                        .TextMatrix(nCtr, 3) = Round(nDedAmt * IIf((oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00"))) > 1500, 0.02, 0.01), 2)
'                                    End If
'
'                                    ' --> employer's contribution
'                                    If nDedAmt > 5000 Then
'                                        .TextMatrix(nCtr, 4) = objdbRs("def_amt")
'                                    Else
'                                        .TextMatrix(nCtr, 4) = Round(nDedAmt * 0.02, 2)
'                                    End If
'
'                                    If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
'                                    nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                                End If
'                            End If
                        
                        Case "005"      ' --> Philhealth/Medicare
                            If lAllDed Then
                                cSqlStmt = "select def_amt from di3673 " & _
                                           " where (empid=" & cQuote & Text2(0).Text & cQuote & ")" & _
                                           " and (dedid='005')" & _
                                           " and (" & IIf(oTempADO("period_stat") = 1, "period1=1", "period2=1") & ")"
                                OpenQueryDNS cSqlStmt, oRecordSet, False
                                
                                cSqlStmt = "select PS, ES from PA7454 where " & (oTempADO("GROSS16231") + nGrossAmt) & " between range1 and range2"
                                OpenQueryDNS cSqlStmt, objdbRs, False
                                If oRecordSet.RecordCount > 0 Then
                                    .TextMatrix(nCtr, 3) = Round(oRecordSet("def_amt") - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0), 2)
                                Else
                                    .TextMatrix(nCtr, 3) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("PS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0), 2)
                                End If
                                If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
                                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
                            End If
                        Case "006"      ' --> Withholding Tax
                            If lAllDed Then
                                '2008-12-17 d2 yung para sa walang rate pag nag edit
                                If oTempADO("emp_stat") = 2 And Val(Text2(4).Text) > gBasicRate Then
    '                                MsgBox Text2(4).Text
                                    lWithTax = True
                                    nCtr2 = nCtr
        '                            cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
        '                                       " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("GROSS16231") + nGrossAmt) & ">=ded_amt2)" & _
        '                                       " order by ded_amt2 desc limit 1"
        '                            OpenQueryDNS cSqlStmt, objdbRs, False
        '                            If objdbRs.RecordCount > 0 Then
        '                                If objdbRs("DED_PCT") > 0 Then
        '                                    nNetAmt = objdbRs("DED_AMT") + (((oTempADO("GROSS16231") + nGrossAmt) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
        '                                Else
        '                                    nNetAmt = 0
        '                                End If
        '                                .TextMatrix(nCtr, 3) = Round(nNetAmt, 2)
        '                            End If
                                Else
                                    lWithTax = False
                                End If
                            End If
                        Case Else       ' --> not fix deduction
                            nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
                            If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
    
    '                        cSqlStmt = "select def_amt, cut_off_amt, acc_amt from di3673 " & _
    '                                   " where (empid=" & cQuote & Text2(0).Text & cQuote & ")" & _
    '                                   " and (dedid=" & cQuote & oRSetDed("DEDID") & cQuote & ")" & _
    '                                   " and (period" & oTempADO("period_stat") & "=1)"
    '                        OpenQueryDNS cSqlStmt, oRecordSet, False
    '                        If oRecordSet.RecordCount > 0 Then
    '                            If (oRecordSet("acc_amt") + oRecordSet("def_amt")) > oRecordSet("cut_off_amt") Then
    '                                .TextMatrix(nCtr, 3) = oRecordSet("cut_off_amt") - oRecordSet("acc_amt")
    '                            Else
    '                                .TextMatrix(nCtr, 3) = oRecordSet("def_amt")
    '                            End If
    '                        Else
    '                        End If
                        
                    End Select
                    
    '                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
                Next nCtr
            
                ' --> 20061005
                If lWithTax Then
                
                    ' --> 20070705, reset to 0, previously used to SA Net total amount...
                    nNetAmt = 0
                    
                    ' --> revised 20070105
                    If Val(.TextMatrix(nCtr2, 4)) <> 0 Then
                        cSqlStmt = "select year(date_end) as year_end from pa7730 where periodid=" & cQuote & lblPeriod.Caption & cQuote
                        OpenQueryDNS cSqlStmt, objdbRs, False
                        
                        If Val(.TextMatrix(nCtr2, 4)) = Val(.TextMatrix(nCtr2, 3)) Then
                            aTmpTax = ComputeTax(lblPeriod.Caption, _
                                                 Text2(0).Text, _
                                                 cDedID, _
                                                 objdbRs("year_end"), _
                                                 nGrossAmt, _
                                                 nTotExempt, _
                                                 False)
                            nNetAmt = aTmpTax(0)
                            .TextMatrix(nCtr2, 3) = Round(nNetAmt, 2)
                            .TextMatrix(nCtr2, 4) = Round(nNetAmt, 2)
                            .TextMatrix(nCtr2, 8) = aTmpTax(1)
                        Else
                            nNetAmt = Format(.TextMatrix(nCtr2, 3), "##0.#00")
                        End If
                    Else
                        cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
                                   " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("TAX1215") + nGrossAmt - nTotExempt - Val(Format(Text5(12).Text, "###0.00"))) & ">=ded_amt2)" & _
                                   " order by ded_amt2 desc limit 1"
                        
                        OpenQueryDNS cSqlStmt, objdbRs, False
                        If objdbRs.RecordCount > 0 Then
                            If objdbRs("DED_PCT") > 0 Then
                                nNetAmt = objdbRs("DED_AMT") + (((oTempADO("TAX1215") + nGrossAmt - nTotExempt - Val(Format(Text5(12).Text, "###0.00"))) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
                            Else
                                nNetAmt = 0
                            End If
                            .TextMatrix(nCtr2, 3) = Round(nNetAmt, 2)
                        End If
                    End If
                    ' --> end revision 20070105
                    
                    nTotAmt = nTotAmt + nNetAmt
                End If
                
            End With
        'End If
    Else
        nTotAmt = 0
    End If
'                Val(Format(Text5(13).Text, "###0.00")) + _

    Text6.Text = Format(nTotAmt, "##,##0.00")                   ' --> Total Deduction
    Text23.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")) - nTotAmt, "##,##0.00")       ' --> Net
    
    Set oRecordSet = Nothing
End Sub


'Sub Compute()
'    Dim nGrossAmt, _
'        nDedAmt, _
'        nTotAmt, _
'        nTotExempt, _
'        nNetAmt As Double, _
'        nCtr2, _
'        nCtr As Integer, _
'        cDedID, _
'        cSqlStmt As String, _
'        aDedAmt As Variant, _
'        oRecordSet As New ADODB.Recordset, _
'        lWithTax As Boolean, _
'        lAllDed As Boolean, _
'        aTmpTax As Variant, _
'        nholiday As Integer, _
'        nholWPay As Integer, _
'        nHolRegDay As Integer, _
'        nHolWPND As Integer, _
'        lHol1 As Boolean, _
'        lHol2 As Boolean
'
'    Dim dStartDate As String, _
'        dEndDate
'
'
'    If oTempADO.RecordCount = 0 Then Exit Sub
'
'    ' --> 20061005
'    For nCtr = 0 To UBound(aTaxExempt)
'        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
'        cDedID = cDedID & aTaxExempt(nCtr) & ","
'    Next nCtr
'    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
'
'    nTotAmt = 0
'    nNetAmt = 0
'    nGrossAmt = 0
'
'    aDedAmt = Array(0#, 0#, 0#, 0#, 0#)
'
''    If Val(Text2(7).Text) <> 1 Then
''        Text9(0).Text = Format(Round((Val(Text5(0).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                                                                    ' --> Reg Day Pay
''        Text9(1).Text = Format(Round((Val(Text5(1).Text) * Val(Format(Text2(4).Text, "###0.000")) * 1.1), 2), "##,##0.00")                                                              ' --> NDiff Days Pay
''
''        'period for date range
''        OpenQueryDNS " SELECT PERIODID, DATE_START, DATE_END, DURATION FROM pa7730 Where periodid = " & cQuote & lblPeriod & cQuote, objdbRs, False
''        dStartDate = IIf(objdbRs.RecordCount > 0, objdbRs("DATE_START"), "")
''        dEndDate = IIf(objdbRs.RecordCount > 0, objdbRs("DATE_END"), "")
''
''
''        '20110905 addition to the august 21 sprcial holiday
''
''        If gCompanyID = "0002" Then
''            cSqlStmt = "select * from di3670 where empid = " & cQuote & Text2(0) & cQuote
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                If objdbRs("emp_stat") = 2 Then
''                    If Val(Text5(2).Text) >= 3 Then
''                        Text5(2).Text = Val(Text5(2).Text) - 1
''
''                        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''
''
''
''                        Text9(2).Text = Format(Round(Text9(2).Text + (((1 * Val(Text2(4).Text)) * 1.5) + Val(Text2(6).Text)), 2), "##,##0.00")                    ' --> Holiday Pay
''                        Text5(2).Text = Val(Text5(2).Text) + 1
''                    Else
''                        If Val(Text5(2).Text) = 2 Then
''                            Text9(2).Text = Format(Round((((2 * Val(Text2(4).Text)) * 1.5) + Val(Text2(6).Text)), 2), "##,##0.00")                    ' --> Holiday Pay
''                        Else
''                            If Val(Text5(2).Text) = 1 Then
''                                Text9(2).Text = Format(Round((((1 * Val(Text2(4).Text)) * 1.5) + Val(Text2(6).Text)), 2), "##,##0.00")                    ' --> Holiday Pay
''                            Else
''                                Text9(2).Text = 0
''                            End If
''                        End If
''                    End If
''                Else
''                    cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                               " left join di3670 b on a.empid=b.empid " & _
''                               " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid = " & cQuote & Text2(0) & cQuote & _
''                               " And a.sun_hr <> 0 And b.emp_stat <>0 "
''                    OpenQueryDNS cSqlStmt, objdbRs, False
''                    If objdbRs.RecordCount > 0 Then
''                        If objdbRs("emp_stat") <> 0 Then
''                            If objdbRs("emp_stat") = 2 Then
''                                Text5(2) = Text5(2) - 1
''                                Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                                Text9(2).Text = Format(Round(Text9(2).Text + ((1 * Text2(4).Text * 0.2)), 2), "##,##0.00")                    ' --> Holiday Pay
''                                Text5(2) = Text5(2) + 1
''                            End If
''                        End If
''                    Else
''                        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                    End If
''                End If
''            End If
''        Else
''            cSqlStmt = "select * from di3670 where empid = " & cQuote & Text2(0) & cQuote
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs("emp_stat") <> 0 Then
''                If objdbRs("wap") = 0 Then
''                    cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                               " left join di3670 b on a.empid=b.empid " & _
''                               " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & Text2(0) & cQuote & " And a.sun_hr <> 0 And b.emp_stat <> 0 "
''                    OpenQueryDNS cSqlStmt, objdbRs, False
''                    If objdbRs.RecordCount > 0 Then
''                        lHol1 = True
''                    Else
''                        lHol1 = False
''                    End If
''
''                    cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                               " left join di3670 b on a.empid=b.empid " & _
''                               " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid =" & cQuote & Text2(0) & cQuote & " And a.sun_nd <> 0 And b.emp_stat <> 0 "
''                    OpenQueryDNS cSqlStmt, objdbRs, False
''                    If objdbRs.RecordCount > 0 Then
''                        lHol2 = True
''                    Else
''                        lHol2 = False
''                    End If
''
''                    If lHol1 = True Or lHol2 = True Then
''                        Text5(2) = Val(Text5(2).Text) - 1
''                        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                        Text9(2).Text = Format(Round(Text9(2).Text + ((1 * Text2(4).Text * 0.2)), 2), "##,##0.00")                    ' --> Holiday Pay
''                        Text5(2) = Text5(2) + 1
''                    Else
''        '                aTimeInfo(7) = aTimeInfo(7) - 1
''        '                aTimeInfo(8) = 0
''                        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                    End If
''                Else
''                    Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                End If
''            Else
''                If objdbRs("wap") = 1 Then
''                    Text9(2).Text = 0
''                Else
''                    Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                End If
''            End If
''        End If
'
'
'
''            cSqlStmt = " select sun_hr,b.emp_stat from di36770 a " & _
''                       " left join di3670 b on a.empid=b.empid " & _
''                       " Where a.Date = " & cQuote & "2011-08-21" & cQuote & " and a.empid = " & cQuote & Text2(0) & cQuote & _
''                       " And a.sun_hr <> 0 And b.emp_stat <>0 "
''            OpenQueryDNS cSqlStmt, objdbRs, False
''            If objdbRs.RecordCount > 0 Then
''                If objdbRs("emp_stat") <> 0 Then
''                    Text5(2) = Text5(2) - 1
''                    Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''                    Text9(2).Text = Format(Round(1 * Val(Text2(6).Text) + ((1 * Text2(4).Text * 1.5)), 2), "##,##0.00")                       ' --> Holiday Pay
''                    Text5(2) = Text5(2) + 1
''                End If
''            Else
''                Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
''            End If
''        End If
'
'
'
''        Text9(2).Text = Format(Round(Val(Text5(2).Text) * Val(Text2(6).Text) + (Val(Text5(2).Text) * Val(Format(Text2(4).Text, "###0.000"))), 2), "##,##0.00")                          ' --> Holiday Pay
'
'
'
'        Text9(3).Text = Format(Round((Val(Text5(3).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25)), 2), "##,##0.00")                                                     ' --> Reg OT Pay
'        Text9(4).Text = Format(Round((Val(Text5(4).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.1 * 1.25)), 2), "##,##0.00")                                               ' --> NDiff OT Pay
'        Text9(5).Text = Format(Round((Val(Text5(5).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25)), 2), "##,##0.00")                                                     ' --> SA Reg OT Pay
'        Text9(6).Text = Format(Round((Val(Text5(6).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.25 * 1.1)), 2), "##,##0.00")                                               ' --> SA NDiff OT Pay
'        Text9(7).Text = Format(Round((Val(Text5(7).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3)), 2), "##,##0.00")                                                      ' --> Sunday Hours Pay
'        Text9(8).Text = Format(Round((Val(Text5(8).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.3)), 2), "##,##0.00")                                                ' --> Sunday OT Hours Pay
'        Text9(9).Text = Format(Round((Val(Text5(14).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.1)), 2), "##,##0.00")                                               ' --> Sunday NDiff Hours Pay
'        Text9(10).Text = Format(Round((Val(Text5(15).Text) * ((Val(Format(Text2(4).Text, "###0.000")) / 8) * 1.3 * 1.1 * 1.3)), 2), "##,##0.00")                                        ' --> Sunday NDiff OT Hours Pay
'    Else
''        MsgBox Format(Text2(4).Text, "###0.00") & vbCrLf &
''               nNoofDays & vbCrLf & _
''               Val(Text5(0).Text) & vbCrLf & _
''               (Val(Format(Text2(4).Text, "###0.00")) / 26.08)
'        Text9(0).Text = Round((Val(Format(Text2(4).Text, "###0.00")) / 2) - ((Val(Format(Text2(4).Text, "###0.00")) / 26.08) * (nNoofDays - Val(Text5(0).Text))), 2)
'        For nCtr = 1 To 8
'            Text9(nCtr).Text = 0
'        Next nCtr
'    End If
'
''    nGrossAmt = RegPay +
''                RegOTPay +
''                NDiffPay +
''                NDiffOTPay +
''                HolPay +
''                PosAllow +
''                COLA +
''                Incentive Leave +
''                13th Month Pay +       ' --> remove first for deduction purposes...
''                Adjustment
'
''    nNetAmt = SARegOTPay +
''              SANDiffOTPay +
''              SunCola +
''              SunPay +
''              SunOTPay +
''              SunNDPay +
''              SunNDOTPay +
''              SAAdjPay
'
'    ' --> cola
'    Text2(8).Text = Round(Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Val(Text5(1).Text) + Val(Text5(0).Text), "###0.000")), 2)
'
'    ' --> sunday cola
'    Text2(9).Text = Round(Val(Format(Text2(6).Text, "###0.00")) * ((Val(Format(Text5(7).Text, "###0.00")) + Val(Format(Text5(14).Text, "###0.00"))) / 8), 2)
'
'    ' --> enhanced 20071009
'    If lExtension Then
'        nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
'                    Val(Format(Text9(3).Text, "###0.00")) + _
'                    Val(Format(Text9(1).Text, "###0.00")) + _
'                    Val(Format(Text9(4).Text, "###0.00")) + _
'                    Val(Format(Text9(2).Text, "###0.00")) + _
'                    Val(Format(Text2(5).Text, "###0.00")) + _
'                    Val(Format(Text2(8).Text, "###0.00")) + _
'                    Val(Format(Text5(12).Text, "###0.00")) + _
'                    Val(Format(Text5(9).Text, "###0.00")) + _
'                    Val(Format(Text5(11).Text, "###0.00"))
'
'        nNetAmt = Val(Format(Text9(5).Text, "###0.00")) + _
'                  Val(Format(Text9(6).Text, "###0.00")) + _
'                  Val(Format(Text2(9).Text, "###0.00")) + _
'                  Val(Format(Text9(7).Text, "###0.00")) + _
'                  Val(Format(Text9(8).Text, "###0.00")) + _
'                  Val(Format(Text9(9).Text, "###0.00")) + _
'                  Val(Format(Text9(10).Text, "###0.00")) + _
'                  Val(Format(Text5(10).Text, "###0.00"))
''        MsgBox Text2(7).Text
'        If Trim(Text2(7).Text) = 2 Then
'            nGrossAmt = nGrossAmt + nNetAmt
'            nNetAmt = 0
'        End If
'    Else
'        If (gCompanyID = "0001") Or (gCompanyID = "0006") Or (gCompanyID = "0005") Then
'            '20100119
''            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
''                        Val(Format(Text9(3).Text, "###0.00")) + _
''                        Val(Format(Text9(1).Text, "###0.00")) + _
''                        Val(Format(Text9(4).Text, "###0.00")) + _
''                        Val(Format(Text9(2).Text, "###0.00")) + _
''                        Val(Format(Text2(5).Text, "###0.00")) + _
''                        Val(Format(Text2(8).Text, "###0.00")) + _
''                        Val(Format(Text5(12).Text, "###0.00")) + _
''                        Val(Format(Text2(9).Text, "###0.00")) + _
''                        Val(Format(Text9(7).Text, "###0.00")) + _
''                        Val(Format(Text9(8).Text, "###0.00")) + _
''                        Val(Format(Text9(9).Text, "###0.00")) + _
''                        Val(Format(Text9(10).Text, "###0.00")) + _
''                        Val(Format(Text5(9).Text, "###0.00")) + _
''                        Val(Format(Text5(11).Text, "###0.00"))
'
'            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
'                        Val(Format(Text9(3).Text, "###0.00")) + _
'                        Val(Format(Text9(1).Text, "###0.00")) + _
'                        Val(Format(Text9(4).Text, "###0.00")) + _
'                        Val(Format(Text9(2).Text, "###0.00")) + _
'                        Val(Format(Text2(5).Text, "###0.00")) + _
'                        Val(Format(Text2(8).Text, "###0.00")) + _
'                        Val(Format(Text5(12).Text, "###0.00")) + _
'                        Val(Format(Text5(9).Text, "###0.00")) + _
'                        Val(Format(Text5(11).Text, "###0.00"))
'
'            nNetAmt = Val(Format(Text9(7).Text, "###0.00")) + _
'                      Val(Format(Text9(8).Text, "###0.00")) + _
'                      Val(Format(Text9(9).Text, "###0.00")) + _
'                       Val(Format(Text9(10).Text, "###0.00")) + _
'                      Val(Format(Text2(9).Text, "###0.00")) + _
'                      Val(Format(Text5(10).Text, "###0.00"))
'
'            If gCompanyID = "0005" Then
'                nGrossAmt = nGrossAmt + nNetAmt
'                nNetAmt = 0
'            End If
'
'
'        Else
'            nGrossAmt = Val(Format(Text9(0).Text, "###0.00")) + _
'                        Val(Format(Text9(3).Text, "###0.00")) + _
'                        Val(Format(Text9(1).Text, "###0.00")) + _
'                        Val(Format(Text9(4).Text, "###0.00")) + _
'                        Val(Format(Text9(2).Text, "###0.00")) + _
'                        Val(Format(Text2(5).Text, "###0.00")) + _
'                        Val(Format(Text2(8).Text, "###0.00")) + _
'                        Val(Format(Text5(12).Text, "###0.00")) + _
'                        Val(Format(Text2(9).Text, "###0.00")) + _
'                        Val(Format(Text9(7).Text, "###0.00")) + _
'                        Val(Format(Text9(8).Text, "###0.00")) + _
'                        Val(Format(Text9(9).Text, "###0.00")) + _
'                        Val(Format(Text9(10).Text, "###0.00")) + _
'                        Val(Format(Text5(9).Text, "###0.00")) + _
'                        Val(Format(Text5(11).Text, "###0.00"))
'
'            nNetAmt = Val(Format(Text9(5).Text, "###0.00")) + _
'                      Val(Format(Text9(6).Text, "###0.00")) + _
'                      Val(Format(Text5(10).Text, "###0.00"))
'        End If
'    End If
'
''    MsgBox Val(Format(Text9(0).Text, "###0.00")) & vbCrLf & Val(Format(Text9(3).Text, "###0.00")) & vbCrLf & Val(Format(Text9(1).Text, "###0.00")) & vbCrLf & Val(Format(Text9(2).Text, "###0.00")) & vbCrLf & Val(Format(Text2(5).Text, "###0.00")) & vbCrLf & (Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Val(Text5(1).Text) + Val(Text5(0).Text), "###0.000")))
'
'    If Check1.Value = vbChecked Then
'        If Val(Text2(7).Text) = 2 Then
'            Text5(13).Text = 0
'        Else
'            If (oTempADO("active") > 0) And (oTempADO("wap") = 0) Then
'                If Val(Text5(13).Text) = 0 Then
'                    cSqlStmt = "select ytd_basic, ytd_cola, ytd_gross, ytd_gross_sa from di3670 where empid=" & cQuote & Text2(0).Text & cQuote
'                    OpenQueryDNS cSqlStmt, objdbRs, False
'                    If (oTempADO("EMP_STAT") = 1) Or (g13Month And (oTempADO("EMP_STAT") = 2)) Then
'                        ' --> 20060724 Text5(13).Text = Format((objdbRs("YTD_BASIC") - objdbRs("YTD_COLA") + (Val(Text9(0).Text) + Val(Text9(1).Text)) - (Val(Text2(6).Text) * Val(Text5(0).Text))) / 12, "###0.00")
'                        Text5(13).Text = Format((objdbRs("YTD_BASIC") + (Val(Format(Text9(0).Text, "###0.00")) + Val(Format(Text9(1).Text, "###0.00")))) / 12, "###0.00")
'                    ElseIf oTempADO("EMP_STAT") = 2 Then
'                        Text5(13).Text = Format((objdbRs("YTD_GROSS") + objdbRs("YTD_GROSS_SA") - objdbRs("YTD_COLA") + nGrossAmt + nNetAmt - (Text2(8).Text)) / 12, "###0.00")
'                    Else
'                        Text5(13).Text = 0
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    ' --> Basic Pay
'    Text3.Text = Format(Val(Format(Text9(0).Text, "###0.00")) + Val(Format(Text9(1).Text, "###0.00")), "##,##0.00")
'    ' --> Gross
'    Text22.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")), "##,##0.00")
'
'    ' --> Net (SA)
'    Text7.Text = Format(nNetAmt, "##,##0.00")
'
''    If gCompanyID = "0005" Then
''        ' --> Gross
'''        MsgBox Text5(13).Text
''        Text22.Text = Format(nGrossAmt + nNetAmt + Val(Format(Text5(13).Text, "###0.00")), "##,##0.00")
''
''        ' --> Net (SA)
''        Text7.Text = Format(0, "##,##0.00")
''    Else
''        ' --> Gross
''        Text22.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")), "##,##0.00")
''
''        ' --> Net (SA)
''        Text7.Text = Format(nNetAmt, "##,##0.00")
''
''    End If
'
'    If Check2.Value <> vbChecked Then
'    '    If (nGrossAmt > 0) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
'
'        If (Val(Text2(7).Text) <> 2) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
'            lAllDed = True
'        Else
'            lAllDed = False
'        End If
'
'        'If (Val(Text2(7).Text) <> 2) And ((oTempADO("emp_stat") > 0) And (oTempADO("wap") = 0)) Then
'            With MSHFlexGrid2
'                nTotExempt = 0
'                nTotAmt = 0
'                For nCtr = 1 To (.Rows - 1)
'                    aDedAmt = Array(0#, 0#, 0#, 0#, 0#)
'                    Select Case .TextMatrix(nCtr, 1)
'                        Case "001"      ' --> SSS Premium
'                            If lAllDed Then
'                                'cSqlStmt = "select ER_SS, EE_SS, ER_EC from pa7770 where " & (oTempADO("GROSS16231") + nGrossAmt) & " between range1 and range2"
'                                '20090616
'                                cSqlStmt = "select ER_SS, EE_SS, ER_EC from pa7770 where " & (oTempADO("GROSS16231") + nGrossAmt + Val(Format(Text5(13).Text, "###0.00"))) & " between range1 and range2"
'                                OpenQueryDNS cSqlStmt, objdbRs, False
'                                .TextMatrix(nCtr, 3) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("EE_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSPREM1215"), 0), 2)
'                                .TextMatrix(nCtr, 4) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("ER_SS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSER1215"), 0), 2)
'        '                        aDedAmt(0) = IIf(objdbRs.RecordCount > 0, objdbRs("EE_TOT"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSPREM1215"), 0)
'        '                        aDedAmt(1) = IIf(objdbRs.RecordCount > 0, objdbRs("ER_TOT"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("SSER1215"), 0)
'                                If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
'                                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                            End If
'
'                        Case "003"      ' --> Pag-Ibig Premium, revised 20070705
'                            If lAllDed Then
'                                cSqlStmt = "select DEF_AMT from pa3330 where dedid='003'"
'                                OpenQueryDNS cSqlStmt, objdbRs, False
'                                If objdbRs.RecordCount > 0 Then
'                                    nDedAmt = (oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00")))
'                                    If (nDedAmt * IIf((oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00"))) > 1500, 0.02, 0.01)) >= 100 Then
'                                        .TextMatrix(nCtr, 3) = objdbRs("def_amt")
'                                    Else
'                                        .TextMatrix(nCtr, 3) = Round(nDedAmt * IIf((oTempADO("BASIC1215") + oTempADO("COLA1215") + Val(Format(Text3.Text, "###0.00")) + Val(Format(Text2(8).Text, "###0.00"))) > 1500, 0.02, 0.01), 2)
'                                    End If
'
'                                    ' --> employer's contribution
'                                    If nDedAmt > 5000 Then
'                                        .TextMatrix(nCtr, 4) = objdbRs("def_amt")
'                                    Else
'                                        .TextMatrix(nCtr, 4) = Round(nDedAmt * 0.02, 2)
'                                    End If
'
'                                    If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
'                                    nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                                End If
'                            End If
'
'                        Case "005"      ' --> Philhealth/Medicare
'                            If lAllDed Then
'                                cSqlStmt = "select def_amt from di3673 " & _
'                                           " where (empid=" & cQuote & Text2(0).Text & cQuote & ")" & _
'                                           " and (dedid='005')" & _
'                                           " and (" & IIf(oTempADO("period_stat") = 1, "period1=1", "period2=1") & ")"
'                                OpenQueryDNS cSqlStmt, oRecordSet, False
'
'                                cSqlStmt = "select PS, ES from PA7454 where " & (oTempADO("GROSS16231") + nGrossAmt) & " between range1 and range2"
'                                OpenQueryDNS cSqlStmt, objdbRs, False
'                                If oRecordSet.RecordCount > 0 Then
'                                    .TextMatrix(nCtr, 3) = Round(oRecordSet("def_amt") - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0), 2)
'                                Else
'                                    .TextMatrix(nCtr, 3) = Round(IIf(objdbRs.RecordCount > 0, objdbRs("PS"), 0) - IIf(oTempADO("period_stat") = 2, oTempADO("PS1215"), 0), 2)
'                                End If
'                                If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
'                                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                            End If
'                        Case "006"      ' --> Withholding Tax
'                            If lAllDed Then
'                                '2008-12-17 d2 yung para sa walang rate pag nag edit
'                                If oTempADO("emp_stat") = 2 And Val(Text2(4).Text) > gBasicRate Then
'    '                                MsgBox Text2(4).Text
'                                    lWithTax = True
'                                    nCtr2 = nCtr
'        '                            cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
'        '                                       " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("GROSS16231") + nGrossAmt) & ">=ded_amt2)" & _
'        '                                       " order by ded_amt2 desc limit 1"
'        '                            OpenQueryDNS cSqlStmt, objdbRs, False
'        '                            If objdbRs.RecordCount > 0 Then
'        '                                If objdbRs("DED_PCT") > 0 Then
'        '                                    nNetAmt = objdbRs("DED_AMT") + (((oTempADO("GROSS16231") + nGrossAmt) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
'        '                                Else
'        '                                    nNetAmt = 0
'        '                                End If
'        '                                .TextMatrix(nCtr, 3) = Round(nNetAmt, 2)
'        '                            End If
'                                Else
'                                    lWithTax = False
'                                End If
'                            End If
'                        Case Else       ' --> not fix deduction
'                            nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                            If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
'
'    '                        cSqlStmt = "select def_amt, cut_off_amt, acc_amt from di3673 " & _
'    '                                   " where (empid=" & cQuote & Text2(0).Text & cQuote & ")" & _
'    '                                   " and (dedid=" & cQuote & oRSetDed("DEDID") & cQuote & ")" & _
'    '                                   " and (period" & oTempADO("period_stat") & "=1)"
'    '                        OpenQueryDNS cSqlStmt, oRecordSet, False
'    '                        If oRecordSet.RecordCount > 0 Then
'    '                            If (oRecordSet("acc_amt") + oRecordSet("def_amt")) > oRecordSet("cut_off_amt") Then
'    '                                .TextMatrix(nCtr, 3) = oRecordSet("cut_off_amt") - oRecordSet("acc_amt")
'    '                            Else
'    '                                .TextMatrix(nCtr, 3) = oRecordSet("def_amt")
'    '                            End If
'    '                        Else
'    '                        End If
'
'                    End Select
'
'    '                nTotAmt = nTotAmt + Val(.TextMatrix(nCtr, 3))
'                Next nCtr
'
'                ' --> 20061005
'                If lWithTax Then
'
'                    ' --> 20070705, reset to 0, previously used to SA Net total amount...
'                    nNetAmt = 0
'
'                    ' --> revised 20070105
'                    If Val(.TextMatrix(nCtr2, 4)) <> 0 Then
'                        cSqlStmt = "select year(date_end) as year_end from pa7730 where periodid=" & cQuote & lblPeriod.Caption & cQuote
'                        OpenQueryDNS cSqlStmt, objdbRs, False
'
'                        If Val(.TextMatrix(nCtr2, 4)) = Val(.TextMatrix(nCtr2, 3)) Then
'                            aTmpTax = ComputeTax(lblPeriod.Caption, _
'                                                 Text2(0).Text, _
'                                                 cDedID, _
'                                                 objdbRs("year_end"), _
'                                                 nGrossAmt, _
'                                                 nTotExempt, _
'                                                 False)
'                            nNetAmt = aTmpTax(0)
'                            .TextMatrix(nCtr2, 3) = Round(nNetAmt, 2)
'                            .TextMatrix(nCtr2, 4) = Round(nNetAmt, 2)
'                            .TextMatrix(nCtr2, 8) = aTmpTax(1)
'                        Else
'                            nNetAmt = Format(.TextMatrix(nCtr2, 3), "##0.#00")
'                        End If
'                    Else
'                        cSqlStmt = "select ded_pct, ded_amt, ded_amt2 from pa8293 " & _
'                                   " where (taxid=" & cQuote & oTempADO("TAXID") & cQuote & ") and (" & (oTempADO("TAX1215") + nGrossAmt - nTotExempt - Val(Format(Text5(12).Text, "###0.00"))) & ">=ded_amt2)" & _
'                                   " order by ded_amt2 desc limit 1"
'
'                        OpenQueryDNS cSqlStmt, objdbRs, False
'                        If objdbRs.RecordCount > 0 Then
'                            If objdbRs("DED_PCT") > 0 Then
'                                nNetAmt = objdbRs("DED_AMT") + (((oTempADO("TAX1215") + nGrossAmt - nTotExempt - Val(Format(Text5(12).Text, "###0.00"))) - objdbRs("DED_AMT2")) * (objdbRs("DED_PCT") / 100))
'                            Else
'                                nNetAmt = 0
'                            End If
'                            .TextMatrix(nCtr2, 3) = Round(nNetAmt, 2)
'                        End If
'                    End If
'                    ' --> end revision 20070105
'
'                    nTotAmt = nTotAmt + nNetAmt
'                End If
'
'            End With
'        'End If
'    Else
'        nTotAmt = 0
'    End If
''                Val(Format(Text5(13).Text, "###0.00")) + _
'
'    Text6.Text = Format(nTotAmt, "##,##0.00")                   ' --> Total Deduction
'    Text23.Text = Format(nGrossAmt + Val(Format(Text5(13).Text, "###0.00")) - nTotAmt, "##,##0.00")       ' --> Net
'
'    Set oRecordSet = Nothing
'End Sub

Function GetDept() As String
    Dim cParam As String
    
    OpenQueryDNS "SELECT DISTINCT DEPID FROM PA87260 WHERE PERIODID=" & cQuote & lblPeriod.Caption & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        DoEvents
        While Not objdbRs.EOF
            cParam = cParam & cQuote & objdbRs("DEPID") & cQuote & ","
            objdbRs.MoveNext
        Wend
    End If
    If Trim(cParam) <> "" Then
        GetDept = "(" & left(cParam, Len(cParam) - 1) & ")"
    Else
        GetDept = ""
    End If
'    GetDept = IIf(Trim(cParam) <> "", "(" & left(cParam, Len(cParam) - 1) & ")", cParam)
End Function

Sub FillGrid()
    Dim cSqlStmt As String, _
        cStartDate As String, _
        cEndDate As String
    
    OpenQueryDNS "select workindays from pa7730 where periodid=" & cQuote & lblPeriod.Caption & cQuote, objdbRs, False
    nNoofDays = IIf(objdbRs.RecordCount > 0, objdbRs("workindays"), 0)
    
    cSqlStmt = "SELECT A.EMPID, A.FULLNAME, " & _
               " A.FIRSTNAME, A.LASTNAME " & _
               " FROM PA87260 A " & _
               " WHERE (A.PERIODID=" & cQuote & lblPeriod.Caption & cQuote & ")" & _
               " AND (A.DEPID=" & cQuote & Text1.Text & cQuote & ")" & _
               " ORDER BY FULLNAME"
    OpenQueryDNS cSqlStmt, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        QueryAttach objdbRs, MSHFlexGrid1, myArray, , , , 1
    Else
        SetGridColumn myArray, MSHFlexGrid1
    End If
End Sub

Private Sub Check2_Click()
    Dim nCtr As Integer
    If Check2.Value <> vbChecked Then
        Compute
    Else
        With MSHFlexGrid2
            For nCtr = 1 To (.Rows - 1)
                .TextMatrix(nCtr, 3) = 0
                .TextMatrix(nCtr, 4) = 0
                .TextMatrix(nCtr, 8) = 0
                .TextMatrix(nCtr, 9) = ""
            Next nCtr
        End With
        Text6.Text = Format("0", "##,##0.00")                                       ' --> Total Deduction
        Text23.Text = Format(Val(Format(Text22.Text, "###0.00")), "##,##0.00")      ' --> Net
    End If
End Sub

Private Sub Combo1_Change()
    MSHFlexGrid1.Col = Combo1.ListIndex + 1
    MSHFlexGrid1.Sort = flexSortGenericAscending
End Sub

Private Sub Command1_Click()
    frmLookup.showPopup 2, " where lineid in " & GetDept
    frmLookup.Show 1
    If Trim(cResult) <> "" Then
        OpenQueryDNS "SELECT * FROM DI5463 WHERE lineid=" & cQuote & cResult & cQuote, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            Text1.Text = cResult
            Label1.Caption = objdbRs("LineName")
            FillGrid
            MSHFlexGrid1_EnterCell
        End If
    End If
    
    If Text1.Enabled Then Text1.SetFocus
End Sub

Private Sub Command10_Click()
    On Error GoTo ErrSave
    Dim nCtr As Integer, nPeriod As Integer, _
        cDedID, _
        cSqlStmt As String, _
        aDedAmt As Variant, _
        nDedAmt As Double, _
        nTotExempt As Double, _
        nDay As Integer, nholiday As Integer
    
    aDedAmt = Array(0#, 0#, 0#, 0#, 0#, 0#)     ' --> for deduction purposes
    ' (0)   -   SSS ER
    ' (1)   -   SSS Premium
    ' (2)   -   Withholding Tax
    ' (3)   -   PhilHealth/Medicare PS
    ' (4)   -   PhilHealth/Medicare ES
    ' (5)   -   ?
    
    ' --> 20061005
    For nCtr = 0 To UBound(aTaxExempt)
        If Trim(aTaxExempt(nCtr)) = "" Then Exit For
        cDedID = cDedID & aTaxExempt(nCtr) & ","
    Next nCtr
    If Trim(cDedID) <> "" Then cDedID = left(cDedID, Len(cDedID) - 1)
    
    Select Case MsgBox("Update transaction for Employee ID#" & Text2(0).Text & "?", vbYesNo, "Update Payroll Transaction")
        Case vbYes
            ' --> retrieve period stat here...
            cSqlStmt = "select workindays, holidays, `status` from pa7730 where periodid=" & cQuote & lblPeriod.Caption & cQuote
            OpenQueryDNS cSqlStmt, objdbRs, False
            nPeriod = IIf(objdbRs.RecordCount > 0, objdbRs("STATUS"), 0)
            nDay = IIf(objdbRs.RecordCount > 0, objdbRs("workindays"), 0)
            nholiday = IIf(objdbRs.RecordCount > 0, objdbRs("holidays"), 0)
            
            If nAdd = 1 Then
                cSqlStmt = "select firstname, mname, lastname, PAGIBIGNO, PHEALTHNUM, TIN, SSNUM, DATE_HIRE, DATE_RES " & _
                           "from di3670 where empid=" & cQuote & Text2(0).Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, False
                
                cSqlStmt = "insert into pa87260(PERIODID, PERIOD_STAT, P_DAY, P_HOLIDAY, EMPID, DEPID, POSID, TAXID, ACTIVE, EMP_STAT, PAYSTATUS, WAP, " & _
                           "RATE_AMT, COLA_AMT, COLA, SUN_COLA, POS_ALLOW, REG_DAY, REG_PAY, REG_OT_HR, REG_OT_PAY, SA_REG_OT, SA_REG_PAY, " & _
                           "NDIFF_DAY, NDIFF_PAY, NDIFF_OT_HR, NDIFF_OT_PAY, SA_NDIFF_OT, SA_NDIFF_PAY, HOLIDAY, HOL_PAY, SUN_HR, SUN_PAY, SUN_OT, SUN_OT_PAY, SUN_ND, SUN_ND_PAY, SUN_ND_OT, SUN_ND_OT_PAY, " & _
                           "ADJ_PAY, SA_ADJ_PAY, OTHER_PAY, LEAVE_PAY, M13PAY, GROSS_PAY, BASICPAY, NET_PAY, SA_NET_PAY, " & _
                           "SSSNUM, TINNUM, PAGIBIGNO, PHEALTHNUM, FULLNAME, FIRSTNAME, MNAME, LASTNAME, DATE_HIRE, DATE_RES)values(" & _
                           cQuote & lblPeriod.Caption & cQuote & "," & nPeriod & "," & nDay & "," & nholiday & "," & _
                           cQuote & Text2(0).Text & cQuote & "," & cQuote & Text1.Text & cQuote & "," & cQuote & oTempADO("posid") & cQuote & "," & cQuote & oTempADO("taxid") & cQuote & "," & _
                           oTempADO("active") & "," & oTempADO("emp_stat") & "," & oTempADO("paystatus") & "," & oTempADO("wap") & "," & _
                           Text2(4).Text & "," & Text2(6).Text & "," & Text2(8).Text & "," & Text2(9).Text & "," & Text2(5).Text & "," & _
                           Val(Format(Text5(0).Text, "###0.00")) & "," & Val(Format(Text9(0).Text, "###0.00")) & "," & Val(Format(Text5(3).Text, "###0.00")) & "," & Val(Format(Text9(3).Text, "###0.00")) & "," & Val(Format(Text5(5).Text, "###0.00")) & "," & Val(Format(Text9(5).Text, "###0.00")) & "," & _
                           Val(Format(Text5(1).Text, "###0.00")) & "," & Val(Format(Text9(1).Text, "###0.00")) & "," & Val(Format(Text5(4).Text, "###0.00")) & "," & Val(Format(Text9(4).Text, "###0.00")) & "," & Val(Format(Text5(6).Text, "###0.00")) & "," & Val(Format(Text9(6).Text, "###0.00")) & "," & _
                           Val(Format(Text5(2).Text, "###0.00")) & "," & Val(Format(Text9(2).Text, "###0.00")) & "," & _
                           Val(Format(Text5(7).Text, "###0.00")) & "," & Val(Format(Text9(7).Text, "###0.00")) & "," & Val(Format(Text5(8).Text, "###0.00")) & "," & Val(Format(Text9(8).Text, "###0.00")) & "," & _
                           Val(Format(Text5(14).Text, "###0.00")) & "," & Val(Format(Text9(9).Text, "###0.00")) & "," & _
                           Val(Format(Text5(15).Text, "###0.00")) & "," & Val(Format(Text9(10).Text, "###0.00")) & "," & _
                           Val(Format(Text5(9).Text, "###0.00")) & "," & Val(Format(Text5(10).Text, "###0.00")) & "," & Val(Format(Text5(11).Text, "###0.00")) & "," & Val(Format(Text5(12).Text, "###0.00")) & "," & Val(Format(Text5(13).Text, "###0.00")) & "," & _
                           Val(Format(Text22.Text, "###0.00")) & "," & Val(Format(Text3.Text, "###0.00")) & "," & Val(Format(Text23.Text, "###0.00")) & "," & Val(Format(Text7.Text, "###0.00")) & "," & _
                           cQuote & objdbRs("SSNUM") & cQuote & "," & cQuote & objdbRs("tin") & cQuote & "," & cQuote & objdbRs("pagibigno") & cQuote & "," & cQuote & objdbRs("phealthnum") & cQuote & "," & _
                           cQuote & Text2(1).Text & cQuote & "," & cQuote & objdbRs("firstname") & cQuote & "," & cQuote & objdbRs("mname") & cQuote & "," & cQuote & objdbRs("lastname") & cQuote & "," & _
                           cQuote & Format(objdbRs("date_hire"), "yyyy-mm-dd") & cQuote & "," & _
                           cQuote & Format(objdbRs("date_res"), "yyyy-mm-dd") & cQuote & ")"
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            
                Log2Audit Name, "Add Payroll transaction of Empl ID#" & Text2(0).Text
            Else
                ' --> save header here
                cSqlStmt = "(PERIODID=" & cQuote & lblPeriod.Caption & cQuote & ") AND " & _
                           "(EMPID=" & cQuote & Text2(0).Text & cQuote & ")"
                
                OpenQueryDNS EditField(Me, "PA87260", cSqlStmt), oTempADO, True
                Script2File EditField(Me, "PA87260", cSqlStmt)
                Log2Audit Name, "Update Payroll transaction of Empl ID#" & Text2(0).Text
                
                ' --> update detail here...
                cSqlStmt = "DELETE FROM PA87263 WHERE PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                           " AND EMPID=" & cQuote & Text2(0).Text & cQuote
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
            End If
        
            With MSHFlexGrid2
            
                ShowProgress 0
                
                nTotExempt = 0
                
                For nCtr = 1 To (.Rows - 1)
                
                    ShowProgress 2, (nCtr / (.Rows - 1)) * 100, , , "Saving " & Trim(.TextMatrix(nCtr, 2)) & "..."
                    
                    If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                        cSqlStmt = "INSERT INTO PA87263(PERIODID, PERIOD_STAT,EMPID,DEDID,CTRL_NO,DED_AMT,DED_AMT2,DED_AMT3,COMPUTED)VALUES(" & _
                                   cQuote & lblPeriod.Caption & cQuote & "," & _
                                   nPeriod & "," & _
                                   cQuote & Text2(0).Text & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                                   cQuote & .TextMatrix(nCtr, 9) & cQuote & "," & _
                                   Val(.TextMatrix(nCtr, 3)) & "," & _
                                   Val(.TextMatrix(nCtr, 4)) & "," & _
                                   Val(.TextMatrix(nCtr, 8)) & "," & _
                                   Val(.TextMatrix(nCtr, 5)) & ")"
                        OpenQueryDNS cSqlStmt, objdbRs, True
                        Script2File cSqlStmt
                    End If
                    
                    Select Case .TextMatrix(nCtr, 1)
                        Case "001"      ' --> SSS Premium
                            aDedAmt(0) = Val(.TextMatrix(nCtr, 3))
                            aDedAmt(1) = Val(.TextMatrix(nCtr, 4))
                            aDedAmt(5) = Val(.TextMatrix(nCtr, 8))
                            If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + aDedAmt(0)

'                        Case "003"      ' --> Pag-Ibig Premium

                        Case "005"      ' --> Philhealth/Medicare
                            aDedAmt(3) = Val(.TextMatrix(nCtr, 3))
                            aDedAmt(4) = Val(.TextMatrix(nCtr, 4))
                            If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + aDedAmt(3)
                            
                        Case "006"      ' --> Withholding Tax
                            aDedAmt(2) = Val(.TextMatrix(nCtr, 3))
                            
                        Case Else
                            If InStr(1, cDedID, .TextMatrix(nCtr, 1)) > 1 Then nTotExempt = nTotExempt + Val(.TextMatrix(nCtr, 3))
                        
                    End Select
                    
                    nDedAmt = nDedAmt + Val(.TextMatrix(nCtr, 3))
                Next nCtr
                
                If Val(Text6.Text) <> 0 Then
                    cSqlStmt = "UPDATE PA87260 SET DED_AMT=" & Round(nDedAmt, 2) & "," & _
                               " SSPREM=" & Round(aDedAmt(0), 2) & "," & _
                               " SSER=" & Round(aDedAmt(1), 2) & "," & _
                               " SSS01=" & Round(aDedAmt(0) + aDedAmt(1), 2) & "," & _
                               " EC001=" & aDedAmt(5) & "," & _
                               " MEDICARE=" & Round(aDedAmt(3), 2) & "," & _
                               " MEDICARE2=" & Round(aDedAmt(4), 2) & "," & _
                               " MED01=" & Round(aDedAmt(3) + aDedAmt(4), 2) & "," & _
                               " TAXABLE=" & Val(Format(Text22.Text, "##00.00")) - nTotExempt - Val(Format(Text5(12).Text, "###0.00")) & "," & _
                               " WTAX=" & Round(aDedAmt(2), 2) & "," & _
                               " NET_PAY=" & Val(Format(Text23.Text, "##00.00")) & _
                               " WHERE PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                               " AND EMPID=" & cQuote & Text2(0).Text & cQuote
                Else
                    cSqlStmt = "UPDATE PA87260 SET NET_PAY=" & Val(Format(Text23.Text, "##00.00")) & "," & _
                               " DED_AMT=0," & _
                               " SSPREM=0," & _
                               " SSER=0," & _
                               " SSS01=0," & _
                               " EC001=0," & _
                               " MEDICARE=0," & _
                               " MEDICARE2=0," & _
                               " MED01=0," & _
                               " TAXABLE=" & Val(Format(Text23.Text, "##00.00")) - Val(Format(Text5(12).Text, "###0.00")) & "," & _
                               " WTAX=0 " & _
                               "WHERE PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                               " AND EMPID=" & cQuote & Text2(0).Text & cQuote
                End If
                OpenQueryDNS cSqlStmt, objdbRs, True
                Script2File cSqlStmt
                
                ShowProgress 4
                
            End With
                
        Case vbNo
        
        Case vbCancel
            GoTo endsave
            
    End Select
    
    ' --> enable form's keypreview automatically...
    KeyPreview = True
    
    Lock2User Me.Name, "EMPID", Text2(0).Text, False
    cSqlStmt = Text2(0).Text
    
    If nAdd = 1 Then
        FillGrid
        ChkDupInGrid cSqlStmt, 1, MSHFlexGrid1
    End If
    
    nAdd = 0
    CtrlPanel Me, nAdd
    ClearAll Me, False, True
    
    Text1.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    
    Combo1.Enabled = True
    Text10.Enabled = True
    Check1.Visible = False
    Check2.Enabled = False
    Check2.Value = vbUnchecked
    
    MSHFlexGrid2.FixedCols = 1
    
    MSHFlexGrid1_EnterCell
    MSHFlexGrid1.Enabled = True
    MSHFlexGrid1.SetFocus
    
endsave:
    Exit Sub
    
ErrSave:
    ErrorMsg Err.Number, Err.Description, "Save Payroll Transaction Entry", Name
End Sub

Private Sub Command11_Click()
    If nAdd = 0 Then
        Unload Me
    Else
        If MsgBox("Are you sure you want to abandon your entry?", vbYesNo + vbCritical, App.Title) = vbYes Then
        
            Lock2User Me.Name, "EMPID", Text2(0).Text, False     ' --> 20050321
            
            nAdd = 0
            CtrlPanel Me, nAdd
            ClearAll Me, False, True
            
            Text1.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            
            Combo1.Enabled = True
            Text10.Enabled = True
            Check1.Visible = False
            Check2.Enabled = False
            Check2.Value = vbUnchecked
            
            MSHFlexGrid2.FixedCols = 1
            
            MSHFlexGrid1_EnterCell
            MSHFlexGrid1.Enabled = True
            MSHFlexGrid1.SetFocus
        End If
        
    End If
End Sub

Sub CreateTmpTran(nMode As Integer)
    On Error GoTo ErrCreate
    Dim cSqlStmt As String
    
    If nMode = 0 Then
        cSqlStmt = " CREATE TABLE tmp87260(" & _
                   " [PERIODNAME] CHAR(100),    [EMPID] CHAR(6),        [DEPTNAME] CHAR(100),       [ACTIVE] INTEGER," & _
                   " [FULLNAME] CHAR(100),      [FNAME] CHAR(100),      [LNAME] CHAR(100), " & _
                   " [RATE_AMT] DOUBLE,         [COLA_AMT] DOUBLE,      [POS_ALLOW] DOUBLE, " & _
                   " [REG_DAY] DOUBLE,          [REG_PAY] DOUBLE, " & _
                   " [REG_OT_HR] DOUBLE,        [REG_OT_PAY] DOUBLE, " & _
                   " [NDIFF_DAY] DOUBLE,        [NDIFF_PAY] DOUBLE, " & _
                   " [NDIFF_OT_HR] DOUBLE,      [NDIFF_OT_PAY] DOUBLE, " & _
                   " [HOLIDAY] DOUBLE,          [HOL_PAY] DOUBLE, " & _
                   " [SA_REG_OT] DOUBLE,        [SA_REG_PAY] DOUBLE, " & _
                   " [SA_NDIFF_OT] DOUBLE,      [SA_NDIFF_PAY] DOUBLE, " & _
                   " [SUN_HR] DOUBLE,           [SUN_PAY] DOUBLE, " & _
                   " [SUN_OT] DOUBLE,           [SUN_OT_PAY] DOUBLE, " & _
                   " [ADJ_PAY] DOUBLE,          [SA_ADJ_PAY] DOUBLE, " & _
                   " [OTHER_PAY] DOUBLE,        [LEAVE_PAY] DOUBLE, " & _
                   " [DED_AMT] DOUBLE,          [GROSS_PAY] DOUBLE, " & _
                   " [NET_PAY] DOUBLE,          [SA_NET_PAY] DOUBLE, " & _
                   " [CMPID] char(4))"
    Else
        cSqlStmt = " CREATE TABLE tmp87263(" & _
                   " [PERIODID] CHAR(5),        [EMPID] CHAR(6)," & _
                   " [DEDID] CHAR(3),           [DEDNAME] CHAR(100)," & _
                   " [AMOUNT] DOUBLE)"
    End If

    oTempConn.Execute cSqlStmt
    While oTempConn.State = adStateExecuting
        DoEvents
    Wend
ErrCreate:
    ' in case table is already existing, let's clear it...
    QueryTemp "DELETE FROM " & IIf(nMode = 0, "tmp87260", "tmp87263"), oTempADO, True
End Sub

Private Sub Command2_Click()
    Dim cSqlStmt As String, _
        ntag As Integer
    
    If nAdd = 1 Then
        OpenQueryDNS "select status from pa7730 where periodid=" & cQuote & lblPeriod.Caption & cQuote, objdbRs, False
        ntag = IIf(objdbRs.RecordCount > 0, objdbRs("status"), 0)
        
        
        
        frmLookup.showPopup 3, "where (empid not in (select empid from pa87260 where periodid=" & cQuote & lblPeriod.Caption & cQuote & "))"
        frmLookup.Combo1.ListIndex = 2
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            cSqlStmt = "SELECT " & _
                       ntag & " as period_stat," & _
                       " 0 as gross16231, " & _
                       " 0 as basic1215, " & _
                       " 0 as ssprem1215, " & _
                       " 0 as sser1215, " & _
                       " 0 as ps1215, " & _
                       " 0 as tax1215, " & _
                       " '' as taxid, " & _
                       " a.empid, " & _
                       "  concat(a.lastname,', ',a.firstname,if(trim(a.mname)='','',concat(' ',left(a.mname,1),'. '))) as fullname, " & _
                       "  a.depid, " & _
                       "  ifnull(c.linename,'') as department, " & _
                       "  a.active, a.emp_stat, a.wap, " & _
                       "  a.rate_amt, a.cola_amt, " & _
                       "  a.paystatus, " & _
                       "  a.pos_allow, " & _
                       "  a.posid, " & _
                       "  ifnull(b.posname,'') as position " & _
                       "FROM di3670 a " & _
                       "  left join di7670 b on a.posid=b.posid " & _
                       "  left join di5463 c on a.depid=c.lineid " & _
                       "where a.empid=" & cQuote & cResult & cQuote
            OpenQueryDNS cSqlStmt, oTempADO, False
            If oTempADO.RecordCount > 0 Then
                Text2(0).Text = cResult
                Text2(1).Text = oTempADO("fullname")
                Text2(2).Text = oTempADO("position")
                Text2(3).Text = IIf(oTempADO("emp_stat") = 0, "WAP", IIf(oTempADO("emp_stat") = 1, "Contractual", "Regular"))
                Text2(4).Text = oTempADO("rate_amt")
                Text2(5).Text = oTempADO("pos_allow")
                Text2(6).Text = oTempADO("cola_amt")
                Label27.Caption = IIf(oTempADO("active") = 1, "Resigned", IIf(oTempADO("active") = 2, "Finished Contract", ""))
                Text2(10).Text = oTempADO("BACCNTNO")
                
                Text1.Text = oTempADO("depid")
                Label1.Caption = oTempADO("department")
                
                If oTempADO("wap") = 0 Then
                    cSqlStmt = "select dedid, " & _
                               " dedname, " & _
                               " 0, 0, 0, dedtag, dedtype, 0, '' " & _
                               "from pa3330 " & _
                               "where " & IIf(ntag = 1, "period1=1", "period2=1") & _
                               " order by dedid"
                    OpenQueryDNS cSqlStmt, objdbRs, False
                    If objdbRs.RecordCount > 0 Then
                        QueryAttach objdbRs, MSHFlexGrid2, myArray2, False, , , 1
                    Else
                        SetGridColumn myArray2, MSHFlexGrid2
                    End If
                End If
                Compute
            End If
        End If
        
        Text2(0).SetFocus
    Else
        frmLookup.showPopup 12, " where a.periodid=" & cQuote & lblPeriod.Caption & cQuote
        frmLookup.Show 1
        If Trim(cResult) <> "" Then
            OpenQueryDNS "select depid from pa87260 where empid=" & cQuote & cResult & cQuote, objdbRs, False
            Text1.Text = objdbRs("depid")
            
            OpenQueryDNS "SELECT * FROM DI5463 WHERE lineid=" & cQuote & objdbRs("depid") & cQuote, objdbRs, False
            If objdbRs.RecordCount > 0 Then
                Label1.Caption = objdbRs("LineName")
                FillGrid
                If ChkDupInGrid(cResult, 1, MSHFlexGrid1) Then
                    With MSHFlexGrid1
                        .SetFocus
                        .TopRow = .Row - 1
                        .Row = .Row - 1
                        .RowSel = .Row
                        .ColSel = .Cols - .FixedCols
                    End With
                    MSHFlexGrid1_EnterCell
                End If
    '            MSHFlexGrid1_EnterCell
            End If
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim cSqlStmt As String, _
        nCtr As Integer
        
    CreateTmpTran 0     ' --> header
    CreateTmpTran 1     ' --> detail
    
    ShowProgress 0
    
    cSqlStmt = "INSERT INTO TMP87260(PERIODNAME,EMPID,DEPTNAME,[ACTIVE],FULLNAME,FNAME,LNAME,RATE_AMT,COLA_AMT,POS_ALLOW," & _
               "REG_DAY,REG_PAY,REG_OT_HR,REG_OT_PAY,NDIFF_DAY,NDIFF_PAY,NDIFF_OT_HR,NDIFF_OT_PAY,[HOLIDAY],HOL_PAY," & _
               "SA_REG_OT,SA_REG_PAY,SA_NDIFF_OT,SA_NDIFF_PAY,SUN_HR,SUN_PAY,SUN_OT,SUN_OT_PAY,ADJ_PAY,SA_ADJ_PAY," & _
               "OTHER_PAY,LEAVE_PAY,DED_AMT,GROSS_PAY,NET_PAY,SA_NET_PAY)VALUES(" & _
               cQuote & "for the period " & EncodeStr2(lblDuration.Caption) & cQuote & "," & _
               cQuote & Text2(0).Text & cQuote & "," & _
               cQuote & EncodeStr2(Label1.Caption) & cQuote & "," & _
               oTempADO("ACTIVE") & "," & _
               cQuote & EncodeStr2(Text2(1).Text) & cQuote & ",'',''," & _
               Text2(4).Text & "," & _
               (Val(Format(Text2(6).Text, "###0.00")) * Val(Format(Text5(0).Text, "###0.00"))) & "," & _
               Text2(5).Text & "," & _
               Text5(0).Text & "," & Format(Text9(0).Text, "###0.00") & "," & _
               Text5(3).Text & "," & Format(Text9(3).Text, "###0.00") & "," & _
               Text5(1).Text & "," & Format(Text9(1).Text, "###0.00") & "," & _
               Text5(4).Text & "," & Format(Text9(4).Text, "###0.00") & "," & _
               Text5(2).Text & "," & Format(Text9(2).Text, "###0.00") & "," & _
               Text5(5).Text & "," & Format(Text9(5).Text, "###0.00") & "," & _
               Text5(6).Text & "," & Format(Text9(6).Text, "###0.00") & "," & _
               Text5(7).Text & "," & Format(Text9(7).Text, "###0.00") & "," & _
               Text5(8).Text & "," & Format(Text9(8).Text, "###0.00") & "," & _
               Format(Text5(9).Text, "###0.00") & "," & Format(Text5(10).Text, "###0.00") & "," & _
               Format(Text5(11).Text, "###0.00") & "," & Format(Text5(12).Text, "###0.00") & "," & _
               Format(Text6.Text, "###0.00") & "," & Format(Text22.Text, "###0.00") & "," & Format(Text23.Text, "###0.00") & "," & Format(Text7.Text, "###0.00") & ")"
    QueryTemp cSqlStmt, objdbRs, True
    
    With MSHFlexGrid2
        For nCtr = 1 To (.Rows - 1)
            ShowProgress 2, (nCtr / (.Rows - 1)) * 100
            If Trim(.TextMatrix(nCtr, 1)) <> "" Then
                cSqlStmt = "INSERT INTO TMP87263(PERIODID,EMPID,DEDID,DEDNAME,[AMOUNT])VALUES(" & _
                           cQuote & lblPeriod.Caption & cQuote & "," & _
                           cQuote & Text2(0).Text & cQuote & "," & _
                           cQuote & .TextMatrix(nCtr, 1) & cQuote & "," & _
                           cQuote & EncodeStr2(.TextMatrix(nCtr, 2)) & cQuote & "," & _
                           .TextMatrix(nCtr, 3) & ")"
                QueryTemp cSqlStmt, objdbRs, True
            End If
        Next nCtr
    End With
    
    ShowProgress 3
    
    QueryTemp "SELECT * FROM TMP87260 WHERE EMPID=" & cQuote & Text2(0).Text & cQuote, objdbRs, False
    If objdbRs.RecordCount > 0 Then
        GenerateReport "Payslip Preview", "PRV87260.RPT", , True
    End If
    
    ShowProgress 4
End Sub

Private Sub Command7_Click()
    Dim nCtr As Integer
    
    DoEvents
        
    nAdd = 1
    
    ClearAll Me, True, True
    CtrlPanel Me, nAdd
    
    Text1.Text = ""
    Text2(2).Text = ""
    Text2(3).Text = ""
    Label1.Caption = ""
    Label27.Caption = ""
    
    Text1.Enabled = True
    Text2(0).Enabled = True
    Text2(1).Enabled = False
    Text2(4).Enabled = False
    Text2(5).Enabled = False
    Text2(6).Enabled = False

    Command1.Enabled = True
    Command2.Enabled = True
    
    MSHFlexGrid1.Enabled = False
    
    Combo1.Enabled = False
    Text10.Enabled = False
    Check1.Visible = True
    Check2.Enabled = True
    Check2.Value = vbUnchecked
    
    For nCtr = 0 To 8
        Text9(nCtr).Enabled = False
    Next nCtr
    
    Text2(0).SetFocus
End Sub

Private Sub Command8_Click()
    Dim nCtr As Integer
    
    DoEvents
    If Not isDataLock(Me.Name, "EMPID", Text2(0).Text) Then
        Lock2User Me.Name, "EMPID", Text2(0).Text, True
        nAdd = 2
        
        ClearAll Me, True, False
        CtrlPanel Me, nAdd
        
        Text1.Enabled = False
        Text2(0).Enabled = False
        Text2(1).Enabled = False
        Text2(4).Enabled = False
        Text2(5).Enabled = False
        Text2(6).Enabled = False

        Command1.Enabled = False
        Command2.Enabled = False
        
        MSHFlexGrid1.Enabled = False
        
        Combo1.Enabled = False
        Text10.Enabled = False
        Check1.Visible = True
        Check2.Enabled = True
        Check2.Value = vbUnchecked
        
        MSHFlexGrid2.FixedCols = 3
        
        For nCtr = 0 To 10
            Text9(nCtr).Enabled = False
        Next nCtr
        
        Text5(0).SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            SendKeys vbTab
            
        Case vbKeyUp
            SendKeys "+{TAB}"   ' --> previous control
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Log2Audit Name, "OPEN"
    
    myArray = Array("TXT:[Emp ID]:8:True", _
                    "TXT:[Employee Name]:30:True", _
                    "TXT:[First Name]:20:False", _
                    "TXT:[Last Name]:20:False")

    myArray2 = Array("TXT:[Ded ID]:3:False", _
                     "TXT:[Name]:24:True", _
                     "NUM:[Amount]:11.2:True", _
                     "NUM:[Amount2]:10.2:True", _
                     "NUM:[AUTO]:1:False", _
                     "NUM:[TAG]:1:False", _
                     "NUM:[TYPE]:1:False", _
                     "NUM:[Amount3]:10.2:True", _
                     "TXT:[Ctrl No]:10:False")
    
    Tag = nAccess_Tag
    nAdd = 0
    
    Combo1.ListIndex = 0
    Label27.Caption = ""
    Check2.Enabled = False
    Check2.Value = vbUnchecked
    
    ClearAll Me, False, True
    CtrlPanel Me, nAdd
    
    FillGrid
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
    OpenQueryDNS "SELECT * FROM PA87260 WHERE PERIODID=" & cQuote & lblPeriod & cQuote & _
                 " AND EMPID=" & cQuote & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & cQuote, oTempADO, False
    If oTempADO.RecordCount > 0 Then
        GetFields Me, oTempADO
        
        cSqlStmt = "select BACCNTNO from di3670 where empid=" & cQuote & oTempADO("empid") & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        Text2(10).Text = IIf(objdbRs.RecordCount > 0, Trim(objdbRs("BACCNTNO")), "")
        
        cSqlStmt = "select ifnull(a.posname,'') as position from di7670 a " & _
                   " left join di3670 b on a.posid=b.posid where b.empid=" & cQuote & oTempADO("empid") & cQuote
        OpenQueryDNS cSqlStmt, objdbRs, False
        
        Text2(2).Text = IIf(objdbRs.RecordCount > 0, objdbRs("POSITION"), "")
        Text2(3).Text = IIf(oTempADO("emp_stat") = 0, "WAP", IIf(oTempADO("emp_stat") = 1, "Contractual", "Regular"))
        Label27.Caption = IIf(oTempADO("active") = 0, "", IIf(oTempADO("active") = 1, "Resigned", IIf(oTempADO("active") = 2, "Finished Contract", "Terminated")))
        
        If (oTempADO("emp_stat") = 1) Or (oTempADO("emp_stat") = 2) Then
            If oTempADO("paystatus") <> 2 Then
        
                cSqlStmt = "SELECT A.DEDID, " & _
                           " IFNULL(B.DEDNAME,'Undefined Deduction') as DEDNAME, " & _
                           " A.DED_AMT, A.DED_AMT2, A.COMPUTED, B.DEDTAG, B.DEDTYPE, A.DED_AMT3, A.CTRL_NO " & _
                           "FROM PA87263 A LEFT JOIN PA3330 B ON A.DEDID=B.DEDID " & _
                           "WHERE A.PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                           " AND A.EMPID=" & cQuote & Text2(0).Text & cQuote & _
                           " ORDER BY A.DEDID"
            Else
                cSqlStmt = "SELECT A.DEDID, " & _
                           " IFNULL(B.DEDNAME,'Undefined Deduction') as DEDNAME, " & _
                           " A.DED_AMT, A.DED_AMT2, A.COMPUTED, B.DEDTAG, B.DEDTYPE, A.DED_AMT3, A.CTRL_NO " & _
                           "FROM PA87263 A LEFT JOIN PA3330 B ON A.DEDID=B.DEDID " & _
                           "WHERE A.PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                           " AND A.EMPID=" & cQuote & Text2(0).Text & cQuote & _
                           " AND A.DEDID NOT IN(" & cQuote & "001" & cQuote & "," & cQuote & "002" & cQuote & "," & cQuote & "003" & cQuote & "," & cQuote & "004" & cQuote & "," & cQuote & "005" & cQuote & "," & cQuote & "006" & cQuote & IIf(gCompanyID = "003", "," & cQuote & "007" & cQuote, "") & IIf(gCompanyID = "003", "," & cQuote & "013" & cQuote, "") & ") " & _
                           " ORDER BY A.DEDID"
            End If
        Else
            cSqlStmt = "SELECT A.DEDID, " & _
                       " IFNULL(B.DEDNAME,'Undefined Deduction') as DEDNAME, " & _
                       " A.DED_AMT, A.DED_AMT2, A.COMPUTED, B.DEDTAG, B.DEDTYPE, A.DED_AMT3, A.CTRL_NO " & _
                       "FROM PA87263 A LEFT JOIN PA3330 B ON A.DEDID=B.DEDID " & _
                       "WHERE A.PERIODID=" & cQuote & lblPeriod.Caption & cQuote & _
                       " AND A.EMPID=" & cQuote & Text2(0).Text & cQuote & _
                       " AND A.DEDID NOT IN(" & cQuote & "001" & cQuote & "," & cQuote & "002" & cQuote & "," & cQuote & "003" & cQuote & "," & cQuote & "004" & cQuote & "," & cQuote & "005" & cQuote & "," & cQuote & "006" & cQuote & IIf(gCompanyID = "003", "," & cQuote & "007" & cQuote, "") & IIf(gCompanyID = "003", "," & cQuote & "013" & cQuote, "") & ") " & _
                       " ORDER BY A.DEDID"
        End If
                   
                   
        OpenQueryDNS cSqlStmt, objdbRs, False
        If objdbRs.RecordCount > 0 Then
            QueryAttach objdbRs, MSHFlexGrid2, myArray2, False, , , 1
        Else
            SetGridColumn myArray2, MSHFlexGrid2
        End If
    Else
        Text2(2).Text = ""
        Text2(3).Text = ""
        Label27.Caption = ""
        SetGridColumn myArray2, MSHFlexGrid2
    End If
    
    If nAdd <> 0 Then Compute
'    If MSHFlexGrid1.Enabled Then MSHFlexGrid1.SetFocus
End Sub

Private Sub MSHFlexGrid1_GotFocus()
    MSHFlexGrid1_EnterCell
End Sub

Private Sub MSHFlexGrid2_DblClick()
    MSHFlexGrid2_KeyDown vbKeyReturn, 0
End Sub

Private Sub MSHFlexGrid2_GotFocus()
    If nAdd <> 0 Then KeyPreview = False
End Sub

Private Sub MSHFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd = 0 Then Exit Sub
    
    If oTempADO("wap") = 1 Then Exit Sub
    
    With MSHFlexGrid2
        Select Case KeyCode
            Case vbKeyReturn
                Command11.Cancel = False
                txtFlex.ZOrder 0
                txtFlex.Visible = True
                txtFlex.Width = .CellWidth + 25
                txtFlex.Height = .CellHeight
                txtFlex.left = .CellLeft + .left
                txtFlex.top = .CellTop + .top - 10
                txtFlex.Text = .Text
                txtFlex.SelStart = 0
                txtFlex.SelLength = Len(.Text)
                txtFlex.SetFocus
        End Select
    End With
End Sub

Private Sub MSHFlexGrid2_LostFocus()
    On Error Resume Next
    If Screen.ActiveForm.Name = Me.Name Then KeyPreview = Screen.ActiveForm.ActiveControl.Name <> "txtFlex"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If nAdd <> 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Text1.Text) = "" Then
                Command1_Click
            Else
                OpenQueryDNS "SELECT * FROM DI5463 WHERE lineid=" & cQuote & Text1.Text & cQuote, objdbRs, False
                If objdbRs.RecordCount > 0 Then
                    Label1.Caption = objdbRs("LineName")
                    FillGrid
                Else
                    Label1.Caption = ""
                    If Text1.Enabled Then
                        MsgBox "Department ID not found!", vbCritical, App.Title
                        Text1.SetFocus
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Text10_Change()
    With MSHFlexGrid1
        DoEvents
        .Redraw = False
        .Row = 1
        Do While .Row < .Rows - 1 And _
                 UCase(left(.TextMatrix(.Row, Combo1.ListIndex + 1), Len(Trim(Text10.Text)))) <> UCase(Trim(Text10.Text))
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
    End With
    MSHFlexGrid1_EnterCell
End Sub

Private Sub Text5_Change(Index As Integer)
'    If nAdd <> 0 Then Compute
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    If nAdd = 1 Then If Val(Text5(Index).Text) = 0 Then Text5(Index).Text = 0
    
    Text5(Index).SelStart = 0
    Text5(Index).SelLength = Len(Text5(Index).Text)
End Sub

Private Sub Text5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If nAdd <> 0 Then Compute
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
    If nAdd <> 0 Then
        Compute
        Cancel = Not IsNumeric(Text5(Index).Text)
    End If
End Sub
Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    With MSHFlexGrid2
        Select Case KeyCode
            Case vbKeyReturn
                .TextMatrix(.Row, 3) = txtFlex.Text
                
                ' --> addendum for 0 value, 20070705
                If InStr(1, "001,003,005", .TextMatrix(.Row, 1)) > 0 Then
                    ' --> added 20070907
                    If .TextMatrix(.Row, 1) = "003" Then
                        .TextMatrix(.Row, 4) = txtFlex.Text
                    End If
                    If Val(txtFlex.Text) = 0 Then .TextMatrix(.Row, 4) = 0
                End If
                
                Compute2
                
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

Private Sub txtFlex_Validate(Cancel As Boolean)
    If nAdd <> 0 Then Cancel = Not IsNumeric(txtFlex.Text)
End Sub

