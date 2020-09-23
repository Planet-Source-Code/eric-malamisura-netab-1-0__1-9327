VERSION 5.00
Object = "{F1909D6D-FB9D-11D3-B06C-00500427A693}#1.0#0"; "xuiTreeView6.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "First Time Running!  Netab Browser By Elucid Software"
   ClientHeight    =   6528
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6528
   ScaleWidth      =   7212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox contPopuplst 
      Height          =   5412
      Left            =   2520
      ScaleHeight     =   5364
      ScaleWidth      =   4644
      TabIndex        =   12
      Top             =   0
      Width           =   4692
      Begin VB.CheckBox Check6 
         Caption         =   "Use Caption List"
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   1572
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Use URL List"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   1332
      End
      Begin VB.Frame Frame4 
         Caption         =   "Block Caption List"
         Height          =   2172
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   4452
         Begin VB.ListBox List2 
            Height          =   1392
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   4212
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Update List"
            Height          =   252
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1332
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Remove"
            Height          =   252
            Left            =   3360
            TabIndex        =   20
            Top             =   1800
            Width           =   972
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Add"
            Height          =   252
            Left            =   2400
            TabIndex        =   21
            Top             =   1800
            Width           =   972
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Edit"
            Height          =   252
            Left            =   1560
            TabIndex        =   22
            Top             =   1800
            Width           =   852
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Block URL List"
         Height          =   2172
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4452
         Begin VB.CommandButton Command4 
            Caption         =   "Update List"
            Height          =   252
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   1332
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Remove"
            Height          =   252
            Left            =   3360
            TabIndex        =   15
            Top             =   1800
            Width           =   972
         End
         Begin VB.ListBox List1 
            Height          =   1392
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4212
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Add"
            Height          =   252
            Left            =   2400
            TabIndex        =   16
            Top             =   1800
            Width           =   972
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Edit"
            Height          =   252
            Left            =   1560
            TabIndex        =   18
            Top             =   1800
            Width           =   852
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSettings.frx":0000
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   1680
         TabIndex        =   27
         Top             =   4680
         Width           =   2892
      End
   End
   Begin VB.PictureBox contStartup 
      Height          =   492
      Left            =   6360
      ScaleHeight     =   444
      ScaleWidth      =   804
      TabIndex        =   1
      Top             =   6000
      Width           =   852
      Begin VB.Frame Frame2 
         Caption         =   "Window"
         Height          =   1212
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   3132
         Begin VB.CheckBox Check4 
            Caption         =   "Tray Icon"
            Height          =   192
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1092
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Last Ran"
            Height          =   252
            Left            =   1680
            TabIndex        =   10
            Top             =   720
            Width           =   972
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Maximized"
            Height          =   252
            Left            =   1680
            TabIndex        =   9
            Top             =   480
            Width           =   1092
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Minimized"
            Height          =   252
            Left            =   1680
            TabIndex        =   8
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Settings"
         Height          =   1572
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4452
         Begin VB.CheckBox Check3 
            Caption         =   "Start With Pagemaster"
            Height          =   252
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   2292
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Internet Settings"
            Height          =   252
            Left            =   2760
            TabIndex        =   5
            Top             =   1200
            Width           =   1572
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Open IE Homepage"
            Height          =   252
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1812
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check If Default Browser"
            Height          =   252
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   2652
         End
      End
   End
   Begin xuiTreeView6.TreeView TreeView1 
      Height          =   5412
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2412
      _ExtentX        =   4255
      _ExtentY        =   9546
      Lines           =   0   'False
      LabelEditing    =   0   'False
      PlusMinus       =   0   'False
      RootLines       =   0   'False
      ToolTips        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxScrollTime   =   0
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

