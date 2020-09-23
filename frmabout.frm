VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Netab - By Elucid Software Inc."
   ClientHeight    =   4368
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   6012
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4368
   ScaleWidth      =   6012
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1092
      Left            =   4200
      ScaleHeight     =   1092
      ScaleWidth      =   1692
      TabIndex        =   5
      Top             =   120
      Width           =   1692
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Credits:"
         Height          =   252
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1572
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Elucid Software Crew:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Eric Malamisura"
         ForeColor       =   &H000000C0&
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Andy Minnich"
         ForeColor       =   &H000000C0&
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1572
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   372
      Left            =   4440
      TabIndex        =   4
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Our Website At:"
      Height          =   252
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "http://elucidsoftware.hypermart.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   2160
      MouseIcon       =   "frmabout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   2532
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmabout.frx":030A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1812
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmabout.frx":04E0
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   1236
      Left            =   120
      Picture         =   "frmabout.frx":0595
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Image1.Width = Image1.Width * Screen.TwipsPerPixelX
'Image1.Height = Image1.Height * Screen.TwipsPerPixelY
'Picture1.Width = Picture1.Width * Screen.TwipsPerPixelX
'Picture1.Height = Picture1.Height * Screen.TwipsPerPixelY
'Label3.Width = Label3.Width * Screen.TwipsPerPixelX
'Label3.Height = Label3.Height * Screen.TwipsPerPixelY
ResizeControls Me, 0
End Sub

Private Sub Label4_Click()
OpenIt Me, "http://elucidsoftware.hypermart.net"

End Sub
