VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "Toolbar2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4668
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4668
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Height          =   492
      Left            =   0
      Top             =   0
      Width           =   8052
      _ExtentX        =   14203
      _ExtentY        =   868
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DoubleTopBorder =   -1  'True
      DoubleBottomBorder=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   1
      CaptionOptions  =   0
      ShowSeparators  =   -1  'True
      ButtonKey1      =   "BACK"
      ButtonToolTipText1=   "BACK"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)

End Sub
