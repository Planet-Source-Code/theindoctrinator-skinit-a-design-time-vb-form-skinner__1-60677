VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.NeoSkinIt NeoSkinIt1 
      Height          =   5250
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9260
      SkinPath        =   "C:\Documents and Settings\srx\Desktop\SkinIt\Skins\Default"
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      CaptionTop      =   360
      LableColor      =   192
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000004&
      Height          =   675
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   480
      TabIndex        =   2
      Top             =   2745
      Width           =   2145
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1320
      Width           =   3375
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        'SkinObject1.DoButtonsFlat Me
End Sub

Private Sub SkinObject1_GotFocus()
        NeoSkinIt1.DoButtonsFlat Me
End Sub
