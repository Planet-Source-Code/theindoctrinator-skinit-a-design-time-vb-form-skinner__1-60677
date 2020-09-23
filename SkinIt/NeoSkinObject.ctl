VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl NeoSkinIt 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BackStyle       =   0  'Transparent
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   765
   ScaleWidth      =   750
   Tag             =   "SkinObject"
   ToolboxBitmap   =   "NeoSkinObject.ctx":0000
   Begin VB.PictureBox pic_LeftCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   1200
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox pic_DownBorder 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   720
      ScaleHeight     =   150
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic_RightBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pic_Borders 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic_LeftBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pic_RightCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2400
      ScaleHeight     =   720
      ScaleWidth      =   1440
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      Begin VB.Image img_MinimizeBtn 
         Height          =   300
         Left            =   540
         ToolTipText     =   "˜æ˜ÊÑíä"
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_RestoreBtn 
         Height          =   300
         Left            =   270
         ToolTipText     =   "ÈÇÒÔÊ"
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_CloseBtn 
         Height          =   300
         Left            =   0
         ToolTipText     =   "ÈÓÊä"
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox pic_CenterCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1200
      ScaleHeight     =   720
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lbl_Caption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList iml_Skin 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "ÎØ ÇÚáÇã ÎØÇí ÈÑäÇãå"
      Top             =   0
      Width           =   15
   End
   Begin VB.Image img_Logo 
      Height          =   750
      Left            =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "NeoSkinIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const SRCCOPY = &HCC0020
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000

'Const DefMaximizeBtn = 1
Const DefMinimizeBtn = 1
Const DefCaption = "Neo Caption"
Const DefLable = ""
Const DefBackColor = 0
Const DefForeColor = 0
Const DefCaptionTop = 195
Const DefCaptionColor = 0
Const DefLableColor = 0
Const DefSkinNo = 0

'Dim v_bMaximizeBtn As Boolean
Dim v_bMinimizeBtn As Boolean
Dim v_sCaption As String
Dim v_sLable As String
Dim v_sSkinPath As String
Dim v_oBackColor As OLE_COLOR
Dim v_oForeColor As OLE_COLOR
Dim v_iCaptionTop As Integer
Dim v_oCaptionColor As OLE_COLOR
Dim v_oLableColor As OLE_COLOR
Dim v_iMouseX, v_iMouseY As Integer
Dim v_oForm As Form
Dim v_sSkinNo As Integer
Dim v_oFormResize As Boolean

Public Enum NeoSkinMode
    NeoDefault = 0
    NeoTitanium = 1
    NeoBlue = 2
    NeoDeco = 3
    NeoHolograph = 4
    NeoTreasureChest = 5
    NeoALPI = 6
    NeoDoesntSuck = 7
    NeoSteelBlade = 8
    NeoWazoo = 9
    NeoSteelRain = 10
    NeoCoupe = 11
    NeoBoilerRoom = 12
    NeoExecutive = 13
    NeoWeaponx = 14
    NeoWinXP = 15
End Enum

Event Click()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Sub LoadSkin() '(m_Form As Form)
Dim v_iCenterImgFrequency As Integer
Dim v_iLoop As Integer
Dim v_lRtn As Long
Dim FormObject As Object
On Error GoTo Err_Find
    UserControl.Width = v_oForm.Width
    UserControl.Height = v_oForm.Height
    v_oForm.MousePointer = vbHourglass
'    SkinMode = v_sSkinNo
    With UserControl


        .lblMessage.Width = .Width
        
        .lblMessage.Top = .Height - .lblMessage.Height - 150 ' 150 is down Border Height
        
        v_oForeColor = v_oForm.ForeColor
        v_oBackColor = v_oForm.BackColor
        .BackColor = v_oBackColor
        .ForeColor = v_oForm.ForeColor

'        .pic_LeftCaption.Refresh

        .pic_CenterCaption.Left = .pic_LeftCaption.Width
        .lbl_Caption.Width = .pic_CenterCaption.Width

'        .pic_LeftCaption.Refresh
'        .pic_RightCaption.Refresh
        .pic_RightCaption.Left = .Width - .pic_RightCaption.Width

'        .pic_LeftCaption.Refresh
'        .pic_RightCaption.Refresh
'        .pic_CenterCaption.Refresh
        .pic_CenterCaption.Width = .Width - .pic_LeftCaption.Width - .pic_RightCaption.Width

        .img_CloseBtn.Left = .pic_RightCaption.Width - .img_CloseBtn.Width - 75

        .img_RestoreBtn.Left = .pic_RightCaption.Width - .img_RestoreBtn.Width - .img_CloseBtn.Width - 75

        .img_MinimizeBtn.Left = .pic_RightCaption.Width - .img_MinimizeBtn.Width - .img_CloseBtn.Width - 75

'        .pic_LeftBorder.Cls
        .pic_LeftBorder.Top = .pic_LeftCaption.Height
        .pic_LeftBorder.Height = .Height - .pic_LeftCaption.Height
'        .pic_RightBorder.Cls
'        .pic_LeftCaption.Refresh
'        .pic_RightCaption.Refresh
'        .pic_CenterCaption.Refresh
'        .pic_RightBorder.Refresh
        .pic_RightBorder.Left = .Width - 150
        .pic_RightBorder.Top = .pic_RightCaption.Height
        .pic_RightBorder.Height = v_oForm.Height - .pic_RightCaption.Height
        
'        .pic_LeftCaption.Refresh
'        .pic_RightCaption.Refresh
'        .pic_CenterCaption.Refresh
'        .pic_RightBorder.Refresh
'        .pic_LeftBorder.Refresh
'        .pic_RightBorder.Refresh

'        .pic_DownBorder.Cls
        .pic_DownBorder.Top = UserControl.Height - 150
        .pic_DownBorder.Width = UserControl.Width

'        .pic_LeftCaption.Refresh
'        .pic_RightCaption.Refresh
'        .pic_CenterCaption.Refresh
'        .pic_RightBorder.Refresh
'        .pic_LeftBorder.Refresh
'        .pic_RightBorder.Refresh
'        .pic_DownBorder.Refresh
        v_iCenterImgFrequency = Abs((.pic_CenterCaption.Width / Screen.TwipsPerPixelX) / 50)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_CenterCaption.hDC, v_iLoop * 50, 0, 100, 48, .pic_CenterCaption.hDC, 0, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_iCenterImgFrequency = Abs(((.Height - .pic_LeftCaption.Height) / Screen.TwipsPerPixelY) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency - 1
                v_lRtn = BitBlt(.pic_LeftBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 0, 0, SRCCOPY)
                v_lRtn = BitBlt(.pic_RightBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 30, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 9)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 9, 0, 9, 10, .pic_Borders.hDC, 20, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 10, 10, .pic_Borders.hDC, 10, 0, SRCCOPY)
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (v_oForm.Width / Screen.TwipsPerPixelX) - 10, 0, 10, 10, .pic_Borders.hDC, 40, 0, SRCCOPY)

        .lbl_Caption.Top = CaptionTop
        .lbl_Caption.ForeColor = CaptionColor
        .lbl_Caption.Caption = v_sCaption
    End With
    v_oForm.MousePointer = vbDefault
Err_Find: Exit Sub
End Sub

Public Property Get MinimizeBtn() As Boolean
    MinimizeBtn = v_bMinimizeBtn
End Property

Public Property Let MinimizeBtn(ByVal m_MinimizeBtn As Boolean)
    v_bMinimizeBtn = m_MinimizeBtn
    PropertyChanged "Minimize"
End Property

Public Property Get Caption() As String
    Caption = v_sCaption
End Property

Public Property Let Caption(ByVal m_Caption As String)
    v_sCaption = m_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Lable() As String
    Lable = v_sLable
End Property

Public Property Let Lable(ByVal m_Lable As String)
    v_sLable = m_Lable
    PropertyChanged "Lable"
End Property

Public Property Get SkinPath() As String
    SkinPath = v_sSkinPath
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)
    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = v_oBackColor
End Property

Public Property Let BackColor(ByVal m_BackColor As OLE_COLOR)
    v_oBackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get CaptionTop() As Integer
    CaptionTop = v_iCaptionTop
End Property

Public Property Let CaptionTop(ByVal m_CaptionTop As Integer)
    v_iCaptionTop = m_CaptionTop
    PropertyChanged "CaptionTop"
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = v_oCaptionColor
End Property

Public Property Let SkinMode(ByVal m_SkinMode As NeoSkinMode)
Dim v_iCenterImgFrequency As Integer
Dim v_iLoop As Integer
Dim v_lRtn As Long
    v_sSkinNo = m_SkinMode
    Select Case v_sSkinNo
        Case NeoDefault
            v_sSkinPath = App.Path & "\Skins\Default"
            v_iCaptionTop = 360
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoTitanium
            v_sSkinPath = App.Path & "\Skins\Titanium"
            v_iCaptionTop = 195
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoBlue
            v_sSkinPath = App.Path & "\Skins\Blue"
            v_iCaptionTop = 250
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoDeco
            v_sSkinPath = App.Path & "\Skins\Deco"
            v_iCaptionTop = 300
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoHolograph
            v_sSkinPath = App.Path & "\Skins\Holograph"
            v_iCaptionTop = 285
            v_oCaptionColor = &HFFFFFF
            v_oLableColor = &HC0&
        Case NeoTreasureChest
            v_sSkinPath = App.Path & "\Skins\TreasureChest"
            v_iCaptionTop = 240
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoALPI
            v_sSkinPath = App.Path & "\Skins\ALPI"
            v_iCaptionTop = 135
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoDoesntSuck
            v_sSkinPath = App.Path & "\Skins\Doesnt_Suck"
            v_iCaptionTop = 270
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoSteelBlade
            v_sSkinPath = App.Path & "\Skins\SteelBlade"
            v_iCaptionTop = 405
            v_oCaptionColor = &HFFFFFF
            v_oLableColor = &HC0&
        Case NeoWazoo
            v_sSkinPath = App.Path & "\Skins\Wazoo"
            v_iCaptionTop = 375
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoSteelRain
            v_sSkinPath = App.Path & "\Skins\SteelRain"
            v_iCaptionTop = 250
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoCoupe
            v_sSkinPath = App.Path & "\Skins\Coupe"
            v_iCaptionTop = 180
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoBoilerRoom
            v_sSkinPath = App.Path & "\Skins\BoilerRoom"
            v_iCaptionTop = 255
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoExecutive
            v_sSkinPath = App.Path & "\Skins\Executive"
            v_iCaptionTop = 370
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoWeaponx
            v_sSkinPath = App.Path & "\Skins\Weaponx"
            v_iCaptionTop = 135
            v_oCaptionColor = &H0&
            v_oLableColor = &HC0&
        Case NeoWinXP
            v_sSkinPath = App.Path & "\Skins\WinXP"
            v_iCaptionTop = 120
            v_oCaptionColor = &HFFFFFF
            v_oLableColor = &HC0&
    End Select
    With UserControl
        .iml_Skin.ListImages.Add 1, , LoadPicture(v_sSkinPath & "\img_Caption_Left.bmp")
        .iml_Skin.ListImages.Add 2, , LoadPicture(v_sSkinPath & "\img_Caption_Center.bmp")
        .iml_Skin.ListImages.Add 3, , LoadPicture(v_sSkinPath & "\img_Caption_Right.bmp")
        .iml_Skin.ListImages.Add 4, , LoadPicture(v_sSkinPath & "\img_Button_Close.gif")
        .iml_Skin.ListImages.Add 5, , LoadPicture(v_sSkinPath & "\img_Button_Restore.gif")
        .iml_Skin.ListImages.Add 6, , LoadPicture(v_sSkinPath & "\img_Button_Minimize.gif")
        .iml_Skin.ListImages.Add 7, , LoadPicture(v_sSkinPath & "\img_Borders.bmp")
        
        .pic_CenterCaption.Picture = .iml_Skin.ListImages(2).Picture
        .pic_Borders.Picture = .iml_Skin.ListImages(7).Picture

        .pic_LeftCaption.Cls
        .pic_LeftCaption.Picture = .iml_Skin.ListImages(1).Picture

        .pic_RightCaption.Cls
        .pic_RightCaption.Picture = .iml_Skin.ListImages(3).Picture

        .pic_LeftCaption.Visible = True
        .pic_CenterCaption.Visible = True
        .pic_RightCaption.Visible = True
        .pic_LeftBorder.Visible = True
        .pic_RightBorder.Visible = True
        .pic_DownBorder.Visible = True
        .img_Logo.Visible = False
        .lblMessage.Left = .pic_LeftBorder.Width
        .pic_LeftCaption.Top = 0
        .pic_LeftCaption.Left = 0
        .pic_RightCaption.Top = 0


        .pic_CenterCaption.Top = 0

        .img_CloseBtn.Top = 45
        .img_RestoreBtn.Top = 45
        .img_MinimizeBtn.Top = 45
        .pic_DownBorder.Left = 0
        .pic_DownBorder.Height = 150


        .img_CloseBtn.Picture = .iml_Skin.ListImages(4).Picture

        .img_RestoreBtn.Picture = .iml_Skin.ListImages(5).Picture

        .img_MinimizeBtn.Picture = .iml_Skin.ListImages(6).Picture

        v_iCenterImgFrequency = Abs((.pic_CenterCaption.Width / Screen.TwipsPerPixelX) / 50)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_CenterCaption.hDC, v_iLoop * 50, 0, 100, 48, .pic_CenterCaption.hDC, 0, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_iCenterImgFrequency = Abs(((.Height - .pic_LeftCaption.Height) / Screen.TwipsPerPixelY) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency - 1
                v_lRtn = BitBlt(.pic_LeftBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 0, 0, SRCCOPY)
                v_lRtn = BitBlt(.pic_RightBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 30, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 9)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 9, 0, 9, 10, .pic_Borders.hDC, 20, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 10, 10, .pic_Borders.hDC, 10, 0, SRCCOPY)
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (v_oForm.Width / Screen.TwipsPerPixelX) - 10, 0, 10, 10, .pic_Borders.hDC, 40, 0, SRCCOPY)
        .pic_LeftCaption.Refresh
        .pic_RightCaption.Refresh
        .pic_CenterCaption.Refresh
        .pic_RightBorder.Refresh
        .pic_LeftBorder.Refresh
        .pic_RightBorder.Refresh
        .pic_DownBorder.Refresh

    End With
    Call LoadSkin
    PropertyChanged "SkinMode"
End Property

Public Property Get SkinMode() As NeoSkinMode
    SkinMode = v_sSkinNo
End Property

Public Property Let CaptionColor(ByVal m_CaptionColor As OLE_COLOR)
    v_oCaptionColor = m_CaptionColor
    PropertyChanged "CaptionColor"
End Property

Public Property Get LableColor() As OLE_COLOR
    LableColor = v_oLableColor
End Property

Public Property Let LableColor(ByVal m_LableColor As OLE_COLOR)
    v_oLableColor = m_LableColor
    PropertyChanged "LableColor"
End Property

Public Sub img_CloseBtn_Click()
    Unload Screen.ActiveForm
End Sub


Public Sub img_MinimizeBtn_Click()
    Screen.ActiveForm.WindowState = 1
End Sub

Public Sub img_RestoreBtn_Click()
    Screen.ActiveForm.WindowState = 0
    UserControl.img_RestoreBtn.Visible = False
'    Call LoadSkin
End Sub

Public Sub lbl_Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        v_iMouseX = X
        v_iMouseY = Y
    End If
End Sub

Public Sub lbl_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (v_oForm.WindowState <> 2) Then
        Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
        Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
    End If
End Sub

Public Sub pic_CenterCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        v_iMouseX = X
        v_iMouseY = Y
    End If
End Sub

Public Sub pic_CenterCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (v_oForm.WindowState <> 2) Then
        Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
        Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Call UserControl_Resize
End Sub

Public Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    DoButtonsFlat v_oForm
End Sub

Public Sub UserControl_InitProperties()
    v_bMinimizeBtn = DefMinimizeBtn
    v_sCaption = DefCaption
    v_sLable = DefLable
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oBackColor = DefBackColor
    v_oForeColor = DefForeColor
    v_oCaptionColor = DefCaptionColor
    v_oLableColor = DefLableColor
    Set v_oForm = UserControl.ParentControls.Item(0)
    v_sSkinNo = DefSkinNo
    SkinMode = v_sSkinNo
    
End Sub

Public Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.img_RestoreBtn.Visible = False
    UserControl.img_MinimizeBtn.Left = UserControl.pic_RightCaption.Width - UserControl.img_MinimizeBtn.Width - UserControl.img_CloseBtn.Width - 75

    v_bMinimizeBtn = PropBag.ReadProperty("MinimizeBtn", DefMinimizeBtn)
    If v_bMinimizeBtn = False Then
        UserControl.img_MinimizeBtn.Visible = False
    Else
        UserControl.img_MinimizeBtn.Visible = True
    End If
    
    v_sCaption = PropBag.ReadProperty("Caption", DefCaption)
    UserControl.lbl_Caption.Caption = v_sCaption

    v_sLable = PropBag.ReadProperty("Lable", DefLable)
    UserControl.lblMessage.Caption = v_sLable

    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Caption.ForeColor = v_oForeColor

    v_iCaptionTop = PropBag.ReadProperty("CaptionTop", DefCaptionTop)
    UserControl.lbl_Caption.Top = v_iCaptionTop

    v_oCaptionColor = PropBag.ReadProperty("CaptionColor", DefCaptionColor)
    UserControl.lbl_Caption.ForeColor = v_oCaptionColor

    v_oLableColor = PropBag.ReadProperty("LableColor", DefLableColor)
    UserControl.lblMessage.ForeColor = v_oLableColor

    v_sSkinNo = PropBag.ReadProperty("SkinMode", DefSkinNo)
    Set v_oForm = UserControl.ParentControls.Item(0)
    SkinMode = v_sSkinNo
End Sub

Private Sub UserControl_Resize()
Dim befWidth As Single
Dim befHeight As Single
Dim UserControlObject As Object
On Error Resume Next
    Set v_oForm = UserControl.ParentControls.Item(0)
    Set UserControlObject = FindItem
    UserControlObject.Align = 1
    UserControlObject.Left = 0
    UserControlObject.Top = 0
    v_oForm.BorderStyle = 0
    befWidth = UserControl.Width
    befHeight = UserControl.Height
    UserControl.Width = v_oForm.Width
    UserControl.Height = v_oForm.Height
    Call LoadSkin
    If Err.Number <> 0 Then
        UserControl.Width = befWidth
        UserControl.Height = befHeight
        Call LoadSkin
    End If
End Sub
Public Sub CallResize()
        UserControl_Resize
End Sub
Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("MaximizeBtn", v_bMaximizeBtn, DefMaximizeBtn)
    Call PropBag.WriteProperty("MinimizeBtn", v_bMinimizeBtn, DefMinimizeBtn)
    Call PropBag.WriteProperty("Caption", v_sCaption, DefCaption)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("CaptionTop", v_iCaptionTop, DefCaptionTop)
    Call PropBag.WriteProperty("CaptionColor", v_oCaptionColor, DefCaptionColor)
    Call PropBag.WriteProperty("LableColor", v_oLableColor, DefLableColor)
    Call PropBag.WriteProperty("SkinMode", v_sSkinNo, DefSkinNo)
     
'    Call UserControl_Resize
'    v_oForm.Controls.Item(UserControl. Align = 1
End Sub
Private Function FindItem() As Object
    Dim ObjectCont As Object
    For Each ObjectCont In v_oForm
        If TypeName(ObjectCont) = "SkinObject" Then
            Set FindItem = ObjectCont
            Exit For
        End If
    Next
End Function
Public Sub DoButtonsFlat(Container As Form, Optional AllButtons As Boolean = True, Optional cmdName As CommandButton)
Dim Button
Dim Value As Boolean


If AllButtons = True Then
        For Each Button In Container.Controls
            If TypeOf Button Is CommandButton Or TypeOf Button Is Frame Then
                
                SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
                Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
            End If
        Next
ElseIf AllButtons = False And Not TypeOf cmdName Is CommandButton Then
         MsgBox "Please Provide a valid commandbutton name for making it flat"
         Exit Sub
Else
                SetWindowLong cmdName.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
                cmdName.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End If


End Sub
