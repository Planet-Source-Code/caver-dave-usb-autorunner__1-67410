VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   645
   ClientLeft      =   6450
   ClientTop       =   5655
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   630
      Picture         =   "FrmSys.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   540
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   15
      Picture         =   "FrmSys.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Interval        =   1000
      Left            =   1260
      Top             =   105
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mrestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'\|/declarations for any menus with bitmaps
   Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

   Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
     ByVal nPos As Long) As Long

   Private Declare Function SetMenuItemBitmaps Lib "user32" _
     (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
      ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

   Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
     (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
     ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 
' Constant for SetMenuItemBitmaps
    Private Const MF_BYPOSITION = &H400&
 
' Constants for LoadImage
    Private Const IMAGE_BITMAP = &O0
    Private Const LR_LOADFROMFILE = 16
    Private Const LR_CREATEDIBSECTION = 8192
'/|\declarations for any menus with bitmaps

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 12
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer

Public Property Let Tooltip(Value As String)
   'nid.szTip = Value & vbNullChar
   nid.szTip = "USB Auto Run"
End Property

Public Property Get Tooltip() As String
   Tooltip = nid.szTip
End Property

Public Property Let Interval(Value As Integer)
   TmrFlash.Interval = Value
   UpdateIcon NIM_MODIFY
End Property

Public Property Get Interval() As Integer
   Interval = TmrFlash.Interval
End Property

Public Property Let TrayIcon(Value)
   TmrFlash.Enabled = False
   On Error Resume Next
   ' Value can be a picturebox, image, form or string
   Select Case TypeName(Value)
      Case "PictureBox", "Image"
         Me.Icon = Value.Picture
         TmrFlash.Enabled = False
         RaiseEvent TIcon(Me)
      Case "String"
         If (UCase(Value) = "DEFAULT") Then
            TmrFlash.Enabled = True
            Me.Icon = Flash2.Picture
            RaiseEvent TIcon(Me)
         Else
            ' Sting is filename; load icon from picture file.
            TmrFlash.Enabled = True
            Me.Icon = LoadPicture(Value)
            RaiseEvent TIcon(Me)
         End If
      Case Else
         ' It's a form ?
         Me.Icon = Value.Icon
         RaiseEvent TIcon(Me)
   End Select
   If Err.Number <> 0 Then TmrFlash.Enabled = True
   UpdateIcon NIM_MODIFY
End Property

Private Sub Form_Load()
   Me.Icon = Flash1
   RaiseEvent TIcon(Me)
   Me.Visible = False
   TmrFlash.Enabled = True
   Tooltip = "USB Auto Run"
   mAbout.Caption = "About " & "USB AutoRunner"
   UpdateIcon NIM_ADD
   
Call MnuImg
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim msg As Long
   
   ' The Form_MouseMove is intercepted to give systray mouse events.
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If
      
   Select Case msg
      Case WM_RBUTTONDBLCLK
         RaiseEvent Click("RBUTTONDBLCLK")
      Case WM_RBUTTONDOWN
         RaiseEvent Click("RBUTTONDOWN")
      Case WM_RBUTTONUP
         ' Popup menu: selectively enable items dependent on context.
         Select Case FSys.Visible
            Case True
               Select Case FSys.WindowState
'                  Case vbMaximized
'                     mMaximize.Enabled = False
'                     mMinimize.Enabled = True
'                     mRestore.Enabled = False
'                  Case vbNormal
'                     mMaximize.Enabled = True
'                     mMinimize.Enabled = True
'                     mRestore.Enabled = False
'                  Case vbMinimized
'                     mMaximize.Enabled = True
'                     mMinimize.Enabled = False
'                     mRestore.Enabled = True
                  Case Else
'                     mMaximize.Enabled = True
'                     mMinimize.Enabled = True
                     mrestore.Enabled = True
               End Select
            Case Else
               mrestore.Enabled = True
'               mMaximize.Enabled = True
 '              mMinimize.Enabled = False
         End Select
         
         RaiseEvent Click("RBUTTONUP")
         PopupMenu mPopupMenu
      Case WM_LBUTTONDBLCLK
         RaiseEvent Click("LBUTTONDBLCLK")
         mRestore_Click
      Case WM_LBUTTONDOWN
         RaiseEvent Click("LBUTTONDOWN")
      Case WM_LBUTTONUP
         RaiseEvent Click("LBUTTONUP")
      Case WM_MBUTTONDBLCLK
         RaiseEvent Click("MBUTTONDBLCLK")
      Case WM_MBUTTONDOWN
         RaiseEvent Click("MBUTTONDOWN")
      Case WM_MBUTTONUP
         RaiseEvent Click("MBUTTONUP")
      Case WM_MOUSEMOVE
         RaiseEvent Click("MOUSEMOVE")
      Case Else
         RaiseEvent Click("OTHER....: " & Format$(msg))
   End Select
End Sub

Private Sub FSys_Resize()
   ' Event generated my main form. WindowState is stored in LastWindowState, so that
   ' it may be re- set when the menu item "Restore" is selected.
   If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState
End Sub

Private Sub FSys_Unload(Cancel As Integer)
   ' Important: remove icon from tray, and unload this form when
   ' the main form is unloaded.
   UpdateIcon NIM_DELETE
   Unload Me
End Sub

Private Sub mAbout_Click()
frmAbout.Show
End Sub
Public Sub mExit_Click()
   Unload FSys
   Unload USBAutoRunFrm
   End
End Sub

Private Sub mRestore_Click()
   ' Don't "restore"  FSys is visible and not minimized.
   If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub
   ' Restore LastWindowState
   FSys.WindowState = LastWindowState
   FSys.Visible = True
   SetForegroundWindow FSys.hwnd
End Sub

Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nid
      .cbSize = Len(nid)
      .hwnd = Me.hwnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
   End With
   Shell_NotifyIcon Value, nid
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      ' Cancel by setting Cancel = 1, minimize and hide main window.
      Cancel = 1
      F.WindowState = vbMinimized
      F.Hide
   End If
End Sub

Public Sub MeResize(ByRef F As Form)
   Select Case F.WindowState
      Case vbNormal, vbMaximized
         ' Store LastWindowState
         LastWindowState = F.WindowState
      Case vbMinimized
         F.Hide
   End Select
End Sub

Private Sub TmrFlash_Timer()
   ' Change icon.
   Static LastIconWasFlash1 As Boolean
   LastIconWasFlash1 = Not LastIconWasFlash1
   Select Case LastIconWasFlash1
      Case True
         Me.Icon = Flash2
      Case Else
         Me.Icon = Flash1
   End Select
   RaiseEvent TIcon(Me)
   UpdateIcon NIM_MODIFY
End Sub

Sub MnuImg()
'loads & creates menus with bitmap images
Dim hMenu, hMenu1, hMenu2, hMenu3 As Long
Dim hSubMenu, hSubMenu1, hSubMenu2, hSubMenu3 As Long
Dim hMenuImg, hMenuImg1, hMenuImg2, hMenuImg3 As Long
Dim sFileName, sFileName1, sFileName2, sFileName3 As String

 'Get the bitmap.
'popmenu
        sFileName = App.Path & "\menus\about.bmp"
        sFileName1 = App.Path & "\menus\restore.bmp"
        sFileName2 = App.Path & "\menus\exit.bmp"
        
'Menu1
     hMenuImg1 = LoadImage(0, sFileName, IMAGE_BITMAP, 10, 10, _
                LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
     hMenuImg2 = LoadImage(0, sFileName1, IMAGE_BITMAP, 10, 10, _
                LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
     hMenuImg3 = LoadImage(0, sFileName2, IMAGE_BITMAP, 10, 10, _
                LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
      


' Get the menu item handle.
'1 st Menu
      hMenu = GetMenu(Me.hwnd)
      hSubMenu = GetSubMenu(hMenu, 0) '0 is menu bar item 1'
      hMenu1 = GetMenu(Me.hwnd)
      hSubMenu1 = GetSubMenu(hMenu1, 0)
      hMenu2 = GetMenu(Me.hwnd)
      hSubMenu2 = GetSubMenu(hMenu2, 0)
      hMenu3 = GetMenu(Me.hwnd)
      hSubMenu3 = GetSubMenu(hMenu3, 0)
      

'1 st Menu
      Call SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, hMenuImg1, 0)
      Call SetMenuItemBitmaps(hSubMenu1, 2, MF_BYPOSITION, hMenuImg2, 0)
      Call SetMenuItemBitmaps(hSubMenu2, 4, MF_BYPOSITION, hMenuImg3, 0)
    
End Sub
