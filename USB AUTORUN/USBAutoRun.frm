VERSION 5.00
Begin VB.Form USBAutoRunFrm 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " USB AutoRunner"
   ClientHeight    =   585
   ClientLeft      =   5040
   ClientTop       =   4590
   ClientWidth     =   3060
   Icon            =   "USBAutoRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   3060
   Begin VB.CommandButton cmdMisc 
      BackColor       =   &H000000FF&
      Caption         =   "MIN WIN"
      Height          =   375
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "MINMISE THIS WINDOW"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox CheStart 
      BackColor       =   &H000000FF&
      Caption         =   "CHECK TO START WITH WINDOWS"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "CLICK TO ADD THIS PROGRAM TO THE WINDOWS START UP FOLDER"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "USBAutoRunFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WaitForSingleObject Lib "kernel32" _
           (ByVal hHandle As Long, _
           ByVal dwMilliseconds As Long) As Long
 
         Private Declare Function FindWindow Lib "user32" _
           Alias "FindWindowA" _
           (ByVal lpClassName As String, _
           ByVal lpWindowName As String) As Long
 
         Private Declare Function PostMessage Lib "user32" _
           Alias "PostMessageA" _
           (ByVal hwnd As Long, _
           ByVal wMsg As Long, _
           ByVal wParam As Long, _
           ByVal lParam As Long) As Long
 
         Private Declare Function IsWindow Lib "user32" _
           (ByVal hwnd As Long) As Long
 
         'Constants used by the API functions
        Const WM_CLOSE = &H10
        Const INFINITE = &HFFFFFFFF

Const myerrfilepath = 75
Dim Drive1(26) As Boolean
Dim KnownDrive(26) As Boolean

Dim wShell As New IWshShell_Class
Dim wShortcut As IWshShortcut_Class

Private Sub cmdMisc_Click()
USBAutoRunFrm.WindowState = vbMinimized
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Form_Load()
 If App.PrevInstance Then End
 Call Openerz
 Call Form_Activate
 Drive1DotRefresh
 For i = 1 To 26
  KnownDrive(i) = Drive1(i)
 Next i
 Timer1.Enabled = True
 Load FrmSysTray
   Set FrmSysTray.FSys = Me
 
End Sub

Private Sub Timer1_Timer()
 On Error Resume Next
 Drive1DotRefresh
 For i = 4 To 26
  If Drive1(i) Then ' A USB Drive is currently inserted...
   If KnownDrive(i) = False Then ' Ahh - it just happened!
    KnownDrive(i) = True ' OK, so remember it's there.
    DriveLetter$ = Chr$(64 + i) & ":"
    AutoPlayFile$ = DriveLetter$ & "\AUTORUN.INF"
    If Dir$(AutoPlayFile$) <> "" Then
     Open AutoPlayFile$ For Input As #1
     While Not EOF(1)
      Line Input #1, A$
      If InStr(UCase$(A$), "OPEN=") Or InStr(UCase$(A$), "OPEN =") Then
        S = InStr(A$, "=")
        Program$ = Trim$(Right$(A$, Len(A$) - S))
        Program$ = DriveLetter$ & IIf(Left$(Program$, 1) <> "\", "\", "") & Program$
        Shell Program$, vbNormalFocus
        Close
        GoTo UpdateKnownDrive:
      End If
     Wend
     Close
    End If
    Exit Sub
   End If
   KnownDrive(i) = True
  End If
 Next i
UpdateKnownDrive:
 For i = 4 To 26
  KnownDrive(i) = Drive1(i)
 Next i
End Sub
Sub Drive1DotRefresh()
 On Error GoTo NoDrive:
 For i = 4 To 26
  DriveLetter$ = Chr$(64 + i)
  Drive1(i) = Dir$(DriveLetter$ & ":\AutoRun.inf") <> ""
NextDrive:
 Next i
 T = Timer + 0.1
 While T > Timer
  DoEvents
 Wend
 Exit Sub
NoDrive:
 Resume NextDrive:
End Sub
Private Sub CheStart_Click()
If CheStart.Value = 1 Then
        'savestring HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "USBAutoRun", NormalisePath(App.Path) & "USBAutoRun.exe -min"
    'Call AppAutoStart
    Call CreateIcon
    Text1(0).Text = "YES"
    Call uasSaver
    ElseIf CheStart.Value = 0 Then
        'DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "USBAutoRun"
     'Call AppAutoStop
     Text1(0).Text = "NO"
      Call uasSaver
    End If
End Sub
Private Sub Form_Resize()
' ResizeAll USBAutoRunFrm
'If (ChkHideOnMinimize) Then FrmSysTray.MeResize Me

   'If (Check1) Then
   FrmSysTray.MeResize Me

End Sub
Private Sub Form_Activate()
If Text1(0).Text = "YES" Then
 CheStart.Value = 1
 Call cmdMisc_Click
Else
Text1(0).Text = "NO"
CheStart.Value = 0
Me.WindowState = vbNormal
Me.Show
End If
End Sub
Private Sub uasSaver()
On Error GoTo snafufubar
  'saves the selected file'
  Dim msg As String
  Dim Filehandle As Integer
  Dim X As Integer

  Filehandle = FreeFile

   Open App.Path & "\formshow.uas" For Output As Filehandle
        
          Write #Filehandle, USBAutoRunFrm.Text1(0);

      
      Close #Filehandle

snafufubar:
      If (Err.Number = myerrfilepath) Then
        msg = "you must save a file"
        If MsgBox(msg) = vbOK Then
          USBAutoRunFrm.SetFocus
        End If
      End If
      Exit Sub
End Sub
Private Sub Openerz()
On Error GoTo fubar
  'opens the selected file'
  Dim msg As String
  Dim box1 As String

  Dim Filenumber As Integer

  Filenumber = FreeFile

    Open App.Path & "\formshow.uas" For Input As #Filenumber

        Do While Not EOF(Filenumber)
          Input #Filenumber, box1
          Text1(0).Text = box1
                Loop
      Close #Filenumber

      Exit Sub
fubar:
      If (Err.Number = myerrfilepath) Then
        msg = "you must select a file to open"
        If MsgBox(msg) = vbOK Then
          USBAutoRunFrm.SetFocus
        End If
      End If
      Exit Sub
End Sub
Private Sub CreateIcon()
'Now Add the following Code
   'This will Create an ICON under Start->Programs-> with the Name TEST
   'If U want UR Icon to be only under the START MENU then change wShell.SpecialFolders.Item(2) with wShell.SpecialFolders.Item(1) below
   'If U want UR Icon to be on the DESKTOP then change wShell.SpecialFolders.Item(2) with wShell.SpecialFolders.Item(0) below
   Set wShortcut = wShell.CreateShortcut("C:\Documents and Settings\All Users\Start Menu\Programs\Startup" & "\USB AutoRunner.lnk")
   'The Target Path is the Application name for which U want UR ICON to REFER TO.
   'Make sure U refer the Application Name and the Correct Path
   wShortcut.TargetPath = App.Path & "\USBAutorun.exe"
    wShortcut.IconLocation = App.Path & "\GOLD.ico"
    wShortcut.Save
End Sub
