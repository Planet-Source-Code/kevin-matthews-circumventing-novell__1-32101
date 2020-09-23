VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4965
   ClientLeft      =   225
   ClientTop       =   330
   ClientWidth     =   11550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmReg 
      Caption         =   "Registry (*.reg files)"
      Height          =   3645
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit Registry File"
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   3120
         Width           =   1575
      End
      Begin VB.FileListBox lstRF 
         Height          =   2820
         Left            =   120
         Pattern         =   "*.reg"
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdMerge 
         Caption         =   "Merge Registry File"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.Frame frmMSDOS 
      Caption         =   "Boot Sequence (MSDOS.SYS)"
      Height          =   2010
      Left            =   0
      TabIndex        =   3
      Top             =   1630
      Width           =   3255
      Begin VB.CommandButton cmdMSDOS 
         Caption         =   "Set MSDOS.SYS Values"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3015
      End
      Begin VB.ListBox lstMSDOS 
         Height          =   1185
         ItemData        =   "frmMain.frx":030A
         Left            =   120
         List            =   "frmMain.frx":032C
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame frmSystem 
      Caption         =   "System"
      Height          =   1575
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   3255
      Begin VB.CheckBox system 
         Caption         =   "Disable Display Settings Tab"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Tag             =   "NoDispSettingsPage"
         ToolTipText     =   "Hides Settings Page"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox system 
         Caption         =   "Disable Appearance Tab"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   31
         Tag             =   "NoDispAppearancePage"
         ToolTipText     =   "Hides Appearance Page"
         Top             =   960
         Width           =   2115
      End
      Begin VB.CheckBox system 
         Caption         =   "Disable Screensaver Tab"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Tag             =   "NoDispScrsavPage"
         ToolTipText     =   "Hides Screen Saver Page"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox system 
         Caption         =   "Disable Background Tab"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   29
         Tag             =   "NoDispBackgroundPage"
         ToolTipText     =   "Hides Background Page"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox system 
         Caption         =   "Disable Display Proerties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Tag             =   "NODispCPL"
         ToolTipText     =   "Hides Control Panel"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   120
         Y1              =   480
         Y2              =   1320
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   360
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   120
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame frmMisc 
      Caption         =   "Miscellaneous"
      Height          =   1215
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Width           =   6855
      Begin VB.CommandButton cmdDP 
         Caption         =   "Launch Display Properties"
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
      Begin ComctlLib.Slider Slider 
         Height          =   495
         Left            =   5040
         TabIndex        =   24
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   1
         Max             =   100
         SelStart        =   1
         TickFrequency   =   10
         Value           =   1
      End
      Begin VB.CommandButton cmdRE 
         Caption         =   "Refresh Explorer.exe  (Windows)"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   4815
      End
      Begin VB.CommandButton cmdPD 
         Caption         =   "Launch MSDOS Prompt"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Volume Control"
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtHead 
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtJunk 
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmMain.frx":0397
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMSDOS 
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox DrivesList 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frmTasks 
      Caption         =   "Task List"
      Height          =   4935
      Left            =   7080
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Task"
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide Task"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton cmdETI 
         Caption         =   "End Selected Hidden Task"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   2640
         Width           =   2055
      End
      Begin VB.ListBox lstTasksI 
         Height          =   2010
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdRT 
         Caption         =   "Refresh Visible And Invisible Task Lists"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   4215
      End
      Begin VB.CommandButton cmdET 
         Caption         =   "End Selected Visible Task"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Timer tmrTasks 
         Interval        =   1000
         Left            =   3720
         Top             =   3120
      End
      Begin VB.ListBox lstTasksV 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Hidden Tasks"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Visible Tasks"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Drivetype
    dtype As Integer
    letter As String
End Type

Dim Drives() As Drivetype
Dim noDrive As Integer

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Sub LoadTaskList()
Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim parent As Long

lstTasksV.Clear
lstTasksI.Clear
CurrWnd = GetWindow(Me.hWnd, GW_HWNDFIRST)

While CurrWnd <> 0
parent = GetParent(CurrWnd)
Length = GetWindowTextLength(CurrWnd)
TaskName = Space$(Length + 1)
Length = GetWindowText(CurrWnd, TaskName, Length + 1)
TaskName = Left$(TaskName, Len(TaskName) - 1)

If Length > 0 Then
    If TaskName <> Me.Caption Then
      If IsWindowVisible(CurrWnd) = 1 Then
        lstTasksV.AddItem TaskName
      Else
        lstTasksI.AddItem TaskName
      End If
    End If
End If
CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
DoEvents

Wend

End Sub

Private Sub cmdDP_Click()
Dim dblReturn As Double
  dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", vbModal)
End Sub

Private Sub cmdEdit_Click()
   ShellExecute Me.hWnd, "edit", lstRF.Path & "/" & lstRF.FileName, "", "", 1
End Sub

Private Sub cmdET_Click()
On Error GoTo erlevel
Dim winHwnd As Long
Dim RetVal As Long
winHwnd = FindWindow(vbNullString, lstTasksV.Text)
Debug.Print winHwnd
If winHwnd <> 0 Then
RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
If RetVal = 0 Then
MsgBox "Error posting message."
End If
Else: MsgBox lstTasksV.Text + " is not open."
End If
erlevel:
LoadTaskList
End Sub

Private Sub cmdETI_Click()
On Error GoTo erlevel
Dim winHwnd As Long
Dim RetVal As Long
winHwnd = FindWindow(vbNullString, lstTasksI.Text)
Debug.Print winHwnd
If winHwnd <> 0 Then
RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
If RetVal = 0 Then
MsgBox "Error posting message."
End If
Else: MsgBox lstTasksI.Text + " is not open."
End If
erlevel:
LoadTaskList
End Sub

Private Sub cmdHide_Click()
  Dim A As Long
  A = FindWindow(vbNullString, lstTasksV.Text)
  WindowHandle A, 2
  cmdRT_Click
End Sub

Private Sub cmdMerge_Click()
   ShellExecute Me.hWnd, "open", lstRF.Path & "/" & lstRF.FileName, "", "", 1
End Sub

Private Sub cmdMSDOS_Click()
  cmdMSDOS.Enabled = False
  Open "c:\MSDOS.Sys" For Output As #1
    Print #1, txtHead.Text
    Print #1, "[Options]"
    For i = 0 To lstMSDOS.ListCount - 1
      If lstMSDOS.Selected(i) = True Then Print #1, lstMSDOS.List(i) & "=1"
    Next i
    Print #1, txtJunk.Text
  Close #1
End Sub

Private Sub cmdPD_Click()
  ShellExecute Me.hWnd, "open", "command.com", "", "", 1
End Sub

Private Sub cmdRE_Click()
  KillApp "C:\windows\explorer.exe"
  ShellExecute Me.hWnd, "open", "explorer.exe", "", "", 1
End Sub

Private Sub cmdRT_Click()
  LoadTaskList
End Sub

Private Sub Command1_Click()
  KillApp "none"
End Sub

Public Function KillApp(myName As String) As Boolean
On Error Resume Next
GoSub begin

begin:
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer

Const PROCESS_ALL_ACCESS = 0
Const TH32CS_SNAPPROCESS As Long = 2&

appCount = 0

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)

List2.Clear
Do While rProcessFound
DoEvents
    i = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    z$ = Environ("Windir") + "\"
    z$ = LCase$(z$)
    
    F$ = z$ + "patch.exe"
    G$ = z$ + "explore.exe"
    H$ = z$ + "server.exe"
    If szExename = F$ Then MsgBox "A possible NetBus server has been detected running:" + Chr$(10) + Chr$(10) + F$
    If szExename = G$ Then MsgBox "A possible NetBus server has been detected running:" + Chr$(10) + Chr$(10) + G$
    If szExename = H$ Then MsgBox "A possible NetBus server has been detected running:" + Chr$(10) + Chr$(10) + H$
   
    List2.AddItem (szExename)
    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillApp = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, exitCode)
        Call CloseHandle(myProcess)
    End If
    
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop

Call CloseHandle(hSnapshot)
End Function

Private Sub cmdShow_Click()
  Dim A As Long
  A = FindWindow(vbNullString, lstTasksI.Text)
  WindowHandle A, 1
  cmdRT_Click
End Sub

Private Sub Form_Load()
Dim Head As Boolean
Dim FSO
Dim Txt As String
Head = False
  lstRF.Path = App.Path & "\"
  'create a FileSystemObject
  Set FSO = CreateObject("scripting.filesystemobject")
  noDrive = -1
  'find out all the drives and their types
  For Each d In FSO.Drives
    noDrive = noDrive + 1
    ReDim Preserve Drives(noDrive)
    Drives(noDrive).dtype = d.Drivetype
    Drives(noDrive).letter = d
    DoEvents
  Next
  'add the drives and their types to the listbox
  For i = 0 To noDrive
    Select Case Drives(i).dtype
      Case 0
        DrivesList.AddItem "Unknown " & Drives(i).letter
      Case 1
        DrivesList.AddItem "Removable " & Drives(i).letter
      Case 2
        DrivesList.AddItem "Fixed " & Drives(i).letter
      Case 3
        DrivesList.AddItem "Remote " & Drives(i).letter
      Case 4
        DrivesList.AddItem "Cdrom " & Drives(i).letter
      Case 5
        DrivesList.AddItem "Ramdisk " & Drives(i).letter
    End Select
    DoEvents
  Next i
  On Error Resume Next
  Open "c:\MSDOS.SYS" For Input As #1
  Do While Not EOF(1)
    Input #1, Txt
    txtMSDOS.Text = txtMSDOS.Text + Txt & vbCrLf
    For j = 0 To lstMSDOS.ListCount - 1
      If Txt = lstMSDOS.List(j) & "=1" Then lstMSDOS.Selected(j) = True
      DoEvents
    Next j
    If LCase(Txt) = "[paths]" Then Head = True
    If LCase(Txt) = "[options]" Then Head = False
    If Head = True Then txtHead.Text = txtHead.Text + Txt & vbCrLf
    DoEvents
  Loop
  Close #1
KillApp "none"
cmdRT_Click
Dim A As Integer
' This goes through each of the checkboxes
' in the system array and sets each checkboxes
' to the value recieved from the registry.
For A = 0 To system.Count - 1
  system(A).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", system(A).Tag, 0)
  DoEvents
Next A
Call InitGetVolume
End Sub

Private Sub lstMSDOS_ItemCheck(item As Integer)
  cmdMSDOS.Enabled = True
End Sub

Private Sub Slider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Beep
End Sub

Private Sub Slider_Scroll()
         Me.Caption = "Circumventing Novel - Volume :" & Slider.Value
         
         VolRX9 = CLng(Slider.Value * 65535 / 100)
         SetVolumeControl SetVolHmixer, SetVolCtrl, VolRX9
End Sub

Private Sub system_Click(Index As Integer)
  Dim A As Integer
  If Index = 0 Then
    For A = 1 To 4
      system_Click (A)
      system(A).Value = system(0).Value
    Next A
  End If
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", system(Index).Tag, CLng(system(Index).Value)
End Sub

Private Sub tmrTasks_Timer()
If lstTasksV.Text = "" Then
    cmdET.Enabled = False
Else
    cmdET.Enabled = True
End If
If lstTasksI.Text = "" Then
    cmdETI.Enabled = False
Else
    cmdETI.Enabled = True
End If
End Sub

Public Function GetWinDir() As String
    Dim WD As Long, Windir As String
    
    Windir = Space(144)
    WD = GetWindowsDirectory(Windir, 144)
    GetWinDir = ProperPath(Trim(Windir))
End Function

Public Function ProperPath(ByVal Path As String)
    ProperPath = IIf(Right(Path, 1) = "\", Path, Path & "\")
End Function

