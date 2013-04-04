VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form frmPsWalker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PsWalker"
   ClientHeight    =   9390
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   10470
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   8880
   End
   Begin ComctlLib.ListView lsProcesses 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   16536
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PID"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Image Path"
         Object.Width           =   13229
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Started by"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   320
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Process"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshProcess 
         Caption         =   "Refresh Now"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKillProcess 
         Caption         =   "Kill Process"
      End
      Begin VB.Menu mnuKillTree 
         Caption         =   "Kill Process Tree"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuspend 
         Caption         =   "Suspend"
      End
      Begin VB.Menu mnuResume 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDump 
         Caption         =   "Dump"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuUpdateSpeed 
         Caption         =   "Update Speed"
         Begin VB.Menu mnuHalfSecs 
            Caption         =   "0.5 seconds"
         End
         Begin VB.Menu mnuOneSec 
            Caption         =   "1 second"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuTwoSecs 
            Caption         =   "2 seconds"
         End
         Begin VB.Menu mnuFiveSecs 
            Caption         =   "5 seconds"
         End
         Begin VB.Menu mnuBar2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPauseUpdate 
            Caption         =   "Paused"
         End
      End
   End
End
Attribute VB_Name = "frmPsWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = (DI_MASK Or DI_IMAGE)
Const IDI_APPLICATION As Long = 32512&

Sub InitProcessList(lsProcesses As ListView)
    Dim i As Long, j As Long, nBeenNull As Integer
    Dim dwPID As Long, szImagePath As String, IsAlive As Boolean, szParentName As String
    Dim b As Boolean
    
    Timer1.Enabled = False
    
    nBeenNull = 0
    lsProcesses.ListItems.Clear
    
    For i = 0& To 2147483647 Step 4
        If nBeenNull >= 32767 Then Exit For
        
        DoEvents
        
        IsAlive = IsProcessAlive(i)
        
        If (IsAlive) Then
            szImagePath = StandardizeProcessPath(GetProcessImagePath(i))
            nBeenNull = 0
            
            If szImagePath <> "" Then
                b = DrawFileIconInPictureBox(szImagePath, Picture1)
                If (b) Then ImageList1.ListImages.Add , , Picture1.Image
                
                lsProcesses.ListItems.Add , , GetFileNameByPath(szImagePath), , IIf(b, ImageList1.ListImages.Count, LoadPicture(""))
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(1) = i
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(2) = szImagePath
            Else
                b = DrawFileIconInPictureBox(szImagePath, Picture1)
                If (b) Then ImageList1.ListImages.Add , , Picture1.Image
                
                lsProcesses.ListItems.Add , , "???", , IIf(b, ImageList1.ListImages.Count, LoadPicture(""))
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(1) = i
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(2) = "[Error opening process]"
            End If
        Else
            nBeenNull = nBeenNull + 1
        End If
    Next
    
    For i = 1 To lsProcesses.ListItems.Count
        dwPID = lsProcesses.ListItems(i).SubItems(1)
        
        If PIDExistsInListView(lsProcesses, GetProcessParentPID(dwPID), j) Then
            szParentName = lsProcesses.ListItems(j).Text
        Else
            szParentName = "<Non-existent process>"
        End If
        
        lsProcesses.ListItems(i).SubItems(3) = szParentName & " (" & GetProcessParentPID(dwPID) & ")"
    Next
    
    Timer1.Enabled = True
End Sub

Function GetFileNameByPath(ByVal szFile As String) As String
    If InStr(1, szFile, "\") Then
        GetFileNameByPath = Right$(szFile, InStr(1, StrReverse(szFile), "\") - 1)
    Else
        GetFileNameByPath = szFile
    End If
End Function

Private Sub Form_Load()
    Call MiniEnablePrivilege
    Call InitProcessList(lsProcesses)
    SendMessageLong lsProcesses.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lsProcesses_ItemClick(ByVal Item As ComctlLib.ListItem)
    mnuKillProcess.Enabled = True
    mnuKillTree.Enabled = True
    mnuRestart.Enabled = True
    mnuSuspend.Enabled = True
    mnuResume.Enabled = True
    mnuDump.Enabled = True
End Sub

Private Sub lsProcesses_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        mnuKillProcess.Enabled = False
        mnuKillTree.Enabled = False
        mnuRestart.Enabled = False
        mnuSuspend.Enabled = False
        mnuResume.Enabled = False
        mnuDump.Enabled = False
    End If
End Sub

Private Sub lsProcesses_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnuProcess
End Sub

Private Sub mnuDump_Click()
    Dim ret As Long
    
    ret = DumpProcessMemoryToFile(lsProcesses.SelectedItem.SubItems(1), App.Path & "\" & lsProcesses.SelectedItem.Text & ".dmp")
    
    If ret = 0 Then
        Call MessageBox(Me.hWnd, "Process memory was dumped successfully.", "PsWalker", 0)
    Else
        Call MessageBox(Me.hWnd, "Failed to dump process memory." & vbCrLf & vbCrLf & "Error code: " & ret, "PsWalker", vbCritical)
    End If
End Sub

Private Sub mnuKillTree_Click()
    Dim dwParentPID As Long, dwPID As Long, i As Long, bFailed As Boolean
    
    dwParentPID = lsProcesses.SelectedItem.SubItems(1)
    dwPID = dwParentPID
    
LoopStart:
    For i = 1 To lsProcesses.ListItems.Count
        If GetProcessParentPID(lsProcesses.ListItems(i).SubItems(1)) = dwPID Then
            If IsProcessAlive(lsProcesses.ListItems(i).SubItems(1)) Then
                dwPID = lsProcesses.ListItems(i).SubItems(1)
                GoTo LoopStart
            Else
                If IsProcessAlive(dwParentPID) Then
                    dwPID = dwParentPID
                    GoTo LoopStart
                Else
                    GoTo Finale
                End If
            End If
        End If
    Next
    
    If EndProcess(dwPID) = False Then bFailed = True
    
Finale:
    If (dwPID <> dwParentPID) Then
        dwPID = dwParentPID
        GoTo LoopStart
    End If
    
    If (bFailed) Then
        Call MessageBox(Me.hWnd, "One or more processes could not be killed." & vbCrLf & "The operation did not complete.", "PsWalker", vbCritical)
    End If
End Sub

Private Sub mnuRestart_Click()
    If RestartProcess(lsProcesses.SelectedItem.SubItems(1)) = False Then
        Call MessageBox(Me.hWnd, "Failed to restart process.", "PsWalker", vbCritical)
    End If
End Sub

Private Sub mnuResume_Click()
    If ResumeProcess(lsProcesses.SelectedItem.SubItems(1)) = False Then
        Call MessageBox(Me.hWnd, "Failed to resume process.", "PsWalker", vbCritical)
    End If
End Sub

Private Sub mnuSuspend_Click()
    If SuspendProcess(lsProcesses.SelectedItem.SubItems(1)) = False Then
        Call MessageBox(Me.hWnd, "Failed to suspend process.", "PsWalker", vbCritical)
    End If
End Sub

Private Sub mnuHalfSecs_Click()
    Timer1.Interval = 500
    mnuHalfSecs.Checked = True
    mnuOneSec.Checked = False
    mnuTwoSecs.Checked = False
    mnuFiveSecs.Checked = False
    mnuPauseUpdate.Checked = False
End Sub

Private Sub mnuOneSec_Click()
    Timer1.Interval = 1000
    mnuHalfSecs.Checked = False
    mnuOneSec.Checked = True
    mnuTwoSecs.Checked = False
    mnuFiveSecs.Checked = False
    mnuPauseUpdate.Checked = False
End Sub

Private Sub mnuTwoSecs_Click()
    Timer1.Interval = 2000
    mnuHalfSecs.Checked = False
    mnuOneSec.Checked = False
    mnuTwoSecs.Checked = True
    mnuFiveSecs.Checked = False
    mnuPauseUpdate.Checked = False
End Sub

Private Sub mnuFiveSecs_Click()
    Timer1.Interval = 5000
    mnuHalfSecs.Checked = False
    mnuOneSec.Checked = False
    mnuTwoSecs.Checked = False
    mnuFiveSecs.Checked = True
    mnuPauseUpdate.Checked = False
End Sub

Private Sub mnuPauseUpdate_Click()
    Timer1.Interval = 0
    mnuHalfSecs.Checked = False
    mnuOneSec.Checked = False
    mnuTwoSecs.Checked = False
    mnuFiveSecs.Checked = False
    mnuPauseUpdate.Checked = True
End Sub

Private Sub mnuRefreshProcess_Click()
    Call InitProcessList(lsProcesses)
End Sub

Private Sub mnuKillProcess_Click()
    If EndProcess(lsProcesses.SelectedItem.SubItems(1)) = False Then
        Call MessageBox(Me.hWnd, "Failed to kill process.", "PsWalker", vbCritical)
    End If
End Sub

Function UpdateProcessList()
    Dim i As Long, j As Long, nBeenNull As Integer, ProcessCount As Long
    Dim dwPID As Long, szImagePath As String, IsAlive As Boolean, szParentName As String
    Dim b As Boolean
    
    nBeenNull = 0
    
    ProcessCount = lsProcesses.ListItems.Count
    
    For i = 1 To ProcessCount
        If (i <= ProcessCount) Then
            dwPID = lsProcesses.ListItems(i).SubItems(1)
            
            If (Not IsProcessAlive(dwPID)) Then
                lsProcesses.ListItems.Remove i
                i = i - 1
                ProcessCount = lsProcesses.ListItems.Count
            End If
        End If
    Next
    
    For i = 0& To 2147483647 Step 4
        If nBeenNull >= 32767 Then Exit For
        
        DoEvents
        
        IsAlive = IsProcessAlive(i)
        
        If (IsAlive) Then
            szImagePath = StandardizeProcessPath(GetProcessImagePath(i))
            nBeenNull = 0
            
            If PIDExistsInListView(lsProcesses, GetProcessParentPID(i), j) Then
                szParentName = lsProcesses.ListItems(j).Text
            Else
                szParentName = "<Non-existent process>"
            End If
                
            If (PIDExistsInListView(lsProcesses, i, 0) = False) Then
                b = DrawFileIconInPictureBox(szImagePath, Picture1)
                If (b) Then ImageList1.ListImages.Add , , Picture1.Image
                
                lsProcesses.ListItems.Add , , IIf(szImagePath = "", "???", GetFileNameByPath(szImagePath)), , IIf(b, ImageList1.ListImages.Count, LoadPicture(""))
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(1) = i
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(2) = IIf(szImagePath = "", "[Error opening process]", szImagePath)
                lsProcesses.ListItems(lsProcesses.ListItems.Count).SubItems(3) = szParentName & " (" & GetProcessParentPID(i) & ")"
            End If
        Else
            nBeenNull = nBeenNull + 1
        End If
    Next
End Function

Function PIDExistsInListView(objLv As ListView, ByVal dwPID As Long, Index As Long) As Boolean
    Dim i As Long
    
    For i = 1 To objLv.ListItems.Count
        If objLv.ListItems(i).SubItems(1) = dwPID Then
            Index = i
            PIDExistsInListView = True
            Exit Function
        End If
    Next
End Function

Function DrawFileIconInPictureBox(ByVal szPath As String, ByRef PictureBoxToDrawIn As PictureBox) As Boolean
    Dim hSmallIcon As Long
    
    PictureBoxToDrawIn.Picture = LoadPicture("")
    
    Call ExtractIconEx(szPath, 0, 0, hSmallIcon, 1)
    
    If (hSmallIcon) Then
        Call DrawIconEx(PictureBoxToDrawIn.hdc, 0, 0, hSmallIcon, 16, 16, 0, 0, DI_NORMAL)
    Else
        hSmallIcon = LoadIcon(0, IDI_APPLICATION)
        Call DrawIconEx(PictureBoxToDrawIn.hdc, 0, 0, hSmallIcon, 16, 16, 0, 0, DI_NORMAL)
    End If
    
    DrawFileIconInPictureBox = True
    
    Call DestroyIcon(hSmallIcon)
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call UpdateProcessList
    Timer1.Enabled = True
End Sub
