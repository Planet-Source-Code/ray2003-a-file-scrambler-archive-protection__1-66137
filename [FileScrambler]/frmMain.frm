VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Scrambler"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freKey 
      Caption         =   "Scramble key  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   2415
      Begin MSComDlg.CommonDialog cdl1 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdLoadKeyBit 
         Caption         =   "Load key from bit file"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdCreateKeyRandom 
         Caption         =   "Create new key (random)"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame freDescramble 
      Caption         =   "Archive descrambler  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   1
      Top             =   7200
      Width           =   11295
      Begin VB.CommandButton cmdDescramble 
         Caption         =   "<< DESCRAMBLE ARCHIVE >>"
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdBrowseDescrambleDestination 
         Caption         =   "..."
         Height          =   255
         Left            =   7440
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdBrowseDescrambleFile 
         Caption         =   "..."
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDestination 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Destination:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblScrambledFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Scrambled file:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame freScramble 
      Caption         =   "Scrambler  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      Begin VB.CommandButton cmdBrowseScrambleArchive 
         Caption         =   "..."
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   6000
         Width           =   255
      End
      Begin VB.CommandButton cmdScramble 
         Caption         =   "<< SCRAMBLE ARCHIVE >>"
         Height          =   375
         Left            =   8400
         TabIndex        =   20
         Top             =   6000
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   3915
         Left            =   120
         TabIndex        =   18
         Top             =   1170
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Height          =   4380
         Hidden          =   -1  'True
         Left            =   2880
         System          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2655
      End
      Begin VB.ListBox lstArchFiles 
         Height          =   4155
         Left            =   8520
         TabIndex        =   14
         Top             =   960
         Width           =   2655
      End
      Begin VB.ListBox lstArchDirs 
         Height          =   4155
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdCreateDirInArch 
         Caption         =   "Create directory in archive"
         Height          =   375
         Left            =   5760
         TabIndex        =   12
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Scrambled archive save location:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Label lblArchSaveLoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   6000
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Open directory or select file and press <SPACE> to add dir or file to archive"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   5415
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   11160
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label lblContentLocation 
         Caption         =   "\"
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   720
         Width           =   5415
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   11160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label4 
         Caption         =   "New archive contents"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "File browser"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5295
      End
      Begin VB.Line Line1 
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   5760
      End
      Begin VB.Label Label5 
         Caption         =   "Press <DEL> to remove dir or file from archive contents"
         Height          =   495
         Left            =   8520
         TabIndex        =   15
         Top             =   5280
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseDescrambleDestination_Click()
    frmDescrDest.Show vbModal
End Sub

Private Sub cmdBrowseDescrambleFile_Click()
    frmLocateArchive.Show vbModal
End Sub

Private Sub cmdBrowseScrambleArchive_Click()
    frmSaveArchive.Show vbModal
End Sub

Private Sub cmdCreateDirInArch_Click()
    Dim s As String
    s = InputBox("Name of new directory:")
    If s <> "" Then
        If InStr(s, "\") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "/") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, ":") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "*") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "?") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, """") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "<") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, ">") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If InStr(s, "|") Then MsgBox "Invalid directory name", vbCritical: Exit Sub
        If FindDirByName(xCurr, s) > -1 Then
            MsgBox "This directory name already exists in the current archive tree directory", vbExclamation
            Exit Sub
        End If
        CreateDir xCurr, s
        ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
    End If
End Sub

Private Sub cmdCreateKeyRandom_Click()
    Dim s As String, t As String, l As Long, k As Long
    s = InputBox("Number of bits: ")
    If s <> "" Then
        If IsNumeric(s) Then
            l = CLng(s)
            If l > 3 And l < 8193 Then
                s = ""
                Randomize
                For k = 1 To l
                    s = s & Round(Rnd)
                Next
                cdl1.Filename = ""
                cdl1.Filter = "*.fs-bitkey|*.fs-bitkey"
                cdl1.ShowSave
                If cdl1.Filename <> "" Then
                    t = cdl1.Filename
                    If Right(t, 10) <> ".fs-bitkey" Then t = t & ".fs-bitkey"
                    Open t For Binary As #1
                        Put #1, , s
                    Close #1
                    If LoadScrambleKey(s) Then
                        txtKey.Text = s
                        MsgBox "Key succefully generated, loaded and saved", vbInformation
                    End If
                End If
            Else
                MsgBox "Invalid length"
            End If
        Else
            MsgBox "Invalid length"
        End If
    End If
End Sub

Private Sub cmdDescramble_Click()
    If txtKey.Text = "" Then
        MsgBox "Scramble key not loaded, please set before proceding", vbExclamation
        Exit Sub
    End If
    If lblScrambledFile.Caption = "" Then
        MsgBox "Scrambled archive not located, please set before proceding", vbExclamation
        Exit Sub
    End If
    If lblDestination.Caption = "" Then
        MsgBox "Destination directory not set, please set before proceding", vbExclamation
        Exit Sub
    End If
    If MsgBox("WARNING: This program will hang or crash if the loaded key is not the same as the original scramble key for this archive. Proceed at your own risk!" & vbCrLf & vbCrLf & "Do you want to proceed?", vbYesNo Or vbExclamation) = vbYes Then
        sDescrArchLoc = AddSep(PreviousDir(lblScrambledFile.Caption))
        sDescrExtrLoc = Left(lblDestination.Caption, Len(lblDestination) - 1)
        Unload frmDescrDest
        Unload frmLocateArchive
        Unload Me
        OperMode = 2
        frmProgress.Show
    End If
End Sub

Private Sub cmdLoadKeyBit_Click()
    Dim s As String, l As Long, k As Long
    cdl1.Filename = ""
    cdl1.Filter = "*.fs-bitkey|*.fs-bitkey"
    cdl1.ShowOpen
    If cdl1.Filename <> "" Then
        l = FileLen(cdl1.Filename)
        s = Space(l)
        Open cdl1.Filename For Binary As #1
            Get #1, , s
        Close #1
        For k = 1 To l
            If Not (Mid(s, k, 1) = "0" Or Mid(s, k, 1) = "1") Then
                MsgBox "Invalid key format!", vbCritical
                Exit Sub
            End If
        Next
        If LoadScrambleKey(s) Then
            txtKey.Text = s
            MsgBox "Key succesfully loaded", vbInformation
        End If
    End If
End Sub

Private Sub cmdScramble_Click()
    Dim tmp_a As Currency, tmp_b As Currency, tmp_c As Currency
    If xRoot.tpSubDirs = 0 And xRoot.tpFiles = 0 Then
        MsgBox "Scramble archive is empty, cannot proceed", vbExclamation
        Exit Sub
    End If
    If lblArchSaveLoc.Caption = "" Then
        MsgBox "Scramble archive save location not loaded, please set before proceding", vbExclamation
        Exit Sub
    End If
    If txtKey.Text = "" Then
        MsgBox "Scramble key not loaded, please set before proceding", vbExclamation
        Exit Sub
    End If
    frmPrepare.Show vbModal
    If OperMode = 1 Then
        Unload frmSaveArchive
        Unload Me
        frmProgress.Show
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Len(Dir1.Path) < 4 Then
            MsgBox "Cannot add a root. Please select only sub-directories.", vbExclamation
            Exit Sub
        End If
        If FindDirByName(xCurr, DirName(Dir1.Path)) > -1 Then
            MsgBox "This directory name already exists in the current archive tree directory", vbExclamation
            Exit Sub
        End If
        frmAddDirTree.Show vbModal
        ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
    End If
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If FindFileByName(xCurr, File1.Filename) > -1 Then
            MsgBox "This filename already exists in the current archive tree directory", vbExclamation
            Exit Sub
        End If
        xCurr.AddFile File1.Filename, File1.Path & IIf(Right(File1.Path, 1) = "\", "", "\") & File1.Filename
        ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
    End If
End Sub

Private Sub Form_Load()
    DirInit
End Sub

Private Sub lstArchDirs_DblClick()
    Dim s As String
    If lstArchDirs.ListIndex > -1 Then
        s = lstArchDirs.List(lstArchDirs.ListIndex)
        If s = ".." Then
            Set xCurr = xCurr.tpParent
        Else
            Set xCurr = xCurr.SubDir(FindDirByName(xCurr, s))
        End If
        ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
    End If
End Sub

Private Sub lstArchDirs_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim s As String
    If KeyCode = vbKeyDelete Then
        If lstArchDirs.ListIndex > -1 Then
            s = lstArchDirs.List(lstArchDirs.ListIndex)
            If Not s = ".." Then xCurr.RemoveDir FindDirByName(xCurr, s)
            ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
        End If
    End If
End Sub

Private Sub lstArchFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lstArchFiles.ListIndex > -1 Then
            xCurr.RemoveFile FindFileByName(xCurr, lstArchFiles.List(lstArchFiles.ListIndex))
            ListDir xCurr, lstArchDirs, lstArchFiles, lblContentLocation
        End If
    End If
End Sub
