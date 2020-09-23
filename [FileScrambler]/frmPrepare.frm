VERSION 5.00
Begin VB.Form frmPrepare 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preparing to scramble..."
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtLog 
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
      Height          =   2295
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label lblCurrent 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   7335
   End
End
Attribute VB_Name = "frmPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExecSub()
    Dim l As Long, k As Long, tmp_a As Currency, tmp_b As Currency, tmp_c As Currency
    txtLog.Text = "Key set to " & Len(frmMain.txtKey.Text) & " bits, initializing scrambler...."
    DoEvents
    Sleep 100
    sWritePath = frmMain.lblArchSaveLoc.Caption
    ResetScrambler
    txtLog.Text = txtLog.Text & vbCrLf & "Analyzing archive contents, creating directory tree and file tree...."
    DoEvents
    Sleep 100
    lArchDirCount = 0
    lArchFileCount = 0
    LoadArchDir xRoot
    txtLog.Text = txtLog.Text & vbCrLf & "Loaded " & lArchDirCount & " directories containing " & lArchFileCount & " files"
    Sleep 100
    txtLog.Text = txtLog.Text & vbCrLf & "Retrieving file lengths...."
    DoEvents
    Sleep 100
    ReDim Preserve lArchFileLens(0 To lArchFileCount - 1)
    For l = 0 To lArchFileCount - 1
        k = FileLen(sArchFilePaths(l))
        lArchFileLens(l) = k
        lblCurrent.Caption = sArchFilePaths(l) & "  (" & k & " bytes)"
        DoEvents
    Next
    txtLog.Text = txtLog.Text & vbCrLf & "Estimating total archive size...."
    Sleep 100
    dArchSize = Len(CStr(lArchDirCount)) + 1 + Len(CStr(lArchFileCount)) + 1
    dArchFilesSize = 0
    For l = 0 To lArchDirCount - 1
        lblCurrent.Caption = dArchSize & " bytes"
        DoEvents
        dArchSize = dArchSize + Len(sArchDirs(l)) + 1
    Next
    For l = 0 To lArchFileCount - 1
        lblCurrent.Caption = dArchSize & " bytes"
        DoEvents
        dArchSize = dArchSize + Len(sArchFileDirs(l)) + 1
        dArchSize = dArchSize + Len(sArchFileNames(l)) + 1
        dArchSize = dArchSize + Len(CStr(lArchFileLens(l))) + 1
        dArchSize = dArchSize + lArchFileLens(l)
        dArchFilesSize = dArchFilesSize + lArchFileLens(l)
    Next
    dArchSize = dArchSize + Len(CStr(dArchFilesSize)) + 1
    txtLog.Text = txtLog.Text & vbCrLf & "Archive size estimated to " & dArchSize & " bytes"
    txtLog.Text = txtLog.Text & vbCrLf & "Retrieving free disk space for " & frmMain.lblArchSaveLoc.Caption & "...."
    DoEvents
    Sleep 100
    Call SHGetDiskFreeSpace(Left(frmMain.lblArchSaveLoc.Caption, 3), tmp_a, tmp_b, tmp_c)
    txtLog.Text = txtLog.Text & vbCrLf & "Free disk space for archive location: " & CStr(CDbl(tmp_a * 10000)) & " bytes"
    If (tmp_a * 10000) < dArchSize Then
        MsgBox "You do not have enough free space on the specified location. Please select another location or abort.", vbExclamation
        Unload Me
        Exit Sub
    End If
    txtLog.Text = txtLog.Text & vbCrLf & vbCrLf & "Relaxing CPU, please wait a few seconds...."
    DoEvents
    Sleep 8000
    OperMode = 1
    Unload Me
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    ExecSub
End Sub

Private Sub LoadArchDir(clsDir As clsDir)
    Dim sDir As String, l As Long
    sDir = Mid(clsDir.tpPath & clsDir.tpName & "\", 3)
    lblCurrent.Caption = sDir
    DoEvents
    If Not sDir = "\" Then
        ReDim Preserve sArchDirs(lArchDirCount)
        sArchDirs(lArchDirCount) = sDir
        lArchDirCount = lArchDirCount + 1
    End If
    For l = 0 To clsDir.tpFiles - 1
        ReDim Preserve sArchFileNames(0 To lArchFileCount)
        ReDim Preserve sArchFilePaths(0 To lArchFileCount)
        ReDim Preserve sArchFileDirs(0 To lArchFileCount)
        sArchFileNames(lArchFileCount) = clsDir.Filename(l)
        sArchFilePaths(lArchFileCount) = clsDir.FileTag(l)
        sArchFileDirs(lArchFileCount) = sDir
        lArchFileCount = lArchFileCount + 1
    Next
    For l = 0 To clsDir.tpSubDirs - 1
        LoadArchDir clsDir.SubDir(l)
    Next
End Sub
