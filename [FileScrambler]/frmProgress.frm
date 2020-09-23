VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scrambling progress"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame freProgress 
      Caption         =   "Progress  "
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin MSComctlLib.ProgressBar pbTotal 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdAbort 
         Caption         =   "<< ABORT OPERATION AND EXIT FS >>"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CheckBox chkShowDetailedProgress 
         Caption         =   "Do not show detailed progress (improves performance)"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1440
         Width           =   4935
      End
      Begin MSComctlLib.ProgressBar pbFile 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8400
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblInfoM 
         Caption         =   "Total files: #####, Current file: #####, Current file name: #####################################"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   8055
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "File progress:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Overall progress:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowDetailedProgress_Click()
    If chkShowDetailedProgress.Value = 1 Then
        lblInfo.Visible = False
        lblInfoM.Visible = False
        pbFile.Visible = False
    Else
        lblInfo.Visible = True
        lblInfoM.Visible = True
        pbFile.Visible = True
    End If
End Sub

Private Sub cmdAbort_Click()
    If MsgBox("Are you sure you want to abort?", vbExclamation Or vbYesNo) = vbYes Then
        TerminateScrambler
        End
    End If
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    If OperMode = 1 Then
        Scrambler
    ElseIf OperMode = 2 Then
        Descambler
    Else
        MsgBox "Unexpected error occured! (stupid debugger!! :)", vbCritical
        End
    End If
End Sub

Private Sub Scrambler()
    Dim l As Long, k As Long, m As Long, n As Long
    Dim bytBuff As Byte
    Dim lFileNumX As Long
    pbTotal.Max = dArchSize
    ScrambleStringAndWrite CStr(lArchDirCount) & ">" & CStr(lArchFileCount) & ">" & CStr(dArchFilesSize) & ">"
    For l = 0 To lArchDirCount - 1
        ScrambleStringAndWrite sArchDirs(l) & ">"
        pbTotal.Value = dWrittenTotal
        DoEvents
    Next
    For l = 0 To lArchFileCount - 1
        m = lArchFileLens(l)
        ScrambleStringAndWrite sArchFileDirs(l) & ">" & sArchFileNames(l) & ">" & CStr(m) & ">"
        lFileNumX = FreeFile
        pbTotal.Value = dWrittenTotal
        If chkShowDetailedProgress.Value = 0 Then
            pbFile.Max = m
            lblInfoM.Caption = "Total files: " & lArchFileCount & ", Current file: " & l + 1 & ", Current file name: " & sArchFileNames(l)
        End If
        DoEvents
        Open sArchFilePaths(l) For Binary As lFileNumX
            For k = 1 To m
                Get #lFileNumX, k, bytBuff
                ScrambleAndWrite bytBuff
                If chkShowDetailedProgress.Value = 0 Then
                    pbTotal.Value = dWrittenTotal
                    If Right(CStr(k), 2) = "00" Then
                        If pbFile.Max = m Then
                            pbFile.Value = k
                            DoEvents
                        End If
                    End If
                End If
            Next
        Close #lFileNumX
    Next
    ScrambleAndWrite 255
    TerminateScrambler
    Sleep 500
    MsgBox "Archive scrambling complete! Process done. FileScrambler will now close.", vbInformation
    End
End Sub

Private Sub Descambler()
    Dim l As Double, m As Long, n As Long, o As Long, p As Long
    Dim lDFileNum As Long, sFile As String
    n = CLng(ReadStr)
    m = CLng(ReadStr) ' file count
    l = CDbl(ReadStr) ' total files size
    pbTotal.Max = 1.01 * l
    For o = 1 To n
        MkDir sDescrExtrLoc & ReadStr
    Next
    For n = 1 To m
        sFile = sDescrExtrLoc & ReadStr
        sFile = sFile & ReadStr
        o = CLng(ReadStr)
        If dReadTotal < pbTotal.Max Then pbTotal.Value = dReadTotal
        If chkShowDetailedProgress.Value = 0 Then
            pbFile.Max = o
            lblInfoM.Caption = "Total files: " & m & ", Current file: " & n & ", Current file name: " & sFile
        End If
        DoEvents
        lDFileNum = FreeFile
        Open sFile For Binary As #lDFileNum
            For p = 1 To o
                Put #lDFileNum, p, ReadAndDescramble
                If chkShowDetailedProgress.Value = 0 Then
                    If dReadTotal < pbTotal.Max Then pbTotal.Value = dReadTotal
                    If Right(CStr(p), 2) = "00" Then
                        If pbFile.Max = o Then
                            pbFile.Value = p
                            DoEvents
                        End If
                    End If
                End If
            Next
        Close #lDFileNum
    Next
    MsgBox "Archive descrambling complete! Process done. FileScrambler will now close.", vbInformation
    End
End Sub

Private Function ReadStr() As String
    Dim sBuff As String, bytByte As Byte
    Do
        bytByte = ReadAndDescramble
        If Not Chr(bytByte) = ">" Then sBuff = sBuff & Chr(bytByte)
    Loop While Not Chr(bytByte) = ">"
    ReadStr = sBuff
End Function
