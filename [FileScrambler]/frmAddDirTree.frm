VERSION 5.00
Begin VB.Form frmAddDirTree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adding directory tree to archive tree... Please wait..."
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblInfo 
      Caption         =   "Adding ##..."
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   12375
   End
End
Attribute VB_Name = "frmAddDirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddDirToTree(sLocalPath As String, clsEntryDir As clsDir)
    Dim l As Long, k As Long
    sLocalPath = AddSep(sLocalPath)
    lblInfo.Caption = "Adding " & sLocalPath & "..."
    DoEvents
    CreateDir clsEntryDir, DirName(sLocalPath)
    ListDirX sLocalPath
    If lFilecount > 0 Then
        For l = 0 To lFilecount - 1
            clsEntryDir.SubDir(clsEntryDir.tpSubDirs - 1).AddFile sFiles(l), sLocalPath & sFiles(l)
        Next
    End If
    If lDircount > 0 Then
        k = lDircount
        For l = 0 To k - 1
            AddDirToTree sLocalPath & sDirs(l), clsEntryDir.SubDir(clsEntryDir.tpSubDirs - 1)
            ListDirX sLocalPath
        Next
    End If
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    AddDirToTree frmMain.Dir1.Path, xCurr
    Unload Me
End Sub
