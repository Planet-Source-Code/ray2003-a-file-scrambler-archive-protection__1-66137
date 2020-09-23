VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tpIsRoot As Boolean
Public tpParent As clsDir
Public tpName As String
Public tpPath As String
Public tpSubDirs As Long
Public tpFiles As Long

Private prvt_tpSd() As clsDir
Private prvt_tpFn() As String
Private prvt_tpFt() As String

Public Function SubDir(lIndex As Long) As clsDir
    Set SubDir = prvt_tpSd(lIndex)
End Function

Public Function Filename(lIndex As Long) As String
    Filename = prvt_tpFn(lIndex)
End Function

Public Function FileTag(lIndex As Long) As String
    FileTag = prvt_tpFt(lIndex)
End Function

Public Sub AddDir(clsNewDir As clsDir)
    ReDim Preserve prvt_tpSd(0 To tpSubDirs)
    Set prvt_tpSd(tpSubDirs) = clsNewDir
    tpSubDirs = tpSubDirs + 1
End Sub

Public Sub AddFile(sFilename As String, sFiletag As String)
    ReDim Preserve prvt_tpFn(0 To tpFiles)
    ReDim Preserve prvt_tpFt(0 To tpFiles)
    prvt_tpFn(tpFiles) = sFilename
    prvt_tpFt(tpFiles) = sFiletag
    tpFiles = tpFiles + 1
End Sub

Public Sub RemoveDir(lIndex As Long)
    Dim l As Long
    If lIndex < (tpSubDirs - 1) Then
        Set prvt_tpSd(lIndex) = Nothing
        For l = lIndex To tpSubDirs - 2
            Set prvt_tpSd(l) = prvt_tpSd(l + 1)
            Set prvt_tpSd(l + 1) = Nothing
        Next
    End If
    Set prvt_tpSd(tpSubDirs - 1) = Nothing
    If tpSubDirs > 1 Then ReDim Preserve prvt_tpSd(0 To tpSubDirs - 2)
    tpSubDirs = tpSubDirs - 1
End Sub

Public Sub RemoveFile(lIndex As Long)
    Dim l As Long
    If lIndex < (tpFiles - 1) Then
        For l = lIndex To tpFiles - 2
            prvt_tpFn(l) = prvt_tpFn(l + 1)
            prvt_tpFt(l) = prvt_tpFt(l + 1)
        Next
    End If
    If tpFiles > 1 Then ReDim Preserve prvt_tpFn(0 To tpFiles - 2)
    If tpFiles > 1 Then ReDim Preserve prvt_tpFt(0 To tpFiles - 2)
    tpFiles = tpFiles - 1
End Sub
