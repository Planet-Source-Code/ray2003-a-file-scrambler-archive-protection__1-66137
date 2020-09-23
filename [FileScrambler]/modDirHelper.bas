Attribute VB_Name = "modDirHelper"
Option Explicit
Public xRoot As clsDir
Public xCurr As clsDir

Public Sub DirInit()
    Set xRoot = New clsDir
    xRoot.tpIsRoot = True
    Set xRoot.tpParent = Nothing
    xRoot.tpPath = "\"
    xRoot.tpName = ":"
    Set xCurr = xRoot
    ListDir xRoot, frmMain.lstArchDirs, frmMain.lstArchFiles, frmMain.lblContentLocation
End Sub

Public Sub CreateDir(clsParent As clsDir, sName As String)
    Dim x As New clsDir
    Set x.tpParent = clsParent
    x.tpPath = x.tpParent.tpPath & x.tpParent.tpName & "\"
    x.tpName = sName
    x.tpParent.AddDir x
End Sub

Public Sub ListDir(clsDir As clsDir, objDestDir As ListBox, objDestFil As ListBox, lblCurrent As Label)
    Dim l As Long
    objDestDir.Clear
    objDestFil.Clear
    lblCurrent.Caption = clsDir.tpPath & clsDir.tpName
    If Not clsDir.tpIsRoot Then
        objDestDir.AddItem ".."
        For l = 1 To clsDir.tpSubDirs
            objDestDir.AddItem clsDir.SubDir(l - 1).tpName
        Next
    Else
        For l = 0 To clsDir.tpSubDirs - 1
            objDestDir.AddItem clsDir.SubDir(l).tpName
        Next
    End If
    For l = 0 To clsDir.tpFiles - 1
        objDestFil.AddItem clsDir.Filename(l)
    Next
End Sub

Public Function FindDirByName(clsDir As clsDir, sName As String) As Long
    Dim l As Long
    FindDirByName = -1
    'If clsDir.tpSubDirs > 0 Then
        For l = 0 To clsDir.tpSubDirs - 1
            If clsDir.SubDir(l).tpName = sName Then
                FindDirByName = l
                Exit Function
            End If
        Next
    'End If
End Function

Public Function FindFileByName(clsDir As clsDir, sName As String) As Long
    Dim l As Long
    FindFileByName = -1
    'If clsDir.tpFiles > 0 Then
        For l = 0 To clsDir.tpFiles - 1
            If clsDir.Filename(l) = sName Then
                FindFileByName = l
                Exit Function
            End If
        Next
    'End If
End Function

