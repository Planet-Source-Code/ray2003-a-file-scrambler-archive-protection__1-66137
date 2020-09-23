Attribute VB_Name = "modFileBrowse"
Option Explicit
Public lDrivecount As Long
Public lDircount As Long
Public lFilecount As Long
Public sDrives() As String
Public sDriveNames() As String
Public sDirs() As String
Public sFiles() As String

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub ListDirX(sDir As String)
    On Error Resume Next
    Dim sSub As String
    lDircount = 0
    lFilecount = 0
    ReDim sDirs(lDircount)
    ReDim sFiles(lFilecount)
    sSub = Dir$(sDir, vbArchive + vbDirectory + vbHidden + vbNormal + vbSystem)
    While sSub <> ""
        If (Not (sSub = ".")) And (Not (sSub = "..")) Then
            If (GetAttr(sDir & sSub) And vbDirectory) = vbDirectory Then
                lDircount = lDircount + 1
                ReDim Preserve sDirs(lDircount - 1)
                sDirs(lDircount - 1) = sSub
            Else
                lFilecount = lFilecount + 1
                ReDim Preserve sFiles(lFilecount - 1)
                sFiles(lFilecount - 1) = sSub
            End If
        End If
        sSub = Dir$
    Wend
End Sub

Public Sub ListDrives()
    On Error Resume Next
    Dim sTemp As String
    Dim sTemp2 As String
    Dim iNullSpot As Integer
    Dim lDrive As Long
    lDrivecount = 0
    ReDim sDrives(lDrivecount)
    ReDim sDriveNames(lDrivecount)
    sTemp = String$(2048, 0)
    Call GetLogicalDriveStrings(2047, sTemp)
    Do
        iNullSpot = InStr(sTemp, Chr$(0))
        If iNullSpot > 1 Then
            sTemp2 = UCase$(Left$(sTemp, iNullSpot - 2))
            lDrive = GetDriveType(sTemp2)
            ReDim Preserve sDrives(lDrivecount)
            ReDim Preserve sDriveNames(lDrivecount)
            sDrives(lDrivecount) = sTemp2 & "\"
            Select Case lDrive
                Case 1, 3
                    sDriveNames(lDrivecount) = "[" & IIf(GetDriveName(sTemp2) <> "", GetDriveName(sTemp2), "NONAME") & ",PARTITION]"
                Case 5
                    sDriveNames(lDrivecount) = "[" & IIf(GetDriveName(sTemp2) <> "", GetDriveName(sTemp2), "NONAME") & ",CDROM]"
                Case 4
                    sDriveNames(lDrivecount) = "[" & IIf(GetDriveName(sTemp2) <> "", GetDriveName(sTemp2), "NONAME") & ",REMOTE]"
                Case 2
                    sDriveNames(lDrivecount) = "[NONAME,REMOVABLE]"
                Case Else
                    sDriveNames(lDrivecount) = "[NONAME,UNKNOWN]"
            End Select
            lDrivecount = lDrivecount + 1
            sTemp = Mid$(sTemp, iNullSpot + 1)
        End If
    Loop Until iNullSpot <= 1
End Sub

Public Function PreviousDir(sCurDir As String) As String
    Dim i As Integer
    Dim s As String
    s = IIf(Right(sCurDir, 1) = "\", Left(sCurDir, Len(sCurDir) - 1), sCurDir)
    i = InStr(StrReverse(s), "\")
    If i > 0 Then PreviousDir = Left(sCurDir, Len(sCurDir) - i)
End Function

Public Function AddSep(s As String) As String
    AddSep = s & IIf(Right(s, 1) = "\", "", "\")
End Function

Public Function DirName(sCurDir As String) As String
    Dim i As Integer
    Dim s As String
    s = IIf(Right(sCurDir, 1) = "\", Left(sCurDir, Len(sCurDir) - 1), sCurDir)
    i = InStr(StrReverse(s), "\")
    If i > 0 Then DirName = Right(s, i - 1)
End Function

Private Function GetDriveName(ByVal sDrive As String) As String
    Dim sVolBuf As String, sSysName As String
    Dim lSerialNum As Long, lSysFlags As Long, lComponentLength As Long
    Dim lRet As Long
    sVolBuf = String$(256, 0)
    sSysName = String$(256, 0)
    lRet = GetVolumeInformation(sDrive, sVolBuf, 255, lSerialNum, lComponentLength, lSysFlags, sSysName, 255)
    If lRet > 0 Then GetDriveName = StripTerminator(sVolBuf)
End Function

Private Function StripTerminator(ByVal sString As String) As String
    Dim iZeroPos As Long
    iZeroPos = InStr(sString, Chr$(0))
    If iZeroPos > 0 Then
        StripTerminator = Left$(sString, iZeroPos - 1)
    Else
        StripTerminator = sString
    End If
End Function
