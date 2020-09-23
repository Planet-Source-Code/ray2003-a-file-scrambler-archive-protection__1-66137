Attribute VB_Name = "modScramble"
Option Explicit

Private ScrambleBits(0 To 7) As Byte

Private lScrambleKeyLen As Long
Private bScrambleKey() As Boolean

Private lScrambleBit As Long
Private lScrambleKeyPos As Long

Private bFileOpen As Boolean
Private lFileNum As Long

Public sWritePath As String
Private lWriteNum As Long
Private sWriteFile As String
Private lWrittenToFile As Long
Public dWrittenTotal As Double

Private lReadNum As Long
Private sReadFile As String
Private lReadFromFile As Long
Public dReadTotal As Double

Public Function LoadScrambleKey(sBitStream As String) As Boolean
    Dim l As Long
    LoadScrambleKey = False
    If Len(sBitStream) < 4 Then Exit Function
    lScrambleKeyLen = Len(sBitStream) + 1
    ReDim bScrambleKey(0 To lScrambleKeyLen - 1)
    bScrambleKey(0) = True
    For l = 1 To lScrambleKeyLen - 1
        bScrambleKey(l) = IIf(Mid(sBitStream, l, 1) = "0", False, True)
    Next l
    LoadScrambleKey = True
End Function

Public Sub ResetScrambler()
    lScrambleBit = 0
    lScrambleKeyPos = 0
    ScrambleBits(0) = 2 ^ 0
    ScrambleBits(1) = 2 ^ 1
    ScrambleBits(2) = 2 ^ 2
    ScrambleBits(3) = 2 ^ 3
    ScrambleBits(4) = 2 ^ 4
    ScrambleBits(5) = 2 ^ 5
    ScrambleBits(6) = 2 ^ 6
    ScrambleBits(7) = 2 ^ 7
End Sub

Public Sub TerminateScrambler()
    bFileOpen = False
    Close #lFileNum
End Sub

Public Sub ScrambleStringAndWrite(sString As String)
    Dim l As Long
    For l = 1 To Len(sString)
        ScrambleAndWrite (Asc(Mid(sString, l, 1)))
    Next
End Sub

Public Sub ScrambleAndWrite(bytByte As Byte)
    If Not bFileOpen Then
        lWriteNum = 0
        lWrittenToFile = 0
        dWrittenTotal = 0
        sWriteFile = sWritePath & MakeLen(CStr(lWriteNum), 4, "0", True) & ".fs-data"
        lFileNum = FreeFile
        Open sWriteFile For Binary As #lFileNum
        bFileOpen = True
    End If
    lWrittenToFile = lWrittenToFile + 1
    dWrittenTotal = dWrittenTotal + 1
    Put #lFileNum, lWrittenToFile, ScrambleByte(bytByte)
    If lWrittenToFile = 104857599 Then ' (100 MB - 1 Byte)
        Close #lFileNum
        lWriteNum = lWriteNum + 1
        lWrittenToFile = 0
        sWriteFile = sWritePath & MakeLen(CStr(lWriteNum), 4, "0", True) & ".fs-data"
        lFileNum = FreeFile
        Open sWriteFile For Binary As #lFileNum
    End If
End Sub

Public Function ReadAndDescramble() As Byte
    Dim bytByte As Byte
    If Not bFileOpen Then
        lReadNum = 0
        lReadFromFile = 0
        dReadTotal = 0
        sReadFile = sDescrArchLoc & MakeLen(CStr(lReadNum), 4, "0", True) & ".fs-data"
        If (Dir$(sReadFile) = "") Then
            MsgBox "Fatal error!" & vbCrLf & vbCrLf & "Could not find file '" & sReadFile & "'. Program will now abort.", vbCritical
            End
        End If
        ResetScrambler
        lFileNum = FreeFile
        Open sReadFile For Binary As #lFileNum
        bFileOpen = True
    End If
    dReadTotal = dReadTotal + 1
    lReadFromFile = lReadFromFile + 1
    Get #lFileNum, lReadFromFile, bytByte
    ReadAndDescramble = ScrambleByte(bytByte)
    If lReadFromFile = 104857599 Then ' (100 MB - 1 Byte)
        Close #lFileNum
        lReadNum = lReadNum + 1
        lReadFromFile = 0
        sReadFile = sDescrArchLoc & MakeLen(CStr(lReadNum), 4, "0", True) & ".fs-data"
        If (Dir$(sReadFile) = "") Then
            MsgBox "Fatal error!" & vbCrLf & vbCrLf & "Could not find file '" & sReadFile & "'. Program will now abort.", vbCritical
            End
        End If
        lFileNum = FreeFile
        Open sReadFile For Binary As #lFileNum
    End If
End Function

Public Function ScrambleByte(bytByte As Byte) As Byte
    If Not bScrambleKey(lScrambleKeyPos) Then
        ScrambleByte = bytByte Xor 255
        Call RotateScrambleBit(1)
        Call RotateScrambleKeyPos
    Else
        ScrambleByte = bytByte Xor ScrambleBits(lScrambleBit)
        Call RotateScrambleBit(2)
        Call RotateScrambleKeyPos
    End If
End Function

Private Sub RotateScrambleBit(lAmount As Long)
    lScrambleBit = lScrambleBit + lAmount
    If Not lScrambleBit < 8 Then lScrambleBit = lScrambleBit Mod 8
End Sub

Private Sub RotateScrambleKeyPos()
    lScrambleKeyPos = (lScrambleKeyPos + 1)
    If Not (lScrambleKeyPos < lScrambleKeyLen) Then lScrambleKeyPos = lScrambleKeyPos Mod lScrambleKeyLen
End Sub

Private Function MakeLen(sIn As String, lLen As Long, sChar As String, bLeftside As Boolean) As String
    Dim l As Long
    If Len(sIn) < lLen Then
        If bLeftside Then
            MakeLen = String(lLen - Len(sIn), sChar) & sIn
        Else
            MakeLen = sIn & String(lLen - Len(sIn), sChar)
        End If
    ElseIf Len(sIn) > lLen Then
        If bLeftside Then
            MakeLen = Right(sIn, lLen)
        Else
            MakeLen = Left(sIn, lLen)
        End If
    Else
        MakeLen = sIn
    End If
End Function
