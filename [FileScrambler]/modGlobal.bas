Attribute VB_Name = "modGlobal"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long

Public OperMode As Long ' 1=scramble, 2=descramble

Public lArchDirCount As Long
Public sArchDirs() As String

Public lArchFileCount As Long
Public sArchFileNames() As String ' File name
Public sArchFilePaths() As String ' Original path
Public sArchFileDirs() As String  ' Archive location (always ends with \)
Public lArchFileLens() As Long    ' File size

Public dArchSize As Double        ' Total archive size
Public dArchFilesSize As Double   ' Total size of all files in archive

Public sDescrArchLoc As String
Public sDescrExtrLoc As String
