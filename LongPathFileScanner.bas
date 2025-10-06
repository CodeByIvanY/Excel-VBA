'Macro - LongPathFileScanner
'Created by: Ivan Yang
'Created Date: 10/06/2025
'Purpose: A VBA utility that scans and lists .xlsx files in directories with paths exceeding the 256-character limit, using Windows API for extended path support.

Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(0 To 259) As Integer
    cAlternate(0 To 13) As Integer
End Type

Private Function FileNameFromBuffer(buffer() As Integer) As String
    Dim i As Long
    Dim result As String
    i = 0
    Do While buffer(i) <> 0 And i <= UBound(buffer)
        result = result & ChrW(buffer(i))
        i = i + 1
    Loop
    FileNameFromBuffer = result
End Function

Public Sub ShowFiles_356(dict_data As Object, ByVal SourceFolderPath As String, Optional ByVal FilePattern As String)
    Dim folderPath As String

    Dim findData As WIN32_FIND_DATA
    Dim hFind As LongPtr
    Dim fileName As String

    folderPath = SourceFolderPath & "\*.xlsx" ' Enable long path support

    hFind = FindFirstFileW(StrPtr(folderPath), findData)

    If hFind <> -1 Then
        Do
            fileName = FileNameFromBuffer(findData.cFileName)
            If Left(fileName, 1) <> "~" Then
                dict_data.Add SourceFolderPath & "\" & fileName, fileName
            End If
        Loop While FindNextFileW(hFind, findData)
        FindClose hFind
    End If
End Sub

