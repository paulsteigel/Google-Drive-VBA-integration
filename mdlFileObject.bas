Option Explicit
''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

Function GetMimeType(FileName As String) As String
    Dim extStr As String
    GetMimeType = nz(DLookup("MimeType", "tblMimeType", "Extension='" + "." + GetFileExtension(FileName) + "'"), "*/*")
End Function

Function GetFileName(FileName As String) As String
    GetFileName = Mid(FileName, InStrRev(FileName, "\") + 1)
End Function

Function GetFileExtension(FileName As String) As String
    GetFileExtension = Mid(FileName, InStrRev(FileName, ".") + 1)
End Function

Function FileOrDirExists(Pathname As String, Optional FileObject As Boolean = False) As Boolean
'No need to set a reference if you use Late binding
    Dim fso As Object
    Dim filePath As String, lRet As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Pathname = "" Then Exit Function
    If FileObject Then
        FileOrDirExists = fso.FileExists(Pathname)
    Else
        FileOrDirExists = fso.FolderExists(Pathname)
    End If
    Set fso = Nothing
End Function

Function GetFilePath(FileFullName As String) As String
     Dim tArr As Variant
     If FileFullName <> "" Then
        tArr = Split(FileFullName, "\")
        GetFilePath = VBA.Replace(FileFullName, "\" & tArr(UBound(tArr)), "")
    End If
End Function

Public Function GetFileSize(sFilePath As String) As Long
    'On Error Resume Next
    Dim fs As Object, oReadfile As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oReadfile = fs.GetFile(sFilePath)
    GetFileSize = oReadfile.size
    Set oReadfile = Nothing
End Function

Function ReadBinaryFile(FileName)
    Const adTypeBinary = 1
    
    'Create Stream object
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To get binary data.
    BinaryStream.Type = adTypeBinary
    
    'Open the stream
    BinaryStream.Open
    
    'Load the file data from disk To stream object
    BinaryStream.LoadFromFile FileName
    
    'Open the stream And get binary data from the object
    ReadBinaryFile = BinaryStream.Read
End Function

Function GetFileBinary(path As String) As String
    Dim A, fso As Object, file As Object, i As Long, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(path)
    If IsNull(file) Then
        MsgBox ("File not found: " & path)
        Exit Function
    End If
    Set ts = file.OpenAsTextStream()
    A = makeArray(file.size)
    i = 0
    ' Do not replace the following block by readBinary = by ts.readAll(), it would result in broken output, because that method is not intended for binary data
    While Not ts.atEndOfStream
        A(i) = ts.Read(1)
    i = i + 1
    Wend
    ts.Close
    GetFileBinary = Join(A, "")
    Set ts = Nothing
    Set fso = Nothing
End Function

Function WriteFileBinary(FileName As String, ByteArray As Variant)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    
    'Create Stream object
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To save binary data.
    BinaryStream.Type = adTypeBinary
    
    'Open the stream And write binary data To the object
    BinaryStream.Open
    BinaryStream.Write ByteArray
    
    'Save binary data To disk
    BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Private Function makeArray(n) ' Small utility function
    Dim S
    S = Space(n)
    makeArray = Split(S, " ")
End Function

Public Function GetFileText(sFilePath As String) As String
   On Error Resume Next
   Dim fs As Object, oReadfile As Object
    Set fs = CreateObject("scripting.filesystemobject")
    Set oReadfile = fs.OpenTextFile(sFilePath, 1, False, True)
    GetFileText = oReadfile.ReadAll
    Set oReadfile = Nothing
End Function

Function BrowseFileOrDir(Optional BrowseFile As Boolean = True, Optional retTxtObject As Control) As String
    ' Start browsing folder...
    Dim iCnt As Long, FileObj As Object, MsgStr As String
BackOpt:
    If BrowseFile Then
        Set FileObj = Application.FileDialog(3)
        MsgStr = "Chon tap tin de tai len"
    Else
        Set FileObj = Application.FileDialog(4)
        MsgStr = "Chon thu muc"
    End If
    With FileObj  '4=msoFileDialogFolderPicker
        .title = MsgStr
        If .Show Then
            If .SelectedItems(1) <> "" Then
                BrowseFileOrDir = .SelectedItems(1)
                If Not retTxtObject Is Nothing Then retTxtObject = BrowseFileOrDir
            End If
        End If
    End With
ExitSub:
End Function
