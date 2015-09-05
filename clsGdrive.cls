Option Compare Database
Option Explicit

Public Event StatusChanged(ByVal Status As String)
Public Event ProgressChange(ByVal Progress As String)

Private Const MAX_CHUNK = 262144
Private sToken As String

Public Property Let Token(NewValue As String)
    ' Set token for the drive object
    sToken = NewValue
End Property

Public Property Get Token() As String
    ' Set token for the drive object
    Token = sToken
End Property

'Simple upload - best for small file to gooogle....
Function SimpleUpload(FileName As String) As Boolean
    Dim http As Object, url As String
    Dim fileData As String, rspTxt As String
    Dim fName As String
    If Token = "" Then Exit Function
    
    ' get only file name
    fName = GetFileName(FileName)
    ' Get token
    url = "https://www.googleapis.com/upload/drive/v2/files?uploadType=media"
    ' Read and get file
    fileData = ReadBinaryFile(FileName)
    
    Dim jSon As New clsJSONScript
        
    'Post the data To the destination URL
    'Create XMLHTTP/ServerXMLHTTP/WinHttprequest object
    'You can use any of these three objects.
    'Set http = CreateObject("WinHttp.WinHttprequest.5")
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        'Open URL As POST request
        .Open "POST", url, False
    
        'Set Content-Type header
        .SetRequestHeader "Authorization", Token
        .SetRequestHeader "Content-Type", GetMimeType(fName)
        .SetRequestHeader "Content-Length", GetFileSize(FileName)
  
        'Send the form data To URL As POST binary request
        .Send (fileData)
        
        rspTxt = .ResponseText
        If .Status <> 200 Then
            '// Error getting OAuth2 token
            'err.Raise vbObjectError + .Status, Description:="Failed to retrieve data " + .Status & ": " + .ResponseText
            SimpleUpload = False
            GoTo ErrHandler
        End If
        
        SimpleUpload = True
        ' Now we have to change the file name to the orginal file...
        url = "https://www.googleapis.com/drive/v2/files/" & GetFileID(rspTxt)
        .Open "PATCH", url, False
        .SetRequestHeader "Authorization", Token
        .SetRequestHeader "Content-type", "application/json"
        .Send (Utf8BytesFromString("{'title': '" & fName & "'}"))
        rspTxt = .ResponseText
        If .Status <> 200 Then
            '// Error getting OAuth2 token
            'err.Raise vbObjectError + .Status, Description:="Failed to retrieve data " + .Status & ": " + .ResponseText
            SimpleUpload = False
            Debug.Print .ResponseText
            GoTo ErrHandler
        End If
    End With
ErrHandler:
    Err.clear
    Set http = Nothing
    Set jSon = Nothing
End Function

Public Function UnicodeToAscii(sText As String) As String
  Dim x As Long, sAscii As String, ascval As Long
 
  If Len(sText) = 0 Then
    Exit Function
  End If
 
  sAscii = ""
  For x = 1 To Len(sText)
    ascval = AscW(Mid(sText, x, 1))
    If (ascval < 0) Then
      ascval = 65536 + ascval ' http://support.microsoft.com/kb/272138
    End If
    sAscii = sAscii & "&#" & ascval & ";"
  Next
  UnicodeToAscii = sAscii
End Function

Private Function GetFileID(JSONText As String)
    ' This will find the recently updated file
    Dim JsonScript As New clsJSONScript
    Dim jsObj As Object
    With JsonScript
        Set jsObj = .DecodeJsonString(JSONText)
        GetFileID = .GetProperty(jsObj, "id")
    End With
    Set JsonScript = Nothing
End Function

Private Function InsertFile(FileName As String, Optional ByVal bAsync As Boolean) As String
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim sPostData       As String
    Dim sUrl As String, Token As String
    ' Get token
    sUrl = "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart"
        
    Const Boundary = "314159265358979323846"
    Const delimiter = vbCrLf & "--" + Boundary + vbCrLf
    Const close_delim = vbCrLf + "--" + Boundary + "--"
 
    Dim base64Data As String, bodyRequest As String
    Dim MetaData As String, ContentType As String, body As String
    
    ' For reading file
    '--- read file
    nFile = FreeFile
    Open FileName For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    Close nFile
    
    ' add variable
    'contentType = "Application/octet-stream"
    ContentType = "application/msword"
    MetaData = "{'title':'" & GetFileName(FileName) & "','mimeType':'application/msword'}"
    MetaData = Replace(MetaData, "'", """")
    ' convert file to base64data
    'base64Data = Base64EncodeString(GetFile(filename))
    base64Data = GetFileText(FileName)
    
    body = delimiter
    body = body + "X-Upload-Content-Type: application/json" + vbCrLf + vbCrLf
    body = body + MetaData
    body = body + delimiter
    body = body + "Content-Type: " + ContentType + vbCrLf
    'body = body + "Content-Transfer-Encoding: base64" + vbCrLf
    'body = body + vbCrLf
    body = body + base64Data + close_delim
    
    '--- post
    Dim Xhr As Object
    Set Xhr = CreateObject("Microsoft.XMLHTTP")
    With Xhr
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Authorization", Token
        .SetRequestHeader "Content-Type", "multipart/related; boundary=""" + Boundary + """"
        .SetRequestHeader "Content-Length", GetFileSize(FileName)
        .SetRequestHeader "User-Agent", VbAgent
        .Send (body)
        If .Status <> 200 Then
            '// Error getting OAuth2 token
            Debug.Print .ResponseText
        End If
        If Not bAsync Then
            InsertFile = .ResponseText
        End If
        'WriteLog insertFile, "C:\dumpfile.txt"
    End With
    Set Xhr = Nothing
End Function

Private Function pvPostFile(sFileName As String, Optional ByVal bAsync As Boolean) As String
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim sPostData       As String
    Dim sUrl As String, Token As String, Xhr As Object
    sUrl = "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart"
        
    '--- read file
    nFile = FreeFile
    Open sFileName For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    Close nFile
    ' for meta data part
    Dim MetaHeader As Variant
    Dim sFormData As String, d As String, headerJson As String
    Const Boundary = "foo_bar_baz"
    Const delimiter = "\r\n--" + Boundary + "\r\n"
    Const close_delim = "\r\n--" + Boundary + "--"
    
    MetaHeader = Array("'title':'MyWordfile.doc'", "'description':'ntohing special'", "'mimeType':'application/msword'")
    '--- prepare body
    
    sPostData = "--" & Boundary & vbCrLf & sPostData & vbCrLf & "--" & Boundary & "--"
    sPostData = Join(MetaHeader) & vbCrLf & "--" & Boundary & vbCrLf & sPostData & vbCrLf & "--" & Boundary & "--"
        
    '--- post
    Set Xhr = CreateObject("Microsoft.XMLHTTP")
    With Xhr
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Authorization", Token
        .SetRequestHeader "Content-Type", "multipart/related; boundary='" + Boundary + "'"
        .SetRequestHeader "X-Upload-Content-Length", GetFileSize(sFileName)
        .SetRequestHeader "X-Upload-Content-Type", "application/msword"
        .Send pvToByteArray(sPostData)
        
        If Not bAsync Then
            pvPostFile = .ResponseText
        End If
        'WriteLog pvPostFile, "C:\dumpfile.txt"
    End With
exitfunction:
        Set Xhr = Nothing
End Function
 
Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Function UploadResumable(FileName As String, Optional FileID As String = "") As Boolean
    Dim params As String, Xhr As Object, res As String, urls As String, method As String, mimeType As String
    ' set up parameter for the upload
    If Token = "" Then Exit Function
    
    'Step 1: Start a resumable session
    'params = "alt=json"
    params = params + "&uploadType=resumable"
    urls = "https://www.googleapis.com/upload/drive/v2/files"
    
    If FileID <> "" Then
        params = params + "&setModifiedDate=true"
        urls = urls + FileID
        method = "PUT"
    Else
        method = "POST"
    End If
    ' this will breake parameter into ?uploadtype ... type?fiel=...&item=...
    urls = urls + "?" + params
    ' set up request header
    mimeType = GetMimeType(FileName)
    Set Xhr = CreateObject("MSXML2.XMLHTTP")
    With Xhr
        .Open method, urls, False
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .SetRequestHeader "Authorization:", Token
        .SetRequestHeader "X-Upload-Content-Type", mimeType
        '.Setrequestheader "X-Upload-Content-Length", GetFileSize(FileName)
        .SetRequestHeader "User-Agent", VbAgent
        RaiseEvent StatusChanged("Handshaking with Google Drive APIs...")
        
        .Send (Utf8BytesFromString("{'title': '" & GetFileName(FileName) & "'}"))
        Debug.Print .ResponseText
        ' check for error
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
        
    End With
    'Step 2: Save the resumable session URI
    urls = Xhr.getResponseHeader("Location")
    
    'Step 3: Upload the file
    method = "PUT"
    
    With Xhr
        .Open method, urls, False
        .SetRequestHeader "Authorization:", Token
        .SetRequestHeader "Content-Type", mimeType
        .SetRequestHeader "Content-Length", GetFileSize(FileName)
        .SetRequestHeader "User-Agent", VbAgent
        RaiseEvent StatusChanged("Start sending file...")
        
        .Send (ReadBinaryFile(FileName))
        'If the request succeeds, responds is HTTP 201 Created
        ', along with any metadata associated with this resource.
        'If the initial request of the resumable session had been a PUT, to update an existing resource,
        'the success response would be  200 OK, along with any metadata associated with this resource.

        'If the upload request is interrupted or if you receive an HTTP 503 Service Unavailable
        'or any other 5xx response from the server, follow the procedure outlined in resume an interrupted upload.
        Debug.Print .ResponseText
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
    End With
    RaiseEvent StatusChanged("File uploaded succesffully...")
    UploadResumable = True
exitfunction:
    Set Xhr = Nothing
End Function

Function DownloadMedia(FileID As String, SaveFileName As String) As Boolean
    Dim Xhr As Object, method As String, urls As String, xFileName As String
    If Token = "" Then Exit Function
    
    xFileName = GetFileName(SaveFileName)
    urls = "https://www.googleapis.com/drive/v2/files/"
    method = "+G+E+T+"
    urls = urls + FileID + "?alt=media"
    
    Set Xhr = CreateObject("MSXML2.XMLHTTP")
    With Xhr
        .Open Replace(method, "+", ""), urls, False
        .SetRequestHeader "Authorization:", Token
        .SetRequestHeader "Content-Type", GetMimeType(SaveFileName)
        RaiseEvent StatusChanged("Send request for downloading of file[" + xFileName + "]...")
        .Send
         'Wait for request to finish
        RaiseEvent StatusChanged("Waiting for responses...")
        Do While .readyState <> 4
            DoEvents
        Loop
        ' check for error
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
        RaiseEvent StatusChanged("Got the file, now saving...")
        ' get location uri if successfully
        WriteFileBinary SaveFileName, .ResponseBody
    End With
    RaiseEvent StatusChanged("Downloading of [" + xFileName + "] was successfull...")
    DownloadMedia = True
exitfunction:
    Set Xhr = Nothing
End Function

Function Download(FileID As String, SaveFileName As String) As Boolean
    Dim Xhr As Object, method As String, urls As String, xFileName As Stream
    Dim jSon As New clsJSONScript, jsObj As Object
    If Token = "" Then Exit Function
    
    xFileName = GetFileName(SaveFileName)
    
    urls = "https://www.googleapis.com/drive/v2/files/"
    method = "+G+E+T+"
    urls = urls + FileID + "?fields=downloadUrl,title"
    
    Set Xhr = CreateObject("MSXML2.XMLHTTP")
    
    With Xhr
        ' First step to get download url
        .Open Replace(method, "+", ""), urls, False
        .SetRequestHeader "Authorization:", Token
        RaiseEvent StatusChanged("Send request for downloading of file[" + xFileName + "]...")
        
        .Send
         'Wait for request to finish
        WaitMe Xhr
                
        ' check for error
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
        Set jsObj = jSon.DecodeJsonString(.ResponseText)
        ' now start download the file
        
        RaiseEvent StatusChanged("Retrieving download url...")
        urls = jSon.GetProperty(jsObj, "downloadUrl")
        Set jSon = Nothing
        
        'Now using that url and download the file
        .Open Replace(method, "+", ""), urls, False
        .SetRequestHeader "Authorization:", Token
        RaiseEvent StatusChanged("Request for download...")
        .Send
         'Wait for request to finish
        WaitMe Xhr
        '
        ' check for error
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
        
        'Debug.Print .getAllResponseHeaders()
        ' Now save file
        RaiseEvent StatusChanged("Downloaded, saving file...")
        WriteFileBinary SaveFileName, .ResponseBody
    End With
    RaiseEvent StatusChanged("Download of file [" + xFileName + "] completed...")
    Download = True
exitfunction:
    Set Xhr = Nothing
End Function

Private Sub WaitMe(Obj As Object)
    Do While Obj.readyState <> 4
        DoEvents
    Loop
End Sub

Private Sub Tried()
    Dim d As String
    'for parsing parametter to query
    d = "{'q': isRoot=true, 'fields': 'items(id,title,mimeType)}"
End Sub

Function UploadChunk(FileName As String, Optional FileID As String = "") As Boolean
    Dim params As String, Xhr As Object, res As String, urls As String, method As String, mimeType As String
    ' set up parameter for the upload
    If Token = "" Then Exit Function
    
    'Step 1: Start a resumable session
    'params = "alt=json"
    params = params + "&uploadType=resumable"
    urls = "https://www.googleapis.com/upload/drive/v2/files"
    
    If FileID <> "" Then
        params = params + "&setModifiedDate=true"
        urls = urls + FileID
        method = "PUT"
    Else
        method = "POST"
    End If
    ' this will breake parameter into ?uploadtype ... type?fiel=...&item=...
    urls = urls + "?" + params
    ' set up request header
    mimeType = GetMimeType(FileName)
    Set Xhr = CreateObject("MSXML2.XMLHTTP")
    With Xhr
        .Open method, urls, False
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .SetRequestHeader "Authorization:", Token
        .SetRequestHeader "X-Upload-Content-Type", mimeType
        '.Setrequestheader "X-Upload-Content-Length", GetFileSize(FileName)
        .SetRequestHeader "User-Agent", VbAgent
        
        RaiseEvent StatusChanged(Msg("MSG_REQUESTING_UPLOAD_HANDLE"))
        .Send (Utf8BytesFromString("{'title': '" & GetFileName(FileName) & "'}"))
        
        ' check for error
        If .Status <> 200 Then
            ' check for error type and response to user
            GoTo exitfunction
        End If
        
    End With
    'Step 2: Save the resumable session URI
    urls = Xhr.getResponseHeader("Location")
    RaiseEvent StatusChanged(Msg("MSG_UPLOAD_HANDLE_RETRIEVED"))
    
    'Step 3: Upload the file
    If PostBinaryChunks(Xhr, urls, FileName) Then
        ' successfully uploaded all
        UploadChunk = True
        RaiseEvent StatusChanged(Replace(Msg("MSG_UPLOAD_DONE_SUCCESSFULL"), "%%", "[" + FileName + "]"))
    Else
        ' failed at some point...
        RaiseEvent StatusChanged(Msg("MSG_UPLOAD_FAILED"))
    End If
exitfunction:
    Set Xhr = Nothing
End Function

Private Function PostBinaryChunks(xhrObject As Object, url As String, FileName As String, Optional ChunkSize As Long = MAX_CHUNK) As Boolean
    '// Push 'intChunkSize' chunks to client. When near EOS, Read() will terminate at EOS regardless of being asked to read passed EOS
    Dim objStream As Object, bntCount As Long, uFileSize As Long, uFileMimeType As String
    Dim totalTrunks As Long, i As Long, RmnByte As Long
    Const Retries = 5
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Open
        .Type = adTypeBinary
        .LoadFromFile FileName
        uFileSize = GetFileSize(FileName)
        uFileMimeType = GetMimeType(FileName)
        totalTrunks = (uFileSize / ChunkSize) + 1
ResumeUpload:
        While Not .EOS
            ' upload in trunk
            DoEvents
            If MakeRequest(xhrObject, url, bntCount * ChunkSize, .Read(ChunkSize), uFileSize, uFileMimeType) Then
                RaiseEvent ProgressChange(Format(100 * (bntCount + 1) / totalTrunks, "Standard") & "%")
            Else
                ' Failed at some point, now get the remained byte and tried 5 times
                While i <= Retries
                    Exponential_backoff
                    RmnByte = QuerryRemainBytes(xhrObject, url, uFileSize)
                    If RmnByte > 0 Then
                        ' good - server on again
                        objStream.Position = RmnByte + 1
                        i = Retries
                        RaiseEvent StatusChanged(Replace(Msg("MSG_RESUME_FAILED_UPLOAD"), "%%", "[" + FileName + "]"))
                        GoTo ResumeUpload
                    End If
                    i = i + 1
                Wend
                ' seemed no luck after retrying... tell permanent error
                RaiseEvent StatusChanged(Replace(Msg("MSG_UPLOAD_CRITICAL_ERROR_ABORTED"), "%%", "[" + FileName + "]"))
                GoTo ExitProc
            End If
            ' get return number to re-read
            bntCount = bntCount + 1
        Wend
        PostBinaryChunks = True
        RaiseEvent StatusChanged("Uploading completed...")
    End With
ExitProc:
    Set objStream = Nothing
End Function

Private Function Exponential_backoff()
    Sleep 1000 + Int(Rnd(1) * 100)
    'Wait 1 second + random_number_milliseconds and retry the request.
End Function

Private Function MakeRequest(xhrObject As Object, url As String, _
    start As Long, body As Variant, Optional FileLength As Long, _
    Optional strMimeType As String) As Boolean
    ' this function will make a range request header for chunk uploading and forward the data
    Dim reg As Object, reqSize As Long
    On Error GoTo ErrProc
    With xhrObject
        .Open "PUT", url, False
        '.Setrequestheader "Authorization:", Token
        ' set length of transfered bytes
        .SetRequestHeader "Content-Length", MAX_CHUNK
        ' still mimetype
        .SetRequestHeader "Content-Type", strMimeType
        reqSize = UBound(body) + 1
        If reqSize <> 0 Then
            .SetRequestHeader "Content-Range", "bytes " & start & "-" & start + reqSize - 1 & "/" & FileLength
        Else
            .SetRequestHeader "Content-Range", "bytes */" & FileLength
        End If
        .SetRequestHeader "User-Agent", VbAgent
        
        .Send (body)
                
        Select Case .Status
        Case 308: ' still continuing
            ' Remember the current chunk for later resuming of the upload
            Debug.Print .getResponseHeader("Range")
        Case 201, 200: ' finish the file
            Debug.Print .getAllResponseHeaders
        Case 500, 502, 503, 504: 'may be an error
            ' tell the uploader to retry later
            GoTo ErrProc
        End Select
    End With
    
    MakeRequest = True
ErrProc:
End Function

Private Function QuerryRemainBytes(xhrObject As Object, url As String, FileSize As Long) As Long
    ' This function will query Google to get the remained bytes was not uploaded
    Dim tmpStr As String
    'On Error GoTo ErrProc
    With xhrObject
        ' Just follow google guide on this
        'PUT {session_uri} HTTP/1.1
        'Content-Length: 0
        'Content-Range: bytes */2000000

        .Open "PUT", url, False
        .SetRequestHeader "Content-Length", 0
        .SetRequestHeader "Content-Range", "bytes */" & FileSize
        .SetRequestHeader "User-Agent", VbAgent
        .Send
        
        'If successfully, following is the returned data
        'HTTP/1.1 308 Resume Incomplete
        'Content-Length: 0
        'Range: 0-42
        Select Case .Status
        Case 308: ' still continuing
            ' Remember the current chunk for later resuming of the upload
            tmpStr = .getResponseHeader("Range")
            QuerryRemainBytes = Val(Mid(tmpStr, InStr(tmpStr, "-") + 1))
        Case 201, 200: ' finish the file
            
        Case 500, 502, 503, 504: 'may be an error
            ' tell the uploader to retry later
            GoTo ErrProc
        End Select
    End With
ErrProc:
End Function
