Option Compare Database
Option Explicit
Private Token As String, frmImportChoice As Long
Private WithEvents gdrive As clsGdrive
Private ParentFolder As String
Private fLang As String

Private Sub initGdrive()
    If gdrive Is Nothing Then Set gdrive = New clsGdrive
    With gdrive
        .Token = GetToken
    End With
End Sub

Private Sub btnDownload_Click()
    'Implement the downloading of the file...
    Dim SavePath As String
    SavePath = CurrentProject.path
    ' Get thefile id and download
    If nz(lstChildren, "") = "" Then Exit Sub
    
    'Browse for folder to save file
    SavePath = BrowseFileOrDir(False)
    
    If Not FileOrDirExists(SavePath) Then Exit Sub
    initGdrive
    DoCmd.Hourglass True
    ' there two type of downloading we implement in this demo
    ' Downdload media directly and download using generated dowload url
    
    gdrive.DownloadMedia lstChildren, SavePath + "\" + lstChildren.column(1)
    'Download lstParent.column(0), GetToken, SavePath + "\" + lstParent.column(1)
    DoCmd.Hourglass False
End Sub

Private Sub btnList_Click()
    On Error GoTo errorhandler
    DoCmd.Hourglass True
    initGdrive
    RefreshList
    DoCmd.Hourglass False
    btnUpload.Enabled = True
    btnDownload.Enabled = True
    btnDelete.Enabled = True
    
    Exit Sub
errorhandler:
    MsgBox "Error: " & Err.Description
    
End Sub

Function RefreshList(Optional ParentID As String = "")
    'On Error GoTo errorhandler
    lstParent.RowSource = ""
    ' Now we have to list file
    Dim jSon As New clsJSONScript, jsObj As Object, myArr As Variant
    If ParentID = ".." Then
        ParentID = ParentFolder
    Else
        ParentFolder = ParentID ' for later problem
    End If
    With jSon
        Set jsObj = .DecodeJsonString(RetrieveGdriveList(GetToken, ParentID, True))
        ' put this stuff.. to array
        myArr = .JsonArray
    End With
    Dim i As Integer, j As Long, strOut_p As String, strOut_c As String, UseCol As Variant, folderItem As String
    With lstParent
        .RowSource = ""
        ' we just take the Id, title, mimetype, trash, parentID
        'UseCol = Array(0, 1, 2, 4, 5)
        .ColumnCount = 2
        .ColumnWidths = "0;" & .Width
        .BoundColumn = 1
        ' Only show name of object
        For j = 1 To UBound(myArr, 1)
            If UBound(myArr, 2) = 0 Then Exit For
            If InStr(myArr(j, 2), "folder") > 0 Then
                ' only list folder
                strOut_p = strOut_p + ";" + nz(myArr(j, 0), "") + ";" + nz(myArr(j, 1), "")
            End If
            ' for children list
            strOut_c = strOut_c + ";" + nz(myArr(j, 0), "''") + ";" + nz(myArr(j, 1), "''") + ";" + nz(myArr(j, 2), "''")
        Next j
        strOut_p = Mid(strOut_p, 2)
        If ParentID <> "" And ParentID <> "root" Then strOut_p = "root;..;" & strOut_p
                
        ' .. for the root
        .RowSource = strOut_p
        
    End With
    ' now we have to get the content of this selected object
    With lstChildren
        .RowSource = ""
        .RowSourceType = "Value list"
    
        ' we just take the Id, title, mimetype, trash, parentID
        'UseCol = Array(0, 1, 2, 4, 5)
        .ColumnCount = 3
        .ColumnWidths = "0;" & .Width / 2 & ";" & Width / 2
        .ColumnHeads = False
        .BoundColumn = 1
                
        strOut_c = Mid(strOut_c, 2)
        ' .. for the root
        .RowSource = strOut_c
    End With
errorhandler:
    Set jSon = Nothing
End Function

Private Sub btnUpload_Click()
    'On Error GoTo errorhandler
    'This will activate google drive to upload file...
    ' Create a zip file first
    Dim mdePath As String, mdefileName As String, zipFile As String, filePath As String
    frmImportChoice = 1
    
    filePath = BrowseFileOrDir
    
    If Not FileOrDirExists(filePath, True) Then Exit Sub
    mdePath = GetFilePath(filePath)
    DoCmd.Hourglass True
    DoEvents
    initGdrive
    'If SimpleUpload(filePath) Then RefreshList
    'If gdrive.UploadResumable(filePath) Then RefreshList
    If gdrive.UploadChunk(filePath) Then RefreshList
    
    DoCmd.Hourglass False
errorhandler:
    'MsgBox "Error: " & Err.Description
End Sub

Private Sub cmdSettings_Click()
    ' Call the setting up of connection variable
    DoCmd.OpenForm "frmSettings", acNormal, , , , acDialog
    Token = ""
    If AppLanguage <> fLang Then
        fLang = AppLanguage
        SetObjInterface = Me
    End If
End Sub

Private Sub Form_Load()
    SetObjInterface = Me
    fLang = AppLanguage
    lstParent.RowSourceType = "Value list"
    lstParent.RowSource = ""
    With Me
        .InsideHeight = 5000
        .InsideWidth = 7200
    End With
End Sub

Private Function GetToken() As String
    If Token = "" Then Token = GetAuthCode
    GetToken = Token
End Function

Private Sub Form_Resize()
    ' changing size for stuff...
    Dim CrtHeight As Long
    With Me
        CrtHeight = .InsideHeight - .lstChildren.Top - 50
        .lstParent.Height = CrtHeight
        .lstChildren.Height = CrtHeight
        .lstChildren.Width = .InsideWidth - lstChildren.Left - 50
    End With
    DoCmd.Maximize
End Sub

Private Sub gdrive_ProgressChange(ByVal Progress As String)
    AppStatus Replace(Msg("MSG_ACTION_IN_PROGRESS"), "%%", " [" + Progress + "] ")
End Sub

Private Sub gdrive_StatusChanged(ByVal Status As String)
    AppStatus Status
End Sub

Private Sub lstParent_Click()
    Debug.Print lstParent.column(0)
End Sub

Private Sub lstParent_DblClick(Cancel As Integer)
    ' now try to list object on the left folder
    If nz(lstParent, "") = "" Then Exit Sub
    ' Now retrive item under this parrent
    RefreshList lstParent
End Sub
