Option Compare Database
Option Explicit

Private WithEvents gAuth As gAuth2
Private RefreshKey As String ' this will be of help for anycomputer without this key

Private Sub cmdExit_Click()
    Frame24_AfterUpdate
    DoCmd.Close acForm, Me.name, acSaveYes
End Sub

Private Sub cmdGetToken_Click()
    'this will create new config file for rclone
    If Not CheckAllVariable Then Exit Sub
    DoCmd.Hourglass True
    txtToken = GetOAuthCode()
    If txtToken = "" Then GoTo Exit_Sub
    AppProperty("App_client_id", "Google Drive") = txtClientID
    AppProperty("App_client_secret", "Google Drive") = txtClientSecret
Exit_Sub:
    DoCmd.Hourglass False
End Sub

Private Function GetOAuthCode() As String
    ' Get token for accessing drive object
    Set gAuth = New gAuth2
    With gAuth
        .InitClientCredentials txtClientID, txtClientSecret
        .InitEndPoints
        '.LogOnGoogle txtGAccount, txtGPassword
        If .LoginGoogle(txtGAccount, txtGPassword) Then
            If Not .GetNewToken() Then GoTo Exit_Function
            GetOAuthCode = .AuthHeader
        End If
    End With
Exit_Function:
    Set gAuth = Nothing
End Function

Private Sub cmdRefresh_Click()
    'If Not CheckAllVariable Then Exit Sub
    Set gAuth = New gAuth2
    With gAuth
        .InitClientCredentials txtClientID, txtClientSecret, RefreshKey
        .InitEndPoints
        If Not .RefreshToken() Then GoTo Exit_Function
        txtToken = .AuthHeader
    End With
    AppProperty("App_client_id", "Google Drive") = txtClientID
    AppProperty("App_client_secret", "Google Drive") = txtClientSecret
Exit_Function:
    Set gAuth = Nothing
End Sub

Private Sub Form_Load()
    Dim client_id As String, client_secret As String
    'set language
    SetObjInterface = Me
    Me.Frame24 = IIf(DBConfig("APP_LANGUAGE") = "E", 1, 2)
    With Me
        .InsideWidth = 6400
        .InsideHeight = 2500
        ' get from ini the setting if any
        client_id = AppProperty("App_client_id", "Google Drive")
        client_secret = AppProperty("App_client_secret", "Google Drive")
        'always use this refreshkey
        RefreshKey = "1/zfqbZXDifkf5teGbUXGFl2gn5QA2F_7MDIYBtdpsWFzBactUREZofsF9C7PrpE-j"
        If client_id = "N/A" Then
            ' load demo profile
            .txtClientID = "759747656687-rjkm22bit7ob5tufc5sbgg1gsuj48fme.apps.googleusercontent.com"
            .txtClientSecret = "Jf5DqXlZ2G3cOtUtHIraaxvQ"
            .txtGAccount = "khosolieu.app@gmail.com"
            .txtGPassword = "ksl12345"
        Else
            .txtClientID = client_id
            .txtClientSecret = client_secret
        End If
    End With
End Sub

Private Function CheckAllVariable() As Boolean
    Dim Obj As Object
    CheckAllVariable = True
    For Each Obj In Me.Controls
        If Obj.Tag = 1 Then
            If nz(Obj, "") = "" Then
                MsgBox Msg("MSG_SOME_OBJECT_MISSING"), vbInformation
                CheckAllVariable = False
                Exit For
           End If
        End If
    Next
End Function

Private Sub Frame24_AfterUpdate()
    If Frame24 = 1 Then DBConfig("APP_LANGUAGE") = "E" Else DBConfig("APP_LANGUAGE") = "V"
End Sub

Private Sub gAuth_StatusChanged(ByVal Status As String)
    AppStatus Status
End Sub
