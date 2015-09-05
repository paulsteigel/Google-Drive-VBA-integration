' ================================================================================== '
'
' OAuth 2.0 Google Authenticator
' Developed by Kyle Beachill
' licence: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' Inspired loosely by Tim Halls authentication classes in his Excel-Rest library:
' https://github.com/timhall/Excel-REST
' modified by Dang Dinh Ngoc with inclusion of WinhttpRequest protocol on replacement of InternetExplorer
' Vietnam 2015
'
' Features:
'     Simple class to handle Google OAuth 2.0 Authentication
'     Follows the Installed Application Flow
'     Returns Simply the value for the Authorization header in API requests
'
' Gotchas:
'     Tokens are held in plain text in the registry
'
' Required References:
'   - Microsoft Internet Controls ' no longer need
'   - Microsoft XML
'
' ================================================================================== '


Option Compare Database
Option Explicit

Public Event StatusChanged(ByVal Status As String)

'// Simple enum for current authentication status
Private Enum AuthenticationStatus
    NotAuthenticated = 1
    TokenExpired = 2
    Authenticated = 3
End Enum

Private Enum HtmlElementType
    ById = 0
    ByName = 1
    ByClassName = 2
    ByTagName = 3
End Enum

'// Application Client ID and Application Secret
Private strClientId As String
Private strClientSecret As String

'// Authentication codes, tokens and expiry date
Private strTokenKey As String
Private strToken As String
Private strRefreshToken As String
' this is a temporary key for using if missing
Private tmpstrRefreshToken As String

Private dtExpiresWhen As Date
Private strAuthCode As String

'// Url End points for the authentication
Private strAuthUrl As String
Private strTokenUrl As String
Private strRedirectUri As String

'// Internet Explorer variables for initial authentication request
Private oIExplorer As Object

Private strResponseText As String
Private oResponse As Object

'// Save the request object to prevent being created for each token expiry
Private objXMLRequest As Object

Private Const WAIT_TIMEOUT = 300
Private Const ERR_TIMEOUT = 1000
Private Const READYSTATE_COMPLETE = 4
Private Const PAGE_LOADED = 4
Private Xhr As Object
Private LogonCookie As Variant

'// Since we are persisting the credentials to the registry, we need to read these in each time the class
'// is initialized, if they aren't found - these will be default values, "" for strings and 1900/01/01 for te date
Private Sub Class_Initialize()
    
    Dim sDate As String
    strToken = GetSetting("GoogleAuth", "Tokens", "Token")
    strRefreshToken = GetSetting("GoogleAuth", "Tokens", "RefreshKey")
    sDate = GetSetting("GoogleAuth", "Tokens", "TokenExpiry")
    
    If Len(sDate) > 0 Then
         dtExpiresWhen = CDate(sDate)
    Else
         dtExpiresWhen = #1/1/1900#
    End If
    
End Sub

'// Allows the overriding of the default google EndPoints - these are unlikely to change
Public Sub InitEndPoints( _
    Optional ByVal AuthUrl As String = "https://accounts.google.com/o/oauth2/auth", _
    Optional ByVal TokenUrl As String = "https://accounts.google.com/o/oauth2/token", _
    Optional ByVal RedirectUri As String = "urn:ietf:wg:oauth:2.0:oob" _
)
    
    strAuthUrl = AuthUrl
    strTokenUrl = TokenUrl
    strRedirectUri = RedirectUri
End Sub

'// Application ID and Secret will always need passing, since they are required for refresh calls
'// Though these *could* be persisted in the registry also
Public Sub InitClientCredentials(ByVal ClientId As String, ByVal ClientSecret As String, Optional TmpRefeshKey As String = "")
    strClientId = ClientId
    strClientSecret = ClientSecret
    tmpstrRefreshToken = TmpRefeshKey
    ' using application temporary key
    If strRefreshToken = "" Then strRefreshToken = tmpstrRefreshToken
End Sub

'// Simple function to return the authentication status of the currently held credentials
Private Function getAuthenticationStatus() As AuthenticationStatus
        
    '// If the Refresh Token Length is 0 then the initial authentication hasn't occurred
    If Len(strRefreshToken) = 0 Then
        getAuthenticationStatus = NotAuthenticated
        Exit Function
    End If
    
    '// If the refresh date is less than now (with a 10 second buffer) then the token has expired
    If dtExpiresWhen < DateAdd("s", 10, Now()) Then
        getAuthenticationStatus = TokenExpired
        Exit Function
    End If
    
    '// Otherwise the token is valid
    getAuthenticationStatus = Authenticated
End Function

Function LoginGoogle(gAccountName As String, gPassword As String) As Boolean
    Dim oHttp As Object, sHTML As String
    Dim google_accounts_url As String, authentication_url As String, gmailUrl As String, i As Long
    Dim d As String, google_accounts_logout
    
    google_accounts_logout = "https://accounts.google.com/Logout?service=accountsettings"
    google_accounts_url = "http://accounts.google.com"
    authentication_url = "https://accounts.google.com/ServiceLoginAuth"
    gmailUrl = "https://mail.google.com/mail/"
        
    GetWebConnector True
    
    Set oHttp = Xhr
    Dim Cookie As Variant, UserDetails As String
    
    With oHttp
        ' Log out all first
        .Open "POST", google_accounts_logout, False
        .Send
        RaiseEvent StatusChanged(Msg("MSG_LOG_OUT_ALL_ACCOUNTS"))
        
        If .Status <> 200 Then GoTo Exit_Sub
        
        ' now release a get request
        .Open "GET", google_accounts_url, False
        .Send
        RaiseEvent StatusChanged(Msg("MSG_SEND_REQUEST_FOR_LOGGING_IN"))
        
        If .Status <> 200 Then GoTo Exit_Sub
        
        'Debug.Print .getAllResponseHeaders()
        'Cookie = Split(.getResponseHeader("Set-Cookie"), ";")
        
        d = "&GALX=" + GetResponseData(.ResponseText, "GALX", , 0)
        d = d + "&continue=https://www.google.com/?gws_rd=ssl&hl=en&_utf8=?"
        d = d + "&bgresponse=js_disabled&pstMsg=1&dnConn=&checkConnection=youtube:206:1"
        d = d + "&checkedDomains=youtube"
        d = d + "&Email=" + gAccountName + "&Passwd=" + gPassword
        d = d + "&PersistentCookie=Yes&signIn=Sign in"
                
        .Open "POST", authentication_url, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .SetRequestHeader "User-Agent", VbAgent '"Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
        .SetRequestHeader "Host", "accounts.google.com"
        .SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        .SetRequestHeader "Accept-Language", "en-US,en;q=0.5"
        
        'For i = 0 To UBound(Cookie)
        '    .setRequestHeader "Cookie", Cookie(i)
        'Next
        .SetRequestHeader "Connection", "keep-alive"
        .SetRequestHeader "Referer", "https://accounts.google.com/ServiceLoginAuth"
        .SetRequestHeader "Content-Length", Len(d)
        
        .Send (Utf8BytesFromString(d))
        RaiseEvent StatusChanged(Msg("MSG_SUBMITTED_USER_DATA"))
        
        If .Status <> 200 Then GoTo Exit_Sub
        
        'WriteLog .ResponseText, "C:\Logged_account.html", True
        
        ' now check for logiing in
        d = GetResponseData(.ResponseText, "gaia_loginform", "action", 0)
        If d = "" Then
            If GetResponseData(.ResponseText, "title", , 3) = "Google Accounts" Or InStr(.ResponseText, gAccountName) > 0 Then
                RaiseEvent StatusChanged(Msg("MSG_LOGGED_IN"))
                LoginGoogle = True
                GoTo Exit_Sub
            End If
        Else
            ' Check for what was the message
            d = GetResponseData(.ResponseText, "errormsg_0_Email", , 0, True)
            If d <> "" Then GoTo Exit_With_Error
            ' password or email problem
            d = GetResponseData(.ResponseText, "errormsg_0_Passwd", "innerHtml", 0, True)
        End If
    End With
Exit_With_Error:
    RaiseEvent StatusChanged(Msg("MSG_LOG_IN_FAILED") + " [" + d + "]")
Exit_Sub:
End Function

Private Function GetResponseData(ResponseText As String, ItemID As String, _
    Optional PropertyType As String = "value", Optional ItemType As HtmlElementType = 0, _
    Optional GetFirstChildElement As Boolean = False)
    ' This function will check whether the surfing was successfully or not?
    Dim htmlDoc As Object, FindElement As Object, tmpObj As String
    Set htmlDoc = CreateObject("HtmlFile")
    
    ' Load the html object
    htmlDoc.Write ResponseText
    htmlDoc.Close

    With htmlDoc
        Select Case ItemType
        Case 0: ' by id
            Set FindElement = .getElementById(ItemID)
        Case 1: ' by name
        Case 2: ' by class
            Set FindElement = .getElementsByClassName(ItemID)
        Case 3: ' by TagName
            ' jst dont know why every time using by tagname IE show up
            GetResponseData = htmlDoc.title
            'Set FindElement = .getElementsByTagName(ItemID)
        End Select
    End With
    If FindElement Is Nothing Then
        ' not existent - succesffully or mightbe not
    Else
        On Error Resume Next
        ' first - try to get its property if we can
        GetResponseData = FindElement.getAttribute(PropertyType)
        If GetFirstChildElement Then
            If FindElement.hasChildNodes Then
                GetResponseData = FindElement.firstChild.getAttribute(PropertyType)
                While FindElement.children.Length >= 1
                    FindElement.removeChild (FindElement.lastChild)
                Wend
                If FindElement.getAttribute(PropertyType) <> "" Then GetResponseData = FindElement.getAttribute(PropertyType)
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.clear
    Set htmlDoc = Nothing
End Function

Sub LogOnGoogle(gAccount As String, gPassword As String)
    GetWebConnector
    
    ' Log out all first
    RaiseEvent StatusChanged(Msg("MSG_LOG_OUT_ALL_ACCOUNTS"))
    SurfAndWait oIExplorer, "https://accounts.google.com/Logout?service=accountsettings"
        
    'Browse to service login
    SurfAndWait oIExplorer, "https://accounts.google.com/ServiceLogin"
        
    If ElementAvaiable(oIExplorer, "account-chooser-add-account") Then
        'Choose add acount
        ClickButton oIExplorer, "account-chooser-add-account"
        SurfAndWait oIExplorer, , True
    End If
    
    
    'Check for type of loging system
    With oIExplorer
        'enter email
        .Document.all.Email.Value = gAccount
        If Not ElementAvaiable(oIExplorer, "Passwd") Then
            ' two step logging in
            ClickButton oIExplorer, "next"
            SurfAndWait oIExplorer, , True
            .Document.all.Passwd.Value = gPassword
        Else
            .Document.all.Passwd.Value = gPassword
        End If
        ' submit and wait
        Call oIExplorer.Document.all.gaia_loginform.submit
    End With
    RaiseEvent StatusChanged(Msg("MSG_LOGGED_IN"))
    SurfAndWait oIExplorer, , True
End Sub

Private Sub SurfAndWait(ieObject As Object, Optional url As String, _
    Optional SkipSurfing As Boolean = False, _
    Optional GetTitle As Boolean = False)
    If Not SkipSurfing Then Call oIExplorer.Navigate(url)
    If Not GetTitle Then Do: DoEvents: Loop Until WaitUntilLoaded Else Do: DoEvents: Loop Until BrowserComplete
End Sub

Private Function WaitUntilLoaded() As Boolean
    Dim i, j, ready

    ' wait for page to connect
    i = 0
    Do Until oIExplorer.readyState = READYSTATE_COMPLETE
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then Err.Raise ERR_TIMEOUT, , "Timeout"
    Loop

    ' wait for document to load
    Do Until oIExplorer.Document.readyState = "complete"
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then Err.Raise ERR_TIMEOUT, , "Timeout"
    Loop
    
    WaitUntilLoaded = True
End Function

Function GetAuthorisationCode(url As String) As String
    Dim oDOM As Object, frmAction As Object
    Dim ApproveUrl As String, d As String, i As Long, Cookie As Variant
    Dim tmpStr As String, xPos1 As Long, xPos2 As Long
    
    With Xhr
        ' Navigate the consent windows
        .Open "GET", url, False
        .Send
        If .Status <> 200 Then GoTo Exit_Sub
        
        ' Get the submited address
        ApproveUrl = GetResponseData(.ResponseText, "connect-approve", "action")
        
        d = "&bgresponse=''&_utf8=" + ChrW(9731)
        d = d + "&state_wrapper=" + GetResponseData(.ResponseText, "state_wrapper")
        d = d + "&submit_access=true"
        d = d + "&valuenow=true"
        
        .Open "POST", ApproveUrl, False
        
        .SetRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
        '.SetRequestHeader "Accept-Encoding", "gzip, deflate"
        .SetRequestHeader "Accept-Language", "en-US"
        '.setRequestHeader "Accept-Language", "en-US,en;q=0.5"
        .SetRequestHeader "Cache-Control", "no-cache"
        .SetRequestHeader "Connection", "Keep-Alive"
        .SetRequestHeader "Content-Length", Len(d)
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        
        .SetRequestHeader "Host", "accounts.google.com"
        .SetRequestHeader "Referer", url
        .SetRequestHeader "User-Agent", VbAgent
                
        .Send (Utf8BytesFromString(d))
                
        If .Status <> 200 Then GoTo Exit_Sub
        GetAuthorisationCode = GetResponseData(.ResponseText, "code")
        RaiseEvent StatusChanged(Msg("MSG_GOT_THE_CODE"))
    End With
Exit_Sub:
    Set oDOM = Nothing
End Function

Private Function Escape(ByVal url As String) As String
    'URLs cannot contain most special characters.
    'VBScript and JavaScript have built-in Escape functions. In VB we have to write our own

    Dim i As Integer, BadChars As String
    BadChars = "<>%=&!@#£$^()+{[}]|\;:'"",/?"
    For i = 1 To Len(BadChars)
        url = Replace(url, Mid(BadChars, i, 1), "%" & Hex(Asc(Mid(BadChars, i, 1))))
    Next i
    url = Replace(url, " ", "+")
    Escape = url
    
End Function

Function GetNewToken() As Boolean
    GetWebConnector True
    '// Wait for userInteraction
    'SurfAndWait oIExplorer, CreateAuthRequest(), , True
    strAuthCode = GetAuthorisationCode(CreateAuthRequest())
    '// Do we have an Authentication Code?
    If Len(strAuthCode) = 0 Then Exit Function
               
    With Xhr
        .Open "POST", strTokenUrl, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send CreateTokenRequest()

        If .Status <> 200 Then Exit Function
        
        '// Get the credentials from the response
        strToken = GetProp("access_token", .ResponseText)
        strRefreshToken = GetProp("refresh_token")
        dtExpiresWhen = DateAdd("s", CLng(GetProp("expires_in")), Now())
    End With
    
    RaiseEvent StatusChanged(Msg("MSG_TOKEN_RENEWED"))
    
    '// Persist the Refresh key and expiry - the above should only ever need running once per application
    SaveSetting "GoogleAuth", "Tokens", "RefreshKey", strRefreshToken
    SaveSetting "GoogleAuth", "Tokens", "Token", strToken
    SaveSetting "GoogleAuth", "Tokens", "TokenExpiry", CStr(dtExpiresWhen)
    GetNewToken = True
End Function

Function RefreshToken() As Boolean
    GetWebConnector True
    Dim url As String
    url = CreateRefreshRequest()
    With Xhr
        .Open "POST", strTokenUrl, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send url
        If .Status <> 200 Then Exit Function
        
        '// Get the credentials from the response
        strToken = GetProp("access_token", .ResponseText)
        dtExpiresWhen = DateAdd("s", CLng(GetProp("expires_in")), Now())
    End With
    RaiseEvent StatusChanged(Msg("MSG_TOKEN_REFRESHED"))
    RefreshToken = True
    '// Persist new token in registry
    SaveSetting "GoogleAuth", "Tokens", "RefreshKey", strRefreshToken
    SaveSetting "GoogleAuth", "Tokens", "Token", strToken
    SaveSetting "GoogleAuth", "Tokens", "TokenExpiry", CStr(dtExpiresWhen)
End Function

'// Simple function that gets a propery from a single depth JSON formatted string
'// Requires the property name
'// Requires te JSON string on the first pass
Private Function GetProp(strPropName As String, Optional strJSObject As String = "") As String
    Static oScriptControl As Object
    
    If oScriptControl Is Nothing Then Set oScriptControl = CreateObject("ScriptControl")
    With oScriptControl
        .Language = "JScript"
        .AddCode "function getProp(json, prop) { return json[prop]; }"
        
        If Len(strJSObject) > 0 Then
            strResponseText = strJSObject
            Set oResponse = .eval("(" & strJSObject & ")")
        End If
        GetProp = .Run("getProp", oResponse, strPropName)
    End With
End Function

'// Public property to return the Authorisation value header for a request
Public Property Get AuthHeader() As String
    Dim eAuthStatus As AuthenticationStatus
    eAuthStatus = getAuthenticationStatus
    
    If eAuthStatus = NotAuthenticated Then
        Call GetNewToken
    ElseIf eAuthStatus = TokenExpired Then
        RefreshToken
    End If
    AuthHeader = "Bearer " & strToken
End Property

'//===========================================================================================================
'// String building functions for the requests

'// Step 1: The initial url for authentication - Note the scope attribute, this sets what the application can access
Private Function CreateAuthRequest() As String
    ' Generate initial Authentication Request
    ' Using installed application flow: https://developers.google.com/accounts/docs/OAuth2InstalledApp
    CreateAuthRequest = strAuthUrl
    If InStr(1, CreateAuthRequest, "?") < 1 Then: CreateAuthRequest = CreateAuthRequest & "?"
    CreateAuthRequest = CreateAuthRequest & "client_id=" & strClientId
    CreateAuthRequest = CreateAuthRequest & "&redirect_uri=" & strRedirectUri
    CreateAuthRequest = CreateAuthRequest & "&response_type=code"
    CreateAuthRequest = CreateAuthRequest & "&scope=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fdrive&state=state"
End Function

'// Step 2: The initial POST body to get the initial Token and refresh token
Private Function CreateTokenRequest() As String
    CreateTokenRequest = "code=" & strAuthCode
    CreateTokenRequest = CreateTokenRequest & "&client_id=" & strClientId
    CreateTokenRequest = CreateTokenRequest & "&client_secret=" & strClientSecret
    CreateTokenRequest = CreateTokenRequest & "&redirect_uri=" & strRedirectUri
    CreateTokenRequest = CreateTokenRequest & "&grant_type=authorization_code"
End Function

'// Step 3: The POST body to refresh a token after it has expired
Private Function CreateRefreshRequest() As String
    CreateRefreshRequest = "client_id=" & strClientId
    CreateRefreshRequest = CreateRefreshRequest & "&client_secret=" & strClientSecret
    CreateRefreshRequest = CreateRefreshRequest & "&refresh_token=" & strRefreshToken
    CreateRefreshRequest = CreateRefreshRequest & "&grant_type=refresh_token"
End Function

Private Function CreateLoginRequest()
    'https. //accounts.google.com/o/oauth2/auth?client_id=XXXXXXXXXXXXXXXXXXXXXXXXXXXX&response_type=code&scope=openid%20profile%20email&redirect_uri=urn. ietf. wg. oauth. 2.0. oob&login_hint=myemail@gmail.
    CreateLoginRequest = "https://accounts.google.com/ServiceLogin?"
    CreateLoginRequest = CreateLoginRequest & "client_id=" & strClientId
    CreateLoginRequest = CreateLoginRequest & "&refresh_token=" & strRefreshToken
    CreateLoginRequest = CreateLoginRequest & "&grant_type=refresh_token"
End Function


'//===========================================================================================================
'// Event handling for Internet Explorer Object
'// OAuth 2.0 Process flow requires a user to provide access through the browser for initial Authentication
'//Check the title Window, if Success or Denied Found End the IE interaction
Property Get BrowserComplete() As Boolean
    ' Try to wait until ie is completed processing
    BrowserComplete = WaitUntilLoaded()
    Dim theText As String
    theText = oIExplorer.Document.title
    If InStr(1, theText, "Success") > 0 Then
        strAuthCode = oIExplorer.Document.getElementById("code").Value
        RaiseEvent StatusChanged(Msg("MSG_QUIT_IEXPLORER"))
        oIExplorer.Quit
    ElseIf InStr(1, theText, "Denied") > 0 Then
        oIExplorer.Quit
    Else
        '// Click a button then
        ClickButton oIExplorer, "submit_approve_access"
        BrowserComplete = False
        RaiseEvent StatusChanged(Msg("MSG_GOT_THE_CODE"))
    End If
End Property

'//Only return the AuthCode and then move on the next process..
Public Property Get AuthCode() As String
    GetWebConnector
        
    '// Wait for userInteraction
    SurfAndWait oIExplorer, CreateAuthRequest()
    
    '// Do we have an Authentication Code?
    If Len(strAuthCode) = 0 Then Err.Raise vbObjectError + 2, Description:="User cancelled Authentication"
    
    AuthCode = strAuthCode
End Property

Private Function ElementAvaiable(ieObject As Object, BtnId As String) As Boolean
    ' check for an element is exist or not
    On Error GoTo ErrHandler
    Dim goBtn As Object
    Set goBtn = ieObject.Document.getElementById(BtnId)
    If goBtn Is Nothing Then GoTo ErrHandler
    ElementAvaiable = True
    Set goBtn = Nothing
ErrHandler:
End Function

Private Sub ClickButton(ieObject As Object, BtnId As String)
    ' for clicking a button when required
    Dim goBtn As Object
    Set goBtn = ieObject.Document.getElementById(BtnId)
    goBtn.Click
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not oIExplorer Is Nothing Then
        ' only kill what we created
        oIExplorer.Quit
        Set oIExplorer = Nothing
    End If
    Set Xhr = Nothing
End Sub

Private Sub GetWebConnector(Optional ActiveXmlObjectOnly = False)
    'If Xhr Is Nothing Then Set Xhr = CreateObject("MSXML2.ServerXMLHTTP")
    If Xhr Is Nothing Then Set Xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    If ActiveXmlObjectOnly Then Exit Sub
    Dim str As String
    On Error GoTo ErrHandler
    str = oIExplorer.Document.title
    Exit Sub
ErrHandler:
    Set oIExplorer = CreateObject("InternetExplorer.Application")
    ' just sleep 1 second for stuff to be done
    Sleep 1000
    
    With oIExplorer
        .AddressBar = False
        .MenuBar = False
        .Resizable = False
        .Visible = False
    End With
End Sub
