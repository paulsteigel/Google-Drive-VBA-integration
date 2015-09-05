Option Compare Database
Option Explicit
Global Const AppTitle = "Demo GoogleDrive VBA integration"
Global Const StrNull = "N/A"
Global Const VbAgent = "Ms Access VBA browser" '"Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)" '"Cig_manager"

' For ini read write
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Property Let AppProperty(KeyName As String, Optional Section As String = "", Optional ConfigFile As String = "", keyValue As String)
    WritePrivateProfileString AppTitle & IIf(Section = "", "", "-" & Section), KeyName, keyValue, IIf(ConfigFile = "", AppConfigPath, ConfigFile)
End Property

Property Get AppProperty(KeyName As String, Optional Section As String = "", Optional ConfigFile As String = "") As String
    Dim tmpBuffer As String * 255, tRet As Long
    tRet = GetPrivateProfileString(AppTitle & IIf(Section = "", "", "-" & Section), KeyName, StrNull, tmpBuffer, Len(tmpBuffer), IIf(ConfigFile = "", AppConfigPath, ConfigFile))
    AppProperty = Left(tmpBuffer, tRet)
End Property

Sub AppStatus(msgText As String)
    SysCmd acSysCmdSetStatus, msgText
End Sub

Property Get AppConfigPath() As String
    AppConfigPath = CurrentProject.path & "\Config.ini"
End Property

Sub OpenLocation(sPath As String)
    ' For browsing specified location
    Dim retVal
    retVal = VBA.Shell("explorer.exe " & sPath, vbNormalFocus)
End Sub

Sub WriteLog(ErrDesc As String, Optional LogFileName As String = "Error.txt", Optional KillIfExist As Boolean = False)
    Dim txtString As String, FileNames As String
    FileNames = LogFileName
    
    txtString = ErrDesc
    Dim UnicodeFile As Boolean
    
    Const ForAppending = 8
    UnicodeFile = True
    
    Dim fso As Object, ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' check if the file exist
    If KillIfExist Then If FileOrDirExists(LogFileName, True) Then Kill LogFileName
    
    Set ts = fso.OpenTextFile(FileNames, ForAppending, True, UnicodeFile)
    ts.WriteLine txtString

    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
End Sub

Function GetAuthCode(Optional client_id As String = "", Optional client_secret As String = "", _
    Optional txtGAccount As String, Optional txtGPassword As String, _
    Optional IsGetNewToken As Boolean = False) As String
    ' Get token for accessing drive object
    Dim myClass As New gAuth2
    If client_id = "" Then
        client_id = AppProperty("App_client_id", "Google Drive")
        client_secret = AppProperty("App_client_secret", "Google Drive")
    Else
        AppProperty("App_client_id", "Google Drive") = client_id
        AppProperty("App_client_secret", "Google Drive") = client_secret
    End If
    With myClass
        .InitClientCredentials client_id, client_secret
        .InitEndPoints
        If IsGetNewToken Then
            .LogOnGoogle txtGAccount, txtGPassword
            If Not .GetNewToken() Then GoTo Exit_Function
        End If
        GetAuthCode = .AuthHeader
    End With
Exit_Function:
    Set myClass = Nothing
End Function

Property Let SetObjInterface(CallObject As Object)
    ' This will set object face language at runtime rather than do this just one
    Dim iObj As New ADODB.Recordset, iCr As Control, Obj As Object, iCaption As String
    Dim i As Long, fLang As String
    fLang = AppLanguage
    ' Initialize interface recordset
    iObj.Open "Select * from tblCaption where ObjectID='" & CallObject.name & "';", CurrentProject.Connection
    With iObj
        ' Set caption for the object
        On Error GoTo ExitMe
        ' Now set caption for all the label in the object
        While Not iObj.EOF
            CallObject.Controls(.Fields("MsgID")).Caption = .Fields("MsgCap" & fLang)
            If .Fields("MsgID") = "FORM_OR_REPORT_NAME" Then CallObject.Caption = .Fields("Msg" + fLang)
            .MoveNext
        Wend
        .Close
        Set Obj = Nothing
    End With
ExitMe:
End Property

Sub GetObjectCaption()
    ' This will get caption of all object and store in tblCaption
    Dim frmObj As Form, SqlStr As String, CtrObj As Control, i As Long

    For i = 0 To CurrentProject.AllForms.Count - 1
        DoCmd.OpenForm CurrentProject.AllForms.Item(i).name, acDesign, , , , acHidden
        Set frmObj = Forms(CurrentProject.AllForms.Item(i).name)
        For Each CtrObj In frmObj.Controls
            If TypeOf CtrObj Is label Or TypeOf CtrObj Is CommandButton Then
                SqlStr = "INSERT INTO tblCaption(ObjectID, MsgGroup, MsgID, MsgCapV) "
                SqlStr = SqlStr + "VALUES('" + frmObj.name + "',1,'" + CtrObj.name + "','" + CtrObj.Caption + "');"
                CurrentDb.Execute SqlStr
            End If
        Next
        ' now for form/report caption
        SqlStr = "INSERT INTO tblCaption(ObjectID, MsgGroup, MsgID, MsgCapV) "
        SqlStr = SqlStr + "VALUES('" + frmObj.name + "',1,'FORM_OR_REPORT_NAME', '" + frmObj.Caption + "')"
        CurrentDb.Execute SqlStr
        DoCmd.Close acForm, frmObj.name, acSaveNo
    Next
ExitMe:
End Sub

Property Get Msg(MessageID As String) As String
    ' This will read the category table for returning a congigured item
    Msg = nz(DLookup("MsgCapV", "tblCaption", "MsgID='" & MessageID & "'"), "Unknown ID or Data not avaiable")
End Property

Property Let DBConfig(PropertyName As String, PropertyValue As String)
    ' Write to property DB
    Dim PrpVal As String
    '1. Check for whether such property exists
    PrpVal = nz(DLookup("MsgCapV", "tblCaption", "MsgID='" & PropertyName & "'"), "")

    If PrpVal <> "" Then
        PrpVal = "UPDATE tblCaption SET MsgCapV='" & PropertyValue & "' WHERE MsgID='" & PropertyName & "';"
    Else
        PrpVal = "INSERT INTO tblCaption(MsgGroup, MsgID, MsgCapV) VALUES(99,'" & PropertyName & "','" & PropertyValue & "');"
    End If
    CurrentDb.Execute PrpVal
End Property

Property Get DBConfig(PropertyName As String) As String
    DBConfig = nz(DLookup("MsgCapV", "tblCaption", "MsgID='" & PropertyName & "'"), "")
End Property

Property Let AppLanguage(NewValue As String)
    ' set default language to English
    If NewValue = "" Then NewValue = "E"
    DBConfig("APP_LANGUAGE") = NewValue
End Property

Property Get AppLanguage() As String
    AppLanguage = DBConfig("APP_LANGUAGE")
End Property
