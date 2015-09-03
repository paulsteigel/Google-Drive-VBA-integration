Option Explicit
Private ScriptEngine As Object
Private jSonObject As Object

Private JsonArr() As String

Public Property Get DecodeJsonString(ByVal JSonString As String) As Object
    Set jSonObject = ScriptEngine.eval("(" + JSonString + ")")
    Set DecodeJsonString = jSonObject
End Property

Public Function GetProperty(ByVal jSonObject As Object, ByVal propertyName As String) 'As Variant
    GetProperty = ScriptEngine.Run("getProperty", jSonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal jSonObject As Object, ByVal propertyName As String) 'As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", jSonObject, propertyName)
End Function

Public Function GetKeys(ByVal jSonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim index As Integer
    Dim key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", jSonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    index = 0
    For Each key In KeysObject
        KeysArray(index) = key
        index = index + 1
    Next
    GetKeys = KeysArray
End Function

Private Sub Class_Initialize()
    Set ScriptEngine = CreateObject("MSScriptControl.ScriptControl")
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
    ScriptEngine.AddCode "function encode(str) {return encodeURIComponent(str);}"
    ReDim JsonArr(0, 0)
End Sub

Private Sub Class_Terminate()
    Set ScriptEngine = Nothing
End Sub

Function EncodeString(theString As String) As String
    EncodeString = ScriptEngine.Run("encode", theString)
End Function

Function Escape(S As String) As String
    Escape = ScriptEngine.eval("escape(" & dq(S) & ")")
End Function

Private Function dq(S As String) As String
    dq = Chr$(34) & S & Chr$(34)
End Function

Property Get JsonArray() As Variant
    'This will create a 2-D array for retrieved values
    Dim ColKey As New Collection, ColVal As New Collection
    If JsonArr(0, 0) <> "" Then GoTo exit_Proc
    ' First call the flatterning to collection
    ParseCollection jSonObject, ColKey, ColVal
    ' Now call the stuff...
    JsonArray = Object2Excel(ColKey, ColVal)
exit_Proc:
    ' now get data to array...
    JsonArr = JsonArr
    Set ColKey = Nothing
    Set ColVal = Nothing
End Property

Private Sub ParseCollection(jsObj As Object, ColKey As Collection, ColVal As Collection, _
    Optional sParentName As String, Optional RemoveItemIndex As Boolean = True)
    ' This will try to recur the search and build a flat collection with keyname and key value
    ' Keyname shall be applied in a treeview path
    Dim keyValue, KeyName As String, Keys() As String, lCtr As Long, KeyPath  As String
        
    Keys = GetKeys(jsObj)
    
    For lCtr = LBound(Keys) To UBound(Keys)
        KeyName = CStr(Keys(lCtr))
        If RemoveItemIndex Then
            If KeyName <> CStr(lCtr) Then
                KeyPath = sParentName & ">" & KeyName
            Else
                KeyPath = sParentName
            End If
        Else
            KeyPath = sParentName & ">" & KeyName
        End If
        
        keyValue = GetProperty(jsObj, KeyName)
    
        If keyValue = "[object Object]" Or InStr(1, keyValue, "[object Object]") > 0 Then
            ' there are stuff behind
            ParseCollection GetObjectProperty(jsObj, KeyName), ColKey, ColVal, KeyPath, RemoveItemIndex
        Else
            ColKey.Add KeyPath
            ColVal.Add keyValue
        End If
    Next
End Sub

Private Function Object2Excel(KeyObj As Collection, ValueObj As Collection) As Variant
    ' This will send 2 object to a 2-D array
    Dim i As Long, j As Long
    Dim HeaderCol As String, KeyColName As String, MaxCol As Long, cCol As Long, cRow As Long, keyValue As String
    Dim StartKeyName As String
    Dim xPos As Long
    Dim myArr() As String
    ' This is the maximum number of colums this appcan handle... 50
    ReDim myArr(50, 100)
    ' We need 2 collection
    j = KeyObj.Count
    For i = 1 To j
        KeyColName = KeyObj.Item(i)
        keyValue = ValueObj.Item(i)
        If StartKeyName = "" Then StartKeyName = KeyColName
        If StartKeyName = KeyColName Then cRow = cRow + 1
        xPos = InStr(HeaderCol, "[" & KeyColName & "]")
        
        If xPos > 0 Then
            'already have this column name
            cCol = Val(Mid(HeaderCol, xPos - 3))
        Else
            'new column name
            HeaderCol = HeaderCol & "," & Format(MaxCol, "000") & "[" & KeyColName & "]"
            ' first colum will keep column name...
            If MaxCol + 1 > 50 Then GoTo exitfunction ' this is critical error...
            myArr(MaxCol, 0) = KeyColName
            cCol = MaxCol
            MaxCol = MaxCol + 1
        End If
        myArr(cCol, cRow) = keyValue
        cCol = cCol + 1
    Next
    ' Now cleaning up this array and make it correctly row,col
    Dim newArr() As String
    'TestArray myArr
    'GoTo ExitFunction
    i = 0
    ReDim newArr(cRow + 1, MaxCol - 1)
    While myArr(i, 0) <> ""
        ' moving through column first
        j = 0
        newArr(j, i) = myArr(i, 0)
        For j = 1 To cRow + 1
            newArr(j, i) = myArr(i, j)
        Next
        i = i + 1
    Wend
    'TestArray newArr
    Object2Excel = newArr
exitfunction:
End Function

Private Sub TestArray(theArray As Variant)
    Dim outStr As String, i As Long, j As Long, keyValue As String
    For i = 0 To UBound(theArray, 1)
        keyValue = ""
        For j = 1 To UBound(theArray, 2)
            keyValue = keyValue + theArray(i, j) + "||"
        Next j
        outStr = outStr + theArray(i, 0) + ":[" + keyValue + "]" + vbCrLf
    Next i
    'WriteLog outStr, "C:\Resutlt.txt", True
End Sub
