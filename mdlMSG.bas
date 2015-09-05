Option Compare Database
Option Explicit
'=======================================================
' Module for Vietnamized MsgBox function
' This overide default VBA MsgBox function with some small
' modifications of text, button caption....
' Use this MsgBox function like it is in default VBA IDE
'=======================================================
' Import
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr
     
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" _
        Alias "SetDlgItemTextW" _
        (ByVal hDlg As LongPtr, _
         ByVal nIDDlgItem As LongPtr, _
         ByVal lpString As String) As LongPtr
     
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
        Alias "SetWindowsHookExA" _
        (ByVal idHook As LongPtr, _
         ByVal lpfn As LongPtr, _
         ByVal hmod As LongPtr, _
         ByVal dwThreadID As LongPtr) As LongPtr
     
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As LongPtr) As LongPtr
     
    ' Handle to the Hook procedure
    Private hHook As LongPtr
#Else
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
     
    Private Declare Function SetDlgItemText Lib "user32" _
        Alias "SetDlgItemTextW" _
        (ByVal hDlg As Long, _
         ByVal nIDDlgItem As Long, _
         ByVal lpString As String) As Long
     
    Private Declare Function SetWindowsHookEx Lib "user32" _
        Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, _
         ByVal lpfn As Long, _
         ByVal hmod As Long, _
         ByVal dwThreadID As Long) As Long
     
    Private Declare Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As Long) As Long
     
    ' Handle to the Hook procedure
    Private hHook As Long
#End If
' Hook type
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
 
' Constants
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

' Modify this code for English
Private StrYes As String
Private StrNo As String
Private StrOK As String
Private StrCancel As String
Private StrRetry As String
Private StrIgnore As String
Private StrAbort As String

Private Enum MsoAlertCancelType
    msoAlertCancelDefault = &HFFFFFFFF
    msoAlertCancelFifth = 4
    msoAlertCancelFirst = 0
    msoAlertCancelFourth = 3
    msoAlertCancelSecond = 1
    msoAlertCancelThird = 2
End Enum

' Application title
Private App_Title As String

Function MsgBox(MessageTxt As String, Optional msgStyle As VbMsgBoxStyle, Optional DlgCaption As String = "") As VbMsgBoxResult
    Beep
    If App_Title = "" Then App_Title = "Demo app"
    Dim msgBoxIcon As Long, msgButton As Long, btnStyle As Long, ErrLoop As Boolean
    Dim ButtonDefault As Long
    
    ' Determine what button is default....
    Dim btnArr As Variant, i As Long
    btnArr = Array(0, 256, 512, 768)
    For i = 0 To UBound(btnArr)
        btnStyle = msgStyle - btnArr(i)
        If btnStyle < 0 Then
            ButtonDefault = i - 1
            btnStyle = msgStyle - btnArr(i - 1)
            ErrLoop = True
            Exit For
        End If
    Next
    
    ' Determine Icon...
    btnArr = Array(0, 16, 32, 48, 64)
    For i = 0 To UBound(btnArr)
        msgButton = btnStyle - btnArr(i)
        If msgButton <= 0 Then
            If msgButton = 0 Then
                msgBoxIcon = i
                btnStyle = btnStyle - btnArr(i)
            Else
                msgBoxIcon = i - 1
                btnStyle = btnStyle - btnArr(i - 1)
            End If
            ErrLoop = True
            Exit For
        End If
    Next
    If ErrLoop Then
        ' get the button style
        If msgButton < 0 Then msgButton = btnStyle
        
        ' clear error if number of button is smaller than the default setting...
        If ButtonDefault > msgButton Then ButtonDefault = msgButton
    Else
        ButtonDefault = 0
        msgButton = 0
        msgBoxIcon = 0
    End If
    ' Set Hook
    hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, 0, GetCurrentThreadId)
    ' Display the messagebox
    MsgBox = Application.Assistant.DoAlert(IIf(DlgCaption <> "", DlgCaption, App_Title), _
        MessageTxt, msgButton, msgBoxIcon, ButtonDefault, msoAlertCancelDefault, True)
End Function
 
Private Function MsgBoxHookProc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If lMsg = HCBT_ACTIVATE Then
        StrYes = "&C" & ChrW(243)
        StrNo = "&Kh" & ChrW(244) & "ng"
        StrOK = "Ch" & ChrW(7845) & "p nh" & ChrW(7853) & "&n"
        StrCancel = "&H" & ChrW(7911) & "y"
        StrRetry = "&Th" & ChrW(7917) & " l" & ChrW(7841) & "i"
        StrAbort = "&D" & ChrW(7915) & "ng"
        StrIgnore = "&B" & ChrW(7887) & " qua"
  
        SetDlgItemText wParam, IDYES, StrConv(StrYes, vbUnicode)
        SetDlgItemText wParam, IDNO, StrConv(StrNo, vbUnicode)
        SetDlgItemText wParam, IDCANCEL, StrConv(StrCancel, vbUnicode)
        SetDlgItemText wParam, IDOK, StrConv(StrOK, vbUnicode)
        SetDlgItemText wParam, IDABORT, StrConv(StrAbort, vbUnicode)
        SetDlgItemText wParam, IDRETRY, StrConv(StrRetry, vbUnicode)
        SetDlgItemText wParam, IDIGNORE, StrConv(StrIgnore, vbUnicode)
        ' Release the Hook
        UnhookWindowsHookEx hHook
    End If
    MsgBoxHookProc = False
End Function
