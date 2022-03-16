Attribute VB_Name = "Clipboard"
Option Explicit
Option Compare Database

#If VBA7 Then
Declare PtrSafe Function abOpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Declare PtrSafe Function abCloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Declare PtrSafe Function abEmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Declare PtrSafe Function abIsClipboardFormatAvailable Lib "user32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
Declare PtrSafe Function abSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare PtrSafe Function abGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Declare PtrSafe Function abGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare PtrSafe Function abGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Declare PtrSafe Function abGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Boolean
Declare PtrSafe Function abLstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare PtrSafe Function abGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Declare PtrSafe Function abGlobalSize Lib "kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long
#Else
Declare Function abOpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Declare Function abCloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Declare Function abEmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Declare Function abIsClipboardFormatAvailable Lib "user32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
Declare Function abSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function abGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Declare Function abGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function abGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Declare Function abGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Boolean
Declare Function abLstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function abGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Declare Function abGlobalSize Lib "kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long
#End If
Const GHND = &H42
Const CF_TEXT = 1
Const APINULL = 0


Function Text2Clipboard(szText As String)
    Dim wLen As Integer
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim RetVal As Variant
    Dim wFreeMemory As Boolean

    ' Get the length, including one extra for a CHR$(0) at the end.
    wLen = Len(szText) + 1
    szText = szText & Chr$(0)
    hMemory = abGlobalAlloc(GHND, wLen + 1)
    If hMemory = APINULL Then
        MsgBox "Unable to allocate memory."
        Exit Function
    End If
    wFreeMemory = True
    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock memory."
        GoTo T2CB_Free
    End If

    ' Copy our string into the locked memory.
    RetVal = abLstrcpy(lpMemory, szText)
    ' Don't send clipboard locked memory.
    RetVal = abGlobalUnlock(hMemory)

    If abOpenClipboard(0&) = APINULL Then
        MsgBox "Unable to open Clipboard.  Perhaps some other application is using it."
        GoTo T2CB_Free
    End If
    If abEmptyClipboard() = APINULL Then
        MsgBox "Unable to empty the clipboard."
        GoTo T2CB_Close
    End If
    If abSetClipboardData(CF_TEXT, hMemory) = APINULL Then
        MsgBox "Unable to set the clipboard data."
        GoTo T2CB_Close
    End If
    wFreeMemory = False

T2CB_Close:
    If abCloseClipboard() = APINULL Then
        MsgBox "Unable to close the Clipboard."
    End If
    If wFreeMemory Then GoTo T2CB_Free
    Exit Function

T2CB_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global memory."
    End If
End Function

Function Clipboard2Text()
    Dim wLen As Integer
    Dim hMemory As Long
    Dim hMyMemory As Long

    Dim lpMemory As Long
    Dim lpMyMemory As Long

    Dim RetVal As Variant
    Dim wFreeMemory As Boolean
    Dim wClipAvail As Integer
    Dim szText As String
    Dim wSize As Long

    If abIsClipboardFormatAvailable(CF_TEXT) = APINULL Then
        Clipboard2Text = Null
        Exit Function
    End If

    If abOpenClipboard(0&) = APINULL Then
        MsgBox "Unable to open Clipboard.  Perhaps some other application is using it."
        GoTo CB2T_Free
    End If

    hMemory = abGetClipboardData(CF_TEXT)
    If hMemory = APINULL Then
        MsgBox "Unable to retrieve text from the Clipboard."
        Exit Function
    End If
    wSize = abGlobalSize(hMemory)
    szText = space(wSize)

    wFreeMemory = True

    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock clipboard memory."
        GoTo CB2T_Free
    End If

    ' Copy our string into the locked memory.
    RetVal = abLstrcpy(szText, lpMemory)
    ' Get rid of trailing stuff.
    szText = Trim(szText)
    ' Get rid of trailing 0.
    Clipboard2Text = Left(szText, Len(szText) - 1)
    wFreeMemory = False

CB2T_Close:
    If abCloseClipboard() = APINULL Then
        MsgBox "Unable to close the Clipboard."
    End If
    If wFreeMemory Then GoTo CB2T_Free
    Exit Function

CB2T_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global clipboard memory."
    End If
End Function
