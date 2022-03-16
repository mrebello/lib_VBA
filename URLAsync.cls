VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "URLAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'* Tools->References
'WinHttp        Microsoft WinHTTP Services, version 5.1          C:\WINDOWS\system32\winhttpcom.dll

Private WithEvents moWHR As WinHttp.WinHttpRequest
Attribute moWHR.VB_VarHelpID = -1

Public msBufferedResponse As String
Public mbFinished As Boolean

Private Const mbDEFAULT_DEBUG  As Boolean = True
Public mvDebug As Variant

Public Property Get bDebug() As Boolean
    If IsEmpty(mvDebug) Then mvDebug = mbDEFAULT_DEBUG
    bDebug = mvDebug
End Property

Public Property Let bDebug(ByVal bRHS As Boolean)
    mvDebug = bRHS
End Property

Public Sub HttpGet(ByVal sURL As String, bAsync As Boolean)
    On Error GoTo ErrHandler
    Set moWHR = New WinHttp.WinHttpRequest
    
    mbFinished = False
    msBufferedResponse = ""
    
    moWHR.Open Method:="GET", url:=sURL, async:=bAsync
    
    moWHR.Send
    Debug.Print "send called with bAsync=" & bAsync
SingleExit:
    Exit Sub
ErrHandler:
    Debug.Print "Error (" & Err.Number & ") " & Err.Description
    Stop
    Resume
    
End Sub


Private Sub moWHR_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Debug.Print "moWHR_OnError"
End Sub

Private Sub moWHR_OnResponseDataAvailable(Data() As Byte)
    Dim sThisChunk As String
    sThisChunk = StrConv(Data(), vbUnicode)
    
    Debug.Print "moWHR_OnResponseDataAvailable (" & Len(sThisChunk) & ")"
    
    msBufferedResponse = msBufferedResponse & sThisChunk
End Sub

Private Sub moWHR_OnResponseFinished()
    Debug.Print "moWHR_OnResponseFinished"
    mbFinished = True
    
    If msBufferedResponse > "" Then
      On Error Resume Next
      If CurrentProject.AllForms("frm_PF_Reusmo").IsLoaded Then
        Form_frm_PF_Resumo.SetFocus
        Form_frm_PF_Resumo.Undo
        Form_frm_PF_Resumo.Requery
        Form_frm_PF_Resumo.Lista123.Selected(0) = True
      End If
    End If

    ' reload HttpGet "xxxxx", True
End Sub

Private Sub moWHR_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    Dim v
    v = VBA.Split(moWHR.getAllResponseHeaders, vbNewLine)
    Debug.Print "moWHR_OnResponseStart"
End Sub
