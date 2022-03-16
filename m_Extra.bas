Attribute VB_Name = "m_Extra"
Option Explicit
Option Compare Database

#If VBA7 Then
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function StopMouseWheel Lib "MouseHook" (ByVal hwnd As Long, ByVal AccessThreadID As Long, Optional ByVal bNoSubformScroll As Boolean = False, Optional ByVal blIsGlobal As Boolean = False) As Boolean
Private Declare PtrSafe Function StartMouseWheel Lib "MouseHook" (ByVal hwnd As Long) As Boolean
#Else
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function StopMouseWheel Lib "MouseHook" (ByVal hwnd As Long, ByVal AccessThreadID As Long, Optional ByVal bNoSubformScroll As Boolean = False, Optional ByVal blIsGlobal As Boolean = False) As Boolean
Private Declare Function StartMouseWheel Lib "MouseHook" (ByVal hwnd As Long) As Boolean
#End If
Private hLib As Long

Public Function MouseWheelON() As Boolean
  MouseWheelON = StartMouseWheel(Application.hWndAccessApp)
  If hLib <> 0 Then
    hLib = FreeLibrary(hLib)
  End If
End Function


Public Function MouseWheelOFF(Optional NoSubFormScroll As Boolean = False, Optional GlobalHook As Boolean = False) As Boolean
  Dim s As String
  Dim blret As Boolean
  Dim AccessThreadID As Long
  
  On Error Resume Next
  ' Our error string
  s = "Sorry...cannot find the MouseHook.dll file" & vbCrLf
  s = s & "Please copy the MouseHook.dll file to your Windows System folder or into the same folder as this Access MDB."
  
  ' OK Try to load the DLL assuming it is in the Window System folder
  hLib = LoadLibrary("MouseHook.dll")
  If hLib = 0 Then
    ' See if the DLL is in the same folder as this MDB
    ' CurrentDB works with both A97 and A2K or higher
    hLib = LoadLibrary(CurrentDBDir() & "utils\MouseHook.dll")
    If hLib = 0 Then
      MsgBox s, vbOKOnly, "MISSING MOUSEHOOK.dll FILE"
      MouseWheelOFF = False
      Exit Function
    End If
  End If
  
  ' Get the ID for this thread
  AccessThreadID = GetCurrentThreadId()
  ' Call our MouseHook function in the MouseHook dll.
  ' Please not the Optional GlobalHook BOOLEAN parameter
  ' Several developers asked for the MouseHook to be able to work with
  ' multiple instances of Access. In order to accomodate this request I
  ' have modified the function to allow the caller to
  ' specify a thread specific(this current instance of Access only) or
  ' a global(all applications) MouseWheel Hook.
  ' Only use the GlobalHook if you will be running multiple instances of Access!
  MouseWheelOFF = StopMouseWheel(Application.hWndAccessApp, AccessThreadID, NoSubFormScroll, GlobalHook)
End Function

Public Function Twofish_Encode(Text As String, key As String) As String
  If Text = "" Or key = "" Then Exit Function
  Dim T As Twofish
  Set T = New Twofish
  Twofish_Encode = T.EncryptString(Text, key, True)
End Function


Public Function Twofish_Decode(Text_base64 As String, key As String) As String
  If Text_base64 = "" Or key = "" Then Exit Function
  Dim T As Twofish
  Set T = New Twofish
  Twofish_Decode = T.DecryptString(Text_base64, key, True)
End Function


Function IndexText(s As String)
  Dim r As String
  r = UCase(Remove_Accents(s))
  r = Replace(r, " DE ", " ")
  r = Replace(r, " DA ", " ")
  r = Replace(r, " DAS ", " ")
  r = Replace(r, " DO ", " ")
  r = Replace(r, " DOS ", " ")
  r = Replace(r, "CE", "SE")
  r = Replace(r, "CI", "SI")
  r = Replace(r, "S ", " ")
  r = Replace(r, "RR", "R")
  r = Replace(r, "SS", "S")
  r = Replace(r, "LL", "L")
  r = Replace(r, "QU", "K")
  r = Replace(r, "GN", "N")
  r = Subst_Car(r, "WCNZYH-.+\ ", "UKMSI,,,,,,")
  r = Replace(r, ",", "")
  IndexText = r
End Function



Public Function CurrentDBDir() As String
  Dim strDBPath As String, P As Long
  strDBPath = CurrentDb.Name
  P = InStrRev(strDBPath, "\")
  CurrentDBDir = Left$(strDBPath, P)
End Function



Public Function Primeira_Maiuscula(texto As String) As String
  Dim x, r$
  Static Excecao$
  'If Excecao = "" Then Excecao = Le_Configuracao("Primeira_Maiuscula_Excecao")
  For Each x In Split(texto, " ")
    If x > "" Then
      If InStr(Excecao, x & ",") > 0 Then
        r = r & LCase(x) & " "
      Else
        r = r & UCase(Left(x, 1)) & LCase(Mid(x, 2)) & " "
      End If
    End If
  Next
  Primeira_Maiuscula = RTrim(r)
End Function


