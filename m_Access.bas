Attribute VB_Name = "m_Access"
Option Compare Database
Option Explicit

'--- DROP
#If VBA7 Then
Public Declare PtrSafe Function apiCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare PtrSafe Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub sapiDragAcceptFiles Lib "shell32.dll" Alias "DragAcceptFiles" (ByVal hwnd As Long, ByVal fAccept As Long)
Public Declare PtrSafe Sub sapiDragFinish Lib "shell32.dll" Alias "DragFinish" (ByVal hDrop As Long)
Public Declare PtrSafe Function apiDragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
#Else
Public Declare Function apiCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub sapiDragAcceptFiles Lib "shell32.dll" Alias "DragAcceptFiles" (ByVal hwnd As Long, ByVal fAccept As Long)
Public Declare Sub sapiDragFinish Lib "shell32.dll" Alias "DragFinish" (ByVal hDrop As Long)
Public Declare Function apiDragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
#End If

Public Drop_lpPrevWndProc As Long
Public Drop_Text As String
Public Const GWL_WNDPROC   As Long = (-4)
Public Const GWL_EXSTYLE = (-20)
Public Const WM_DROPFILES = &H233
Public Const WM_MOUSEWHELL = &H20A
Public Const WS_EX_ACCEPTFILES = &H10&
'----

'=========== Access - funções referentes a Banco de dados e campos ===========

Public Sub AppTitle(t As String)
  Dim obj As Object
  Const conPropNotFoundError = 3270
  On Error GoTo ErrorHandler
  CurrentDb.Properties!AppTitle = t
  Application.RefreshTitleBar
  Exit Sub
ErrorHandler:
  If Err.Number = conPropNotFoundError Then
    Set obj = CurrentDb.CreateProperty("AppTitle", dbText, t)
    CurrentDb.Properties.Append obj
  Else
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
  End If
  Resume Next
End Sub


Public Sub Remove_Prefix_Tables(Prefix As String, Optional New_Prefix = "")
  ' Remove table names prefix (to use in linked tables)
  Dim t%, x
  t = Len(Prefix)
  If t = 0 Then Exit Sub
  For Each x In CurrentDb.TableDefs
    If Left(x.Name, t) = Prefix Then
      x.Name = New_Prefix + Mid(x.Name, t + 1)
    End If
  Next
End Sub


Public Function fn_Scalar(SQL As String, Optional DSN As String = "", Optional l = dbSeeChanges)
  On Error GoTo fn_Scalar_Err
  Dim q As QueryDef
  Dim a As dao.Recordset
  Set q = CurrentDb.CreateQueryDef("")
  If DSN > "" Then q.Connect = "ODBC;DSN=" & DSN & ";"
  q.SQL = SQL
  Set a = q.OpenRecordset(dbOpenDynaset, l)
  If a.RecordCount = 0 Then
    fn_Scalar = Null
  Else
    fn_Scalar = a.Fields(0).value
  End If
  a.Close
  Exit Function
fn_Scalar_Err:
  If Errors.Count >= 2 Then
    fn_Scalar = "**** " & error & " = " & Errors(Errors.Count - 2).Description
  ElseIf Errors.Count >= 1 Then
    fn_Scalar = "**** " & error & " = " & Errors(Errors.Count - 1).Description
  End If
End Function


Public Function fn_Count(Table As String, Field As String, value As Variant) As Integer
  fn_Count = fn_Scalar("Select count(*) from " & Table & " where " & Field & "=" & SQL_Value(value))
End Function


Public Function fn_Table(SQL As String, Optional DSN As String = "") As dao.Recordset
  Dim q As QueryDef
  Dim r As dao.Recordset
  Set q = CurrentDb.CreateQueryDef("")
  If DSN > "" Then q.Connect = "ODBC;DSN=" & DSN & ";Trusted_Connection=Yes;"
  q.SQL = SQL
  Set fn_Table = q.OpenRecordset(dbOpenDynaset, dbSeeChanges)
End Function


Public Function Exec_SQL(SQL As String, Optional DSN As String = "", Optional timeout As Integer = -1) As Integer
  Dim q As QueryDef
  Set q = CurrentDb.CreateQueryDef("")
  If DSN > "" Then q.Connect = "ODBC;DSN=" & DSN & ";Trusted_Connection=Yes;"
  q.SQL = SQL
  q.ReturnsRecords = False
  If timeout > 0 Then q.ODBCTimeout = timeout
'  q.Execute dbSeeChanges
  On Error GoTo erroexecsql
   q.Execute
  ' ? dao.Errors(0)
  Exec_SQL = q.RecordsAffected
  Exit Function
erroexecsql:
  Dim r As VbMsgBoxResult
  r = MsgBox(Errors_Msg(), vbAbortRetryIgnore)
  If r = vbRetry Then Resume
  If r = vbIgnore Then Resume Next
  Error 9999
End Function


Public Function GetString(r As Recordset) As String
  Dim s As String, x As Integer
  If r.RecordCount > 0 Then
    r.MoveFirst
    While Not r.EOF()
      If Len(s) > 0 Then s = s & vbCrLf
      For x = 0 To r.Fields.Count - 1
        s = s & r.Fields(x) & ";"
      Next x
      r.MoveNext
    Wend
  End If
  GetString = s
End Function


Public Function GUID_Clean(GUID As Variant) As String
  Dim G As String
  If IsNull(GUID) Then
    G = ""
  ElseIf TypeName(GUID) = "Textbox" Or TypeName(GUID) = "ComboBox" Then
    G = StringFromGUID(Nz(GUID, ""))
  Else
    G = GUID
  End If
  If Left(G, 6) = "{guid " Then
    GUID_Clean = Mid(G, 7, 38)
  Else
    GUID_Clean = G
  End If
End Function


Public Function Form_GUID(Form As String, Field As String) As String
  On Error Resume Next
  Form_GUID = GUID_Clean(StringFromGUID(Forms(Form).Form.Recordset(Field)))
End Function


Public Function Next_Code(Table As String, Optional Field As String = "")
  ' Return the next code of the PrimaryKey of table
  Dim a As Recordset, b
  Set a = CurrentDb.OpenRecordset(Table)
  a.Index = "PrimaryKey"
  a.MoveLast
  If Field = "" Then
    Next_Code = a(0) + 1
  Else
    Next_Code = a(Field) + 1
  End If
  a.Close
End Function


Public Function Locate(f As Form, ByVal Code As Variant, Optional lB As ListBox = Nothing)
  ' Localiza o código cod no primeiro campo do recordset do form f.
  ' Caso LB seja passado, posiciona LB em cod
  Dim rs As dao.Recordset
  Dim tipo As String
  tipo = TypeName(Code)
  If tipo = "AccessField" Or tipo = "Field" Then tipo = TypeName(Code.value)
  If tipo <> "GUID" Then
    If Not Left(Code, 1) = "{" Then Code = Str(Nz(Code, 0))
  Else
    Code = GUID_Clean(Code)
  End If
  Set rs = f.RecordsetClone
  If rs.Fields(0).Type = dbGUID Then
    rs.MoveFirst
    Do While Not rs.EOF()
      If Code = GUID_Clean(rs.Fields(0)) Then Exit Do
      rs.MoveNext
    Loop
  Else
    If Not IsNull(Code) Then
      rs.FindFirst "[" & rs.Fields(0).Name & "] = " & Code  ' Str(Nz(lst, 0))
    End If
  End If
  If Not rs.EOF Then f.Bookmark = rs.Bookmark
  If Not lB Is Nothing Then lB.value = Code
End Function



'---- Trava para edição de formulários
' Inserir botão P_Editar no form
' Acertar eventos do Form conforme abaixo:
'
'Private Sub P_Editar_Click()
'  P_Editar.Enabled = Form_Lock(Me, False)
'End Sub
'
'Private Sub Form_AfterUpdate()
'  P_Editar.Enabled = Form_Lock(Me)
'End Sub
'
'Private Sub Form_Current()
'  P_Editar.Enabled = Form_Lock(Me)
'End Sub
'
'Private Sub Form_BeforeUpdate(Cancel As Integer)
'  If Msg_Save_Changes(Cancel) = vbNo Then P_Editar.Enabled = Form_Lock(Me)
'End Sub
'
'Private Sub Form_BeforeInsert(Cancel As Integer)
'  If Msg_Save_Changes(Cancel) = vbNo Then P_Editar.Enabled = Form_Lock(Me)
'End Sub
'
'Private Sub Form_Undo(Cancel As Integer)
'  P_Editar.Enabled = Form_Lock(Me)
'End Sub
'

Public Sub Save_Record()
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
End Sub

Public Sub Undo()
    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
End Sub

Public Function Msg_Save_Changes(ByRef Cancel As Integer) As VbMsgBoxResult
  Dim x As VbMsgBoxResult
  x = MsgBox("Salva as alterações?", vbYesNoCancel + vbDefaultButton1) ' "Save changes?"
  If x = vbYes Then Exit Function
  If x = vbNo Then Undo
  Cancel = (x = vbCancel)
  Msg_Save_Changes = x
End Function

Function Form_Lock(f As Form, Optional trava As Boolean = True) As Boolean
  ' trava todos os objetos do form que não possuem a tag=nb ou backcolor=14674687 para edição
  Dim x, ti As Control, mintab As Integer, i As Long
  mintab = 9999
  For Each x In f.Controls
    If x.Tag <> "nb" Then
On Error Resume Next
      i = 0
      i = x.BackColor
      If i <> 14674687 Then
        x.Locked = trava
        If trava = False Then
          i = 9999
          i = x.TabIndex    ' se não tiver a propriedade, continua com 9999
          If TypeName(ti) <> "Label" And i < mintab And x.Visible = True Then
            Set ti = x
            mintab = i
          End If
        End If
      End If
On Error GoTo 0
    End If
  Next
  If trava = False Then ti.SetFocus
  Form_Lock = trava
End Function



'--------------------

Public Function ListBox_MultValue(l As ListBox, Optional Column = 0, Optional Delimiter = ",", Optional Type_ = "N") As String
' Retorna string com todos os itens selecionados separados por Separador
  Dim r$, x%
  For x = 0 To l.ListCount
    If l.Selected(x) Then r = r & IIf(Type_ <> "N", """", "") & l.Column(Column, x) & IIf(Type_ <> "N", """", "") & Delimiter
  Next x
  If Right(r, Len(Delimiter)) = Delimiter Then r = Left(r, Len(r) - Len(Delimiter))
  ListBox_MultValue = r
End Function


Public Function Make_WHERE(Field As String, FieldDefault As String, lst As ListBox) As String
  Dim f$, campo$
  f = "("
  campo = IIf(Field = "*" Or Field = ".", FieldDefault, Replace(Field, "*", ""))
  If Left(Field, 1) = "*" Then f = f & campo & " IS NULL OR "
  Make_WHERE = f & campo & " In (" & ListBox_MultValue(lst) & "))"
End Function


Public Function Insert_Name(Table As String, Name As String, Optional Field_Code = 0, Optional Field_Name = 1)
  ' Field_Code = number or name of the field that have the PrimaryKey
  ' Field_Name = number or name of the field that will save the Name value
  On Error GoTo IncluiNome_erro
  Dim a As Recordset, b
  If Nz(Name) = "" Then
    Insert_Name = acDataErrDisplay
  ElseIf MsgBox("Inclui o " + Table + " '" + Name + "' ?", vbYesNo, "Inclusão") = vbYes Then
    Set a = CurrentDb.OpenRecordset(Table)
    a.Index = "PrimaryKey"
    a.MoveLast
    b = a(Field_Code) + 1
    a.AddNew
    a(Field_Name) = Name
    a(Field_Code) = b
    a.Update
    a.Close
    Insert_Name = acDataErrAdded
  Else
    Insert_Name = acDataErrDisplay
  End If
  Exit Function
IncluiNome_erro:
  MsgBox Err.Description
  Insert_Name = acDataErrContinue
End Function


Public Function GetFixedSizeTXT(r As Recordset, Optional So_Estrutura As Boolean = False) As String
  Dim s As String, x As Integer, e As Boolean
  If So_Estrutura Then
    e = True
  Else
    r.MoveFirst
  End If
  While (So_Estrutura = False And Not r.EOF()) Or (So_Estrutura And e)
    For x = 0 To r.Fields.Count - 1
      If r.Fields(x).Type = dbDecimal Then
        If e Then
          ' ? r.Fields(x).Properties("decimalPlaces")
          s = s & r.Fields(x).Name & vbTab & "N" & r.Fields(x).CollatingOrder & vbCrLf  ' bug - size está no collatingOrder
        Else
          s = s & Format(r.Fields(x), String(r.Fields(x).CollatingOrder, "0"))
        End If
      ElseIf r.Fields(x).Type = dbText Then
        If e Then
          s = s & r.Fields(x).Name & vbTab & "A" & r.Fields(x).Size & vbCrLf
        Else
          s = s & RPad(Nz(r.Fields(x)), r.Fields(x).Size, " ")
        End If
      Else
        MsgBox "Tipo de campo inválido."
      End If
    Next x
    If So_Estrutura Then
      e = False
    Else
      r.MoveNext
      s = s & vbCrLf
    End If
  Wend
  GetFixedSizeTXT = s
End Function


Public Sub Replace_Table_Connections(Text_to_Find As String, New_Text As String)
  ' Replace text in 'Connect' string of linked tables
  Dim t%, x
  t = Len(Text_to_Find)
  If t = 0 Then Exit Sub
  For Each x In CurrentDb.TableDefs
    If InStr(x.Connect, Text_to_Find) > 0 Then
      x.Connect = Replace(x.Connect, Text_to_Find, New_Text)
      x.RefreshLink
    End If
  Next
End Sub


Public Function ADORecordsetMemory(query As String) As ADODB.Recordset
  Dim rs As ADODB.Recordset
  Dim r As ADODB.Recordset
  Dim f, x%, nc%
  Set rs = New ADODB.Recordset
  Set r = New ADODB.Recordset
  rs.Open query, CurrentProject.Connection
  With r
    nc = rs.Fields.Count - 1
    For x = 0 To nc
      r.Fields.Append rs.Fields(x).Name, rs.Fields(x).Type, rs.Fields(x).DefinedSize, rs.Fields(x).Attributes
    Next x
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .LockType = adLockPessimistic
    .Open
  End With
  rs.MoveFirst
  Do Until rs.EOF
    r.AddNew
    For x = 0 To nc
      r.Fields(x) = rs.Fields(x)
    Next x
    rs.MoveNext
  Loop
  Set ADORecordsetMemory = r
End Function


Public Function AgregaStr(tabela As String, Campo_Agr As String, Campo_Where As String, Valor, Optional Separador As String = ",", Optional OrderBy As String = "")
  Dim s As String
  If IsNull(Valor) Then Exit Function
  s = RTrimEx(GetString(CurrentDb.OpenRecordset("SELECT " & Campo_Agr & " FROM " & tabela & " WHERE " & Campo_Where & "=" & Valor & IIf(OrderBy > "", " ORDER BY " & OrderBy, ""))), ";")
  AgregaStr = RTrimEx(Replace(s, ";" & vbCrLf, Separador), Separador)
End Function


Public Function Picture_In_Report(Img As Image, File As String)
  Img.Picture = File
End Function


'----- Drop files

Sub HookFrm(frm As Form, Optional Drop As Boolean = False)
  If Drop_lpPrevWndProc > 0 Then
    MsgBox "Já tem form Hooked"
  Else
    If Drop Then
      Call apiSetWindowLong(frm.hwnd, GWL_EXSTYLE, apiGetWindowLong(frm.hwnd, GWL_EXSTYLE) Or WS_EX_ACCEPTFILES)
      Call sapiDragAcceptFiles(frm.hwnd, True)
    End If
    Drop_lpPrevWndProc = apiSetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf sDragDrop)
    If Drop_lpPrevWndProc = 0 Then
      MsgBox "Erro"
    End If
  End If
End Sub


Sub UnHookFrm(frm As Form)
  Call apiSetWindowLong(frm.hwnd, GWL_WNDPROC, Drop_lpPrevWndProc)
  Drop_lpPrevWndProc = 0
End Sub


Sub LeFilesDrop(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim lngRet As Long, strTmp As String, intLen As Integer
  Dim lngCount As Long, i As Long, strOut As String
  Const cMAX_SIZE = 500
    strTmp = String$(cMAX_SIZE, 0)
    lngCount = apiDragQueryFile(wParam, &HFFFFFFFF, strTmp, cMAX_SIZE)
    For i = 0 To lngCount - 1
      strTmp = String$(cMAX_SIZE, 0)
      intLen = apiDragQueryFile(wParam, i, strTmp, cMAX_SIZE)
      strOut = strOut & Left$(strTmp, intLen) & ";"
    Next i
    strOut = Left$(strOut, Len(strOut) - 1)
    Call sapiDragFinish(wParam)
    Drop_Text = strOut
    Call apiCallWindowProc(Drop_lpPrevWndProc, hwnd, Msg, wParam, lParam)
''    With lstDrop
''      .RowSourceType = "Value List"
''      .RowSource = strOut
''      Caption = "DragDrop: " & .ListCount & " files dropped."
'    End With
End Sub


Sub sDragDrop(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  On Error Resume Next
  If Msg = WM_MOUSEWHELL Then
  ElseIf Msg = WM_DROPFILES Then
'    LeFilesDrop
  Else
    Call apiCallWindowProc(Drop_lpPrevWndProc, hwnd, Msg, wParam, lParam)
  End If
End Sub


Sub RecordSet_to_Excel(tabela As String, FileName As String)
  DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, tabela, FileName, True
End Sub


Sub Remove_Field(tabela As String, campo As String)
  Dim x
  Set x = CurrentDb.TableDefs(tabela).Fields(campo)
  CurrentDb.TableDefs(tabela).Fields.Delete x.Name
End Sub


Function Errors_Msg() As String
  Dim r, l$
  For Each r In dao.Errors
    l = l & r.Description & vbCrLf
  Next
  Errors_Msg = l
End Function


