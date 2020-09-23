Attribute VB_Name = "Conexion_2"
Public Type INFO_Usuario
    Codigo_Usuario As String
    Nombre_Usuario As String
    Usuario_Es_Admin As String
    Usuario_Passwd As String
    Usuario_LastLogin As String
    Usuario_LastTime As String
End Type

Public Type Empresa
 cID As Long
 cEmpresa As String
 cRepresentante As String
 cCc As String
 cClase As String
 cDireccion As String
 cCiudad As String
End Type
Public Datos_Empresa As Empresa
Public Current_User As INFO_Usuario


Public Function ConnectDB(ByRef vDB As ADODB.Connection, PathFileName As String) As Boolean

On Error GoTo errh
 
    If vDB.State = adStateOpen Then vDB.Close
        
    vDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=jepirr4uniguajira"
    
    ConnectDB = True
    
    Exit Function
    
errh:
    ANS = MsgBox("Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical + vbRetryCancel)
    If ANS = vbCancel Then
        ConnectDB = False
        End
    ElseIf ANS = vbRetry Then
        Resume
    End If
'    WriteErrorLog "modDBMain", "ConnectDB", Err.Description
    'ConnectDB = False
    
End Function

Public Function CloseDB(ByRef vDB As ADODB.Connection)
    vDB.Close
End Function


Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRs As ADODB.Recordset, ssql As String, Optional sHowMSG As Boolean = True, Optional ByRef iErrNumber As Variant, Optional ByRef sErrDescription As Variant) As Boolean
    
On Error GoTo errh

    
    Set vRs = Nothing
    Set vRs = New ADODB.Recordset
  
  
    vRs.Open ssql, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True

    
    Exit Function
    
'-------------------------------------------
errh:
    If sHowMSG = True Then
        MsgBox "modDBMain" & "," & "ConnectRS" & "," & "Unable to connect Recordset / Err: " & err.Description, vbExclamation
    End If
    If Not IsMissing(iErrNumber) Then
        iErrNumber = err.Number
    End If
    If Not IsMissing(sErrDescription) Then
        sErrDescription = err.Description
    End If
    ConnectRS = False
End Function


Public Function Esta_Vacio(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        Esta_Vacio = True
        If UseTagValue = True Then
            MsgBox "El Campo '" & sText.Tag & "' Es requerido!", vbExclamation
        Else
            MsgBox "El campo es requerido.", vbExclamation
        End If
        sText.SetFocus
    Else
        Esta_Vacio = False
    End If
End Function

Public Sub pageFillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal pos_start As Long, ByVal pos_end As Long, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional match_field As String, Optional match_str As String, Optional match_ico As Byte, Optional srcHiddenField As String)

    Dim x As ListItem
    Dim i As Byte, c As Long, old_pt As Long
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.AbsolutePosition = pos_start
    On Error Resume Next
    old_pt = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    DoEvents
    Do
        If match_field = "" Then
            If with_num = True Then
                Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
            Else
                Set x = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), sNumIco, sNumIco)
            End If
        Else
            If sRecordSource.Fields(match_field) = match_str Then
                If with_num = True Then
                    Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, match_ico, match_ico)
                Else
                    Set x = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), match_ico, match_ico)
                End If
            Else
                If with_num = True Then
                    Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
                Else
                    Set x = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), sNumIco, sNumIco)
                End If
            End If
        End If
            If srcHiddenField <> "" Then
                x.Tag = sRecordSource.Fields(srcHiddenField) & "*~~~~~*" & c + pos_start
              Else
                x.Tag = c + pos_start
            End If
            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                             x.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) - 1))
                    Else
                            x.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i)))
                    End If
                Else
                        x.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) + 1))
                End If
            Next i
            
        If sRecordSource.AbsolutePosition >= pos_end Then
            Exit Do
        Else
            sRecordSource.MoveNext
            c = c + 1
        End If
    Loop
    Screen.MousePointer = old_pt
    i = 0: c = 0: old_pt = 0
    Set x = Nothing
End Sub

Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function

Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, i, 1)
    Next i
    RightSplitUF = t
    i = 0
    t = ""
End Function

Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, PrimeData, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then getValueAt = rs.Fields(whichField)
    
    Set rs = Nothing
End Function


Public Function AnyRecordExisted(ByRef vRs As ADODB.Recordset) As Boolean
    If vRs.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRs.Requery
    
    If (vRs.BOF = True) And (vRs.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRs.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function


