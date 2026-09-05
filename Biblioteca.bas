Attribute VB_Name = "Module1"
Function NumReg(NomeTabela As String) As Double
    NumReg = 0
    Dim rsNumReg As New Recordset
    rsNumReg.Open "SELECT COUNT(*) AS Registros FROM " & NomeTabela, frmPrincipal.db, adOpenForwardOnly, adLockReadOnly
    If Not rsNumReg.BOF Then
        If Not IsNull(rsNumReg!Registros) Then NumReg = rsNumReg!Registros
    End If
    rsNumReg.Close: Set rsNumReg = Nothing
End Function

Public Sub PreencheCombo(Combo As Object, modulo As String)
    Dim rsBusca As New Recordset
    rsBusca.Open "SELECT Descricao FROM " & modulo & " WHERE cboNome = '" & Combo.Name & "' ORDER BY ID", frmPrincipal.db, adOpenForwardOnly, adLockReadOnly
    'SELECT Descricao FROM cboTransacoes WHERE cboNome = 'cboTransacoes' ORDER BY ID
    Combo.Clear
    While Not rsBusca.EOF
        Combo.AddItem rsBusca!Descricao
        rsBusca.MoveNext
    Wend
    rsBusca.Close: Set rsBusca = Nothing
End Sub


Public Sub MostraErro(ByVal Err As ErrObject, ByVal Caption As String)

    Dim FoundPos As Integer
    Dim Tamanho As Integer
    Dim Msg As Boolean
    Dim DescricaoErro As String

    Screen.MousePointer = vbDefault

    Msg = False
    DescricaoErro = UCase$(Err.Description)

    Select Case Err.Number

        Case -2147217873, -2147217900

            If InStr(1, DescricaoErro, "PRIMARY KEY", vbTextCompare) > 0 _
            Or InStr(1, DescricaoErro, "UNIQUE KEY", vbTextCompare) > 0 _
            Or InStr(1, DescricaoErro, "UNIQUE INDEX", vbTextCompare) > 0 Then

                MsgBox "Informação já cadastrada!", vbCritical, "Erro"
                Msg = True

            ElseIf InStr(1, DescricaoErro, _
                         "DELETE STATEMENT CONFLICTED WITH THE REFERENCE CONSTRAINT", _
                         vbTextCompare) > 0 Then

                FoundPos = InStr(1, Err.Description, ", column", vbTextCompare)

                If FoundPos > 0 Then

                    Tamanho = (FoundPos - 1) - _
                              (InStr(1, Err.Description, "table ", vbTextCompare) + 7)

                    If Tamanho > 0 Then
                        MsgBox "Exclusão da informação não é permitida." & vbCrLf & _
                               Caption & " ligado(a) ao(à) " & vbCrLf & _
                               Mid$(Err.Description, _
                                    InStr(1, Err.Description, "table ", vbTextCompare) + 11, _
                                    Tamanho - 4) & ".", _
                               vbCritical, "Erro"
                    Else
                        MsgBox "Exclusão da informação não é permitida." & vbCrLf & _
                               Err.Description, _
                               vbCritical, "Erro"
                    End If

                Else
                    MsgBox "Exclusão da informação não é permitida." & vbCrLf & _
                           Err.Description, _
                           vbCritical, "Erro"
                End If

                Msg = True

            ElseIf InStr(1, DescricaoErro, "COLUMN REFERENCE", vbTextCompare) > 0 Then

                FoundPos = (InStr(1, Err.Description, ", column", vbTextCompare) - 1) - _
                           (InStr(1, Err.Description, "table ", vbTextCompare) + 7)

                MsgBox "Exclusão da informação não é permitida." & vbCrLf & _
                       Caption & " ligado(a) ao(à) " & _
                       Mid$(Err.Description, _
                            InStr(1, Err.Description, "table ", vbTextCompare) + 7, _
                            FoundPos) & ".", _
                       vbCritical, "Erro"

                Msg = True

            ElseIf InStr(1, DescricaoErro, "TABLE REFERENCE", vbTextCompare) > 0 Then

                FoundPos = (Len(Err.Description) - 2) - _
                           (InStr(1, Err.Description, "table '", vbTextCompare) + 7)

                MsgBox "Exclusão da informação não é permitida." & vbCrLf & _
                       Caption & " ligado(a) ao(à) " & _
                       Mid$(Err.Description, _
                            InStr(1, Err.Description, "table '", vbTextCompare) + 7, _
                            FoundPos) & ".", _
                       vbCritical, "Erro"

                Msg = True

            ElseIf InStr(1, DescricaoErro, "COLUMN FOREIGN KEY", vbTextCompare) > 0 Then

                FoundPos = (InStr(1, Err.Description, ", column", vbTextCompare) - 1) - _
                           (InStr(1, Err.Description, "table ", vbTextCompare) + 7)

                MsgBox "Informação não está cadastrada em " & _
                       Mid$(Err.Description, _
                            InStr(1, Err.Description, "table ", vbTextCompare) + 7, _
                            FoundPos) & ".", _
                       vbCritical, "Erro"

                Msg = True

            End If

            If Not Msg Then Errors Err


        Case 364

            Err.Clear


        Case 400

            MsgBox "Formulário já está aberto na tela!", _
                   vbCritical, "Erro"


        Case -2147217865

            MsgBox "Não existe log para esta informação!", _
                   vbCritical, "Erro"
        Case 3021, 3265
            MsgBox "Arquivo vazio!", _
                   vbExclamation, "Atenção"
        Case -2147467259

            MsgBox "Não foi possível concluir a operação no banco de dados." & vbCrLf & _
                   "Tente novamente.", _
                   vbCritical, "Erro"
        Case Else

            Errors Err

    End Select

End Sub

Private Sub Errors(ByVal Err As ErrObject)

    Dim errLoop As ADODB.Error
    Dim strError As String
    Dim Contador As Integer

    Contador = 0

    For Each errLoop In frmPrincipal.db.Errors

        Contador = Contador + 1

        strError = "Erro: " & errLoop.Description

        If errLoop.Source <> "" Then
            strError = strError & vbCrLf & _
                       "Fonte: " & errLoop.Source
        End If

        MsgBox strError, vbCritical, "Erro"

    Next errLoop

    If Contador = 0 Then

        strError = "Erro: " & Err.Description

        If Err.Source <> "" Then
            strError = strError & vbCrLf & _
                       "Fonte: " & Err.Source
        End If

        MsgBox strError, vbCritical, "Erro"

    End If

End Sub


