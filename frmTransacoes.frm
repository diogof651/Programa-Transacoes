VERSION 5.00
Begin VB.Form frmTransacoes 
   Caption         =   "Cadastro de Transações"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   Icon            =   "frmTransacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDados 
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   7830
      TabIndex        =   9
      Top             =   0
      Width           =   7895
      Begin VB.TextBox txtIdTransacao 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cboStatusTransacao 
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtNumeroCartao 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtDescricao 
         Height          =   975
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   5775
      End
      Begin VB.TextBox txtDataTransacao 
         Height          =   315
         Left            =   5040
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtValorTransacao 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblIdTransacao 
         Caption         =   "Id Transação"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label lblStatusTransacao 
         Caption         =   "Status"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblNumeroCartao 
         Caption         =   "Número Cartão"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label lblDataTransacao 
         Caption         =   "Data/Hora"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lblValorTransacao 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7845
      TabIndex        =   6
      Top             =   2505
      Width           =   7905
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   420
         Left            =   4920
         TabIndex        =   5
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Anterior"
         Height          =   420
         Left            =   5880
         TabIndex        =   7
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Próximo"
         Height          =   420
         Left            =   6840
         TabIndex        =   8
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Ca&ncelar"
         Height          =   420
         Left            =   3960
         TabIndex        =   4
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   420
         Left            =   3000
         TabIndex        =   3
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   420
         Left            =   2040
         TabIndex        =   2
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   420
         Left            =   1080
         TabIndex        =   1
         Top             =   90
         Width           =   840
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "&Inserir"
         Height          =   420
         Left            =   120
         TabIndex        =   0
         Top             =   90
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const OP_NENHUM As Integer = 0
Private Const OP_INSERIR As Integer = 1
Private Const OP_ALTERAR As Integer = 2

Dim rsDados As Recordset
Private mOperacao As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo TrataErro

    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo TrataErro

    Set rsDados = New Recordset
    CarregaStatus

    mOperacao = OP_NENHUM
    MostraRegistro Maior
    MostraBotao True
    MostraCampos
    txtIdTransacao.Enabled = False

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If Not rsDados Is Nothing Then
        If rsDados.State = adStateOpen Then rsDados.Close
    End If
    Set rsDados = Nothing
End Sub

Private Sub cmdInserir_Click()
    On Error GoTo TrataErro

    LimpaCampos
    mOperacao = OP_INSERIR
    MostraBotao False
    txtDataTransacao.Text = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    txtNumeroCartao.SetFocus

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdAlterar_Click()
    On Error GoTo TrataErro

    If Not RegistroAberto Then Exit Sub

    If TransacaoAprovada(CLng(txtIdTransacao.Text)) Then
        MsgBox "Transação aprovada não pode ser alterada.", vbExclamation, "Atenção"
        Exit Sub
    End If

    mOperacao = OP_ALTERAR
    MostraBotao False
    txtNumeroCartao.SetFocus

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo ExcluirErro

    If Not RegistroAberto Then Exit Sub

    If MsgBox("Confirma a exclusão do registro?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão") = vbYes Then
        ExcluiTransacao CLng(txtIdTransacao.Text)
        MostraRegistro Maior
        MostraBotao True
        MostraCampos
    End If

    cmdInserir.SetFocus

    Exit Sub
ExcluirErro:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível excluir a transação.", vbCritical, "Erro"
End Sub

Private Sub cmdConfirmar_Click()
    On Error GoTo TrataErro

    If mOperacao = OP_NENHUM Then Exit Sub
    If Not CamposValidos Then Exit Sub

    If mOperacao = OP_INSERIR Then
        InsereTransacao
    ElseIf mOperacao = OP_ALTERAR Then
        AlteraTransacao
    End If

    mOperacao = OP_NENHUM
    MostraRegistro CLng(txtIdTransacao.Text)
    MostraBotao True
    MostraCampos
    cmdInserir.SetFocus

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo TrataErro

    mOperacao = OP_NENHUM
    MostraBotao True

    If Trim$(txtIdTransacao.Text) <> "" Then
        MostraRegistro CLng(txtIdTransacao.Text)
    Else
        MostraRegistro Maior
    End If

    MostraCampos
    cmdInserir.SetFocus

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo TrataErro

    frmConTransacoes.CodBusca = -1
    frmConTransacoes.RegAtual = 0
    If Trim$(txtIdTransacao.Text) <> "" Then frmConTransacoes.RegAtual = CDbl(txtIdTransacao.Text)

    frmConTransacoes.Show vbModal

    If frmConTransacoes.CodBusca > 0 Then
        MostraRegistro frmConTransacoes.CodBusca
        MostraCampos
    End If

    Unload frmConTransacoes

    Exit Sub
TrataErro:
    Unload frmConTransacoes
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo TrataErro

    Dim Codigo As Double

    If rsDados Is Nothing Then Exit Sub
    If rsDados.State <> adStateOpen Then Exit Sub
    If rsDados.BOF Or rsDados.EOF Then Exit Sub

    Codigo = Anterior
    If Codigo <> 0 Then
        MostraRegistro Codigo
        MostraCampos
    End If

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdNext_Click()
    On Error GoTo TrataErro

    Dim Codigo As Double

    If rsDados Is Nothing Then Exit Sub
    If rsDados.State <> adStateOpen Then Exit Sub
    If rsDados.BOF Or rsDados.EOF Then Exit Sub

    Codigo = Proximo
    If Codigo <> 0 Then
        MostraRegistro Codigo
        MostraCampos
    End If

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub CarregaStatus()
    On Error GoTo UsaPadrao

    PreencheCombo cboStatusTransacao, "cboTransacoes"
    If cboStatusTransacao.ListCount > 0 Then
        cboStatusTransacao.ListIndex = 0
        Exit Sub
    End If

UsaPadrao:
    cboStatusTransacao.Clear
    cboStatusTransacao.AddItem "Aprovada"
    cboStatusTransacao.AddItem "Pendente"
    cboStatusTransacao.AddItem "Cancelada"
    cboStatusTransacao.ListIndex = 1
End Sub

Private Sub MostraRegistro(Registro As Double)
    If rsDados.State = adStateOpen Then rsDados.Close
    If Registro <> 0 Then rsDados.Open "SELECT * FROM Transacoes WHERE (Id_Transacao = " & CStr(Registro) & ")", frmPrincipal.db, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub MostraCampos()
    LimpaCampos

    If rsDados Is Nothing Then Exit Sub
    If rsDados.State <> adStateOpen Then Exit Sub
    If rsDados.BOF Or rsDados.EOF Then Exit Sub

    txtIdTransacao.Text = CStr(rsDados!Id_Transacao)
    txtNumeroCartao.Text = CStr(rsDados!Numero_Cartao)
    txtValorTransacao.Text = Format$(rsDados!Valor_Transacao, "0.00")
    txtDataTransacao.Text = Format$(rsDados!Data_Transacao, "dd/mm/yyyy hh:nn:ss")

    If Not IsNull(rsDados!Descricao) Then txtDescricao.Text = CStr(rsDados!Descricao)
    If Not IsNull(rsDados!Status_Transacao) Then SelecionaStatus CStr(rsDados!Status_Transacao)
End Sub

Private Function Maior() As Double
    Maior = 0

    Dim rsMaior As New Recordset
    rsMaior.Open "SELECT MAX(Id_Transacao) AS Maior FROM dbo.Transacoes", frmPrincipal.db, adOpenForwardOnly, adLockReadOnly

    If Not rsMaior.BOF Then
        If Not IsNull(rsMaior!Maior) Then Maior = rsMaior!Maior
    End If

    rsMaior.Close
    Set rsMaior = Nothing
End Function

Private Function Proximo() As Double
    Proximo = 0

    Dim rsProximo As New Recordset
    rsProximo.Open "SELECT MIN(Id_Transacao) AS Proximo FROM dbo.Transacoes WHERE (Id_Transacao > " & rsDados!Id_Transacao & ")", frmPrincipal.db, adOpenForwardOnly, adLockReadOnly

    If Not rsProximo.BOF Then
        If Not IsNull(rsProximo!Proximo) Then Proximo = rsProximo!Proximo
    End If

    rsProximo.Close
    Set rsProximo = Nothing
End Function

Private Function Anterior() As Double
    Anterior = 0

    Dim rsAnterior As New Recordset
    rsAnterior.Open "SELECT MAX(Id_Transacao) AS Anterior FROM dbo.Transacoes WHERE (Id_Transacao < " & rsDados!Id_Transacao & ")", frmPrincipal.db, adOpenForwardOnly, adLockReadOnly

    If Not rsAnterior.BOF Then
        If Not IsNull(rsAnterior!Anterior) Then Anterior = rsAnterior!Anterior
    End If

    rsAnterior.Close
    Set rsAnterior = Nothing
End Function

Private Function RegistroAberto() As Boolean
    RegistroAberto = False

    If Trim$(txtIdTransacao.Text) = "" Then
        MsgBox "Nenhum registro selecionado.", vbExclamation, "Atenção"
        Exit Function
    End If

    RegistroAberto = True
End Function

Private Sub LimpaCampos()
    txtIdTransacao.Text = ""
    txtNumeroCartao.Text = ""
    txtValorTransacao.Text = ""
    txtDataTransacao.Text = ""
    txtDescricao.Text = ""
    If cboStatusTransacao.ListCount > 0 Then cboStatusTransacao.ListIndex = -1
End Sub

Private Function CamposValidos() As Boolean
    CamposValidos = False

    If Len(Trim$(txtNumeroCartao.Text)) <> 16 Or Not IsNumeric(txtNumeroCartao.Text) Then
        MsgBox "Número do cartão deve ter 16 dígitos.", vbExclamation, "Atenção"
        txtNumeroCartao.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtValorTransacao.Text) Then
        MsgBox "Valor da transação inválido.", vbExclamation, "Atenção"
        txtValorTransacao.SetFocus
        Exit Function
    End If

    If CCur(txtValorTransacao.Text) <= 0 Then
        MsgBox "Valor da transação deve ser positivo.", vbExclamation, "Atenção"
        txtValorTransacao.SetFocus
        Exit Function
    End If

    If Not IsDate(txtDataTransacao.Text) Then
        MsgBox "Data da transação inválida.", vbExclamation, "Atenção"
        txtDataTransacao.SetFocus
        Exit Function
    End If

    If Len(txtDescricao.Text) > 255 Then
        MsgBox "Descrição deve ter no maximo 255 caracteres.", vbExclamation, "Atenção"
        txtDescricao.SetFocus
        Exit Function
    End If

    If cboStatusTransacao.ListIndex < 0 Then
        MsgBox "Informe o status da transação.", vbExclamation, "Atenção"
        cboStatusTransacao.SetFocus
        Exit Function
    End If

    If mOperacao = OP_ALTERAR Then
        If Trim$(txtIdTransacao.Text) = "" Then
            MsgBox "Informe o código da transação para alterar.", vbExclamation, "Atenção"
            txtIdTransacao.SetFocus
            Exit Function
        End If

        If TransacaoAprovada(CLng(txtIdTransacao.Text)) Then
            MsgBox "Transação aprovada não pode ser alterada.", vbExclamation, "Atenção"
            Exit Function
        End If
    End If

    CamposValidos = True
End Function

Private Sub InsereTransacao()
    Dim cmd As ADODB.Command
    Dim rsNovo As ADODB.Recordset

    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = frmPrincipal.db
        .CommandType = adCmdText
        .CommandText = "SET NOCOUNT ON; INSERT INTO dbo.Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) VALUES (?, ?, ?, ?, ?); SELECT CONVERT(INT, SCOPE_IDENTITY()) AS Id_Transacao"
        .Parameters.Append cmd.CreateParameter("Numero_Cartao", adVarChar, adParamInput, 16, Trim$(txtNumeroCartao.Text))
        .Parameters.Append cmd.CreateParameter("Valor_Transacao", adCurrency, adParamInput, , CCur(txtValorTransacao.Text))
        .Parameters.Append cmd.CreateParameter("Data_Transacao", adDate, adParamInput, , CDate(txtDataTransacao.Text))
        .Parameters.Append cmd.CreateParameter("Descricao", adVarChar, adParamInput, 255, txtDescricao.Text)
        .Parameters.Append cmd.CreateParameter("Status_Transacao", adVarChar, adParamInput, 20, cboStatusTransacao.Text)
    End With

    Set rsNovo = cmd.Execute
    If Not rsNovo Is Nothing Then
        If Not rsNovo.EOF Then txtIdTransacao.Text = CStr(rsNovo!Id_Transacao)
        rsNovo.Close
    End If

    Set rsNovo = Nothing
    Set cmd = Nothing
End Sub

Private Sub AlteraTransacao()
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = frmPrincipal.db
        .CommandType = adCmdText
        .CommandText = "UPDATE dbo.Transacoes SET Numero_Cartao = ?, Valor_Transacao = ?, Data_Transacao = ?, Descricao = ?, Status_Transacao = ? WHERE Id_Transacao = ?"
        .Parameters.Append cmd.CreateParameter("Numero_Cartao", adVarChar, adParamInput, 16, Trim$(txtNumeroCartao.Text))
        .Parameters.Append cmd.CreateParameter("Valor_Transacao", adCurrency, adParamInput, , CCur(txtValorTransacao.Text))
        .Parameters.Append cmd.CreateParameter("Data_Transacao", adDate, adParamInput, , CDate(txtDataTransacao.Text))
        .Parameters.Append cmd.CreateParameter("Descricao", adVarChar, adParamInput, 255, txtDescricao.Text)
        .Parameters.Append cmd.CreateParameter("Status_Transacao", adVarChar, adParamInput, 20, cboStatusTransacao.Text)
        .Parameters.Append cmd.CreateParameter("Id_Transacao", adInteger, adParamInput, , CLng(txtIdTransacao.Text))
        .Execute
    End With

    Set cmd = Nothing
End Sub

Private Sub ExcluiTransacao(ByVal IdTransacao As Long)
    On Error GoTo TrataErro

    frmPrincipal.db.Execute "DELETE FROM dbo.Transacoes WHERE Id_Transacao = " & CStr(IdTransacao)

    Exit Sub
TrataErro:
    Dim NumeroErro As Long
    Dim FonteErro As String
    Dim DescricaoErro As String

    NumeroErro = Err.Number
    FonteErro = Err.Source
    DescricaoErro = Err.Description

    GravaErroTransacao IdTransacao, "DELETE", NumeroErro, DescricaoErro
    Err.Raise NumeroErro, FonteErro, DescricaoErro
End Sub

Private Sub GravaErroTransacao(ByVal IdTransacao As Long, ByVal Operacao As String, ByVal NumeroErro As Long, ByVal DescricaoErro As String)
    On Error Resume Next

    Dim rsErro As New Recordset

    rsErro.Open "SELECT * FROM dbo.Transacoes_Erro WHERE 1 = 0", frmPrincipal.db, adOpenKeyset, adLockOptimistic

    rsErro.AddNew
    rsErro!Id_Transacao = IdTransacao
    rsErro!Operacao = Operacao
    rsErro!Numero_Erro = NumeroErro
    rsErro!Descricao_Erro = DescricaoErro
    rsErro.Update

    rsErro.Close
    Set rsErro = Nothing
End Sub

Private Function TransacaoAprovada(ByVal IdTransacao As Long) As Boolean
    Dim cmd As ADODB.Command
    Dim rsBusca As ADODB.Recordset

    TransacaoAprovada = False

    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = frmPrincipal.db
        .CommandType = adCmdText
        .CommandText = "SELECT Status_Transacao FROM dbo.Transacoes WHERE Id_Transacao = ?"
        .Parameters.Append cmd.CreateParameter("Id_Transacao", adInteger, adParamInput, , IdTransacao)
    End With

    Set rsBusca = cmd.Execute
    If Not rsBusca.EOF Then
        If UCase$(CStr(rsBusca!Status_Transacao)) = "APROVADA" Then TransacaoAprovada = True
    End If

    rsBusca.Close
    Set rsBusca = Nothing
    Set cmd = Nothing
End Function

Private Sub SelecionaStatus(ByVal StatusTransacao As String)
    Dim i As Integer

    For i = 0 To cboStatusTransacao.ListCount - 1
        If UCase$(cboStatusTransacao.List(i)) = UCase$(StatusTransacao) Then
            cboStatusTransacao.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub MostraBotao(bVal As Boolean)
    cmdInserir.Enabled = bVal
    cmdAlterar.Enabled = bVal
    cmdExcluir.Enabled = bVal
    cmdBuscar.Enabled = bVal
    cmdPrevious.Enabled = bVal
    cmdNext.Enabled = bVal

    cmdConfirmar.Enabled = Not bVal
    cmdCancelar.Enabled = Not bVal
    picDados.Enabled = Not bVal
End Sub
