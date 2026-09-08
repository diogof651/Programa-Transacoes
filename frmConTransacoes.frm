VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmConTransacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Transações"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11520
   Icon            =   "frmConTransacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   405
      Left            =   10260
      TabIndex        =   12
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "Exportar Excel"
      Height          =   405
      Left            =   7980
      TabIndex        =   10
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Selecionar"
      Height          =   405
      Left            =   9120
      TabIndex        =   11
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "&Próximo"
      Height          =   405
      Left            =   2940
      TabIndex        =   9
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "&Anterior"
      Height          =   405
      Left            =   1800
      TabIndex        =   8
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   405
      Left            =   10260
      TabIndex        =   6
      Top             =   600
      Width           =   1000
   End
   Begin VB.ComboBox cboStatusTransacao 
      Height          =   315
      Left            =   8340
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtValorTransacao 
      Height          =   315
      Left            =   6900
      TabIndex        =   4
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox txtDataInicial 
      Height          =   315
      Left            =   4140
      TabIndex        =   2
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox txtDataFinal 
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox txtNumeroCartao 
      Height          =   315
      Left            =   180
      MaxLength       =   16
      TabIndex        =   0
      Top             =   600
      Width           =   2115
   End
   Begin VB.TextBox txtDescricao 
      Height          =   315
      Left            =   2460
      MaxLength       =   60
      TabIndex        =   1
      Top             =   600
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid grdBusca 
      Height          =   4275
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   7541
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Id_Transacao"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Numero_Cartao"
         Caption         =   "Cartão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Data_Transacao"
         Caption         =   "Data/Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Valor_Transacao"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Status_Transacao"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Descricao"
         Caption         =   "Descrição"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2039,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2039,811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3374,929
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPagina 
      Caption         =   "Página 0 de 0"
      Height          =   255
      Left            =   180
      TabIndex        =   17
      Top             =   5790
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   8340
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblValor 
      Caption         =   "Valor"
      Height          =   255
      Left            =   6900
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblDataInicial 
      Caption         =   "Data Inicial"
      Height          =   255
      Left            =   4140
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDataFinal 
      Caption         =   "Data Final"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDescricao 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   2460
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblNumeroCartao 
      Caption         =   "Número Cartão"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   360
      Width           =   1395
   End
   Begin VB.Label lblContador 
      Height          =   255
      Left            =   4140
      TabIndex        =   20
      Top             =   5790
      Width           =   2775
   End
End
Attribute VB_Name = "frmConTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodBusca As Double
Public RegAtual As Double

Private Const QTDE_POR_PAGINA As Long = 20

Private mPaginaAtual As Long
Private mTotalRegistros As Long
Private mTotalPaginas As Long
Private rsBusca As Recordset
Private mFormatandoData As Boolean

Private Sub Form_Load()
    On Error GoTo TrataErro

    CodBusca = -1
    mPaginaAtual = 1
    CarregaStatus
    CarregaConsulta

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo TrataErro

    Select Case KeyCode
        Case vbKeyEscape
            CodBusca = -1
            Unload Me
        Case vbKeyReturn
            If TypeName(Me.ActiveControl) = "DataGrid" Then SelecionaRegistro
    End Select

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Set grdBusca.DataSource = Nothing
    If Not rsBusca Is Nothing Then
        If rsBusca.State = adStateOpen Then rsBusca.Close
    End If
    Set rsBusca = Nothing
End Sub

Private Sub cmdPesquisar_Click()
    On Error GoTo TrataErro

    mPaginaAtual = 1
    CarregaConsulta

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdAnterior_Click()
    On Error GoTo TrataErro

    If mPaginaAtual > 1 Then
        mPaginaAtual = mPaginaAtual - 1
        CarregaConsulta
    End If

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdProximo_Click()
    On Error GoTo TrataErro

    If mPaginaAtual < mTotalPaginas Then
        mPaginaAtual = mPaginaAtual + 1
        CarregaConsulta
    End If

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdSelecionar_Click()
    SelecionaRegistro
End Sub

Private Sub cmdExportarExcel_Click()
    On Error GoTo TrataErro

    Dim CaminhoExportador As String
    Dim Comando As String

    If Not PeriodoExportacaoValido Then Exit Sub

    CaminhoExportador = App.Path & "\XYZCartoes.Exportacao\bin\Debug\XYZCartoes.Exportacao.exe"

    If Dir(CaminhoExportador) = "" Then
        MsgBox "Exportador não encontrado:" & vbCrLf & CaminhoExportador, vbExclamation, "Exportar Excel"
        Exit Sub
    End If

    Comando = """" & CaminhoExportador & """ " & _
              """" & Format$(CDate(txtDataInicial.Text), "dd/mm/yyyy") & """ " & _
              """" & Format$(CDate(txtDataFinal.Text), "dd/mm/yyyy") & """ --auto"

    Shell Comando, vbNormalFocus
    Exit Sub

TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub cmdSair_Click()
    CodBusca = -1
    Unload Me
End Sub

Private Sub grdBusca_DblClick()
    SelecionaRegistro
End Sub

Private Sub grdBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SelecionaRegistro
    ElseIf KeyCode = vbKeyEscape Then
        CodBusca = -1
        Unload Me
    End If
End Sub

Private Sub txtNumeroCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdPesquisar_Click
    End If
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdPesquisar_Click
    End If
End Sub

Private Sub txtDataInicial_Change()
    FormataCampoData txtDataInicial
End Sub

Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    TrataTeclaData KeyAscii
End Sub

Private Sub txtDataFinal_Change()
    FormataCampoData txtDataFinal
End Sub

Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    TrataTeclaData KeyAscii
End Sub

Private Sub TrataTeclaData(ByRef KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdPesquisar_Click
    ElseIf KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub FormataCampoData(ByVal Campo As TextBox)
    If mFormatandoData Then Exit Sub

    Dim Digitos As String
    Dim Texto As String
    Dim i As Integer

    Digitos = ""
    For i = 1 To Len(Campo.Text)
        If Mid$(Campo.Text, i, 1) >= "0" And Mid$(Campo.Text, i, 1) <= "9" Then
            Digitos = Digitos & Mid$(Campo.Text, i, 1)
        End If
    Next i

    If Len(Digitos) > 8 Then Digitos = Left$(Digitos, 8)

    Texto = Digitos
    If Len(Digitos) > 2 Then Texto = Left$(Digitos, 2) & "/" & Mid$(Digitos, 3)
    If Len(Digitos) > 4 Then Texto = Left$(Digitos, 2) & "/" & Mid$(Digitos, 3, 2) & "/" & Mid$(Digitos, 5)

    If Campo.Text <> Texto Then
        mFormatandoData = True
        Campo.Text = Texto
        Campo.SelStart = Len(Campo.Text)
        mFormatandoData = False
    End If
End Sub
Private Sub txtValorTransacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdPesquisar_Click
    End If
End Sub

Private Sub cboStatusTransacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdPesquisar_Click
    End If
End Sub

Private Sub CarregaStatus()
    cboStatusTransacao.Clear
    cboStatusTransacao.AddItem ""
    cboStatusTransacao.AddItem "Aprovada"
    cboStatusTransacao.AddItem "Pendente"
    cboStatusTransacao.AddItem "Cancelada"
    cboStatusTransacao.ListIndex = 0
End Sub

Private Sub CarregaConsulta()
    On Error GoTo TrataErro

    Dim SQL As String
    Dim Filtro As String
    Dim LinhaInicial As Long
    Dim LinhaFinal As Long

    If Not FiltrosValidos Then Exit Sub

    Filtro = MontaFiltro
    AtualizaTotal Filtro

    If mTotalRegistros = 0 Then
        mPaginaAtual = 1
    ElseIf mPaginaAtual > mTotalPaginas Then
        mPaginaAtual = mTotalPaginas
    End If

    LinhaInicial = ((mPaginaAtual - 1) * QTDE_POR_PAGINA) + 1
    LinhaFinal = mPaginaAtual * QTDE_POR_PAGINA

    SQL = "WITH Consulta AS (" & _
          " SELECT ROW_NUMBER() OVER (ORDER BY Id_Transacao) AS Linha," & _
          " Id_Transacao, Numero_Cartao," & _
          " CONVERT(VARCHAR(10), Data_Transacao, 103) + ' ' + CONVERT(VARCHAR(8), Data_Transacao, 108) AS Data_Transacao," & _
          " Valor_Transacao, Status_Transacao, Descricao" & _
          " FROM dbo.Transacoes WHERE 1 = 1" & Filtro & _
          ") SELECT Id_Transacao, Numero_Cartao, Data_Transacao, Valor_Transacao, Status_Transacao, Descricao" & _
          " FROM Consulta WHERE Linha BETWEEN " & CStr(LinhaInicial) & " AND " & CStr(LinhaFinal) & _
          " ORDER BY Linha"

    Set grdBusca.DataSource = Nothing
    If Not rsBusca Is Nothing Then
        If rsBusca.State = adStateOpen Then rsBusca.Close
    End If
    Set rsBusca = New Recordset
    rsBusca.CursorLocation = adUseClient
    rsBusca.Open SQL, frmPrincipal.db, adOpenStatic, adLockReadOnly

    Set grdBusca.DataSource = rsBusca
    FormataGrid
    AtualizaPaginacao
    PosicionaRegistroAtual

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub

Private Sub AtualizaTotal(ByVal Filtro As String)
    Dim rsTotal As New Recordset

    mTotalRegistros = 0
    mTotalPaginas = 0

    rsTotal.Open "SELECT COUNT(*) AS Total FROM dbo.Transacoes WHERE 1 = 1" & Filtro, frmPrincipal.db, adOpenForwardOnly, adLockReadOnly
    If Not rsTotal.BOF Then
        If Not IsNull(rsTotal!Total) Then mTotalRegistros = CLng(rsTotal!Total)
    End If
    rsTotal.Close
    Set rsTotal = Nothing

    If mTotalRegistros > 0 Then
        mTotalPaginas = ((mTotalRegistros - 1) \ QTDE_POR_PAGINA) + 1
    Else
        mTotalPaginas = 1
    End If
End Sub

Private Function MontaFiltro() As String
    Dim Filtro As String

    If Trim$(txtNumeroCartao.Text) <> "" Then
        Filtro = Filtro & " AND Numero_Cartao LIKE '" & Replace(Trim$(txtNumeroCartao.Text), "'", "''") & "%'"
    End If

    If Trim$(txtDescricao.Text) <> "" Then
        Filtro = Filtro & " AND Descricao LIKE '%" & Replace(Trim$(txtDescricao.Text), "'", "''") & "%'"
    End If

    If Trim$(txtDataInicial.Text) <> "" And Trim$(txtDataFinal.Text) <> "" Then
        Filtro = Filtro & " AND Data_Transacao >= CONVERT(DATETIME, '" & Format$(CDate(txtDataInicial.Text), "yyyymmdd") & "', 112)"
        Filtro = Filtro & " AND Data_Transacao < DATEADD(DAY, 1, CONVERT(DATETIME, '" & Format$(CDate(txtDataFinal.Text), "yyyymmdd") & "', 112))"
    ElseIf Trim$(txtDataInicial.Text) <> "" Then
        Filtro = Filtro & " AND Data_Transacao >= CONVERT(DATETIME, '" & Format$(CDate(txtDataInicial.Text), "yyyymmdd") & "', 112)"
        Filtro = Filtro & " AND Data_Transacao < DATEADD(DAY, 1, CONVERT(DATETIME, '" & Format$(CDate(txtDataInicial.Text), "yyyymmdd") & "', 112))"
    ElseIf Trim$(txtDataFinal.Text) <> "" Then
        Filtro = Filtro & " AND Data_Transacao >= CONVERT(DATETIME, '" & Format$(CDate(txtDataFinal.Text), "yyyymmdd") & "', 112)"
        Filtro = Filtro & " AND Data_Transacao < DATEADD(DAY, 1, CONVERT(DATETIME, '" & Format$(CDate(txtDataFinal.Text), "yyyymmdd") & "', 112))"
    End If

    If Trim$(txtValorTransacao.Text) <> "" Then
        Filtro = Filtro & " AND Valor_Transacao = " & Replace(Format$(CCur(txtValorTransacao.Text), "0.00"), ",", ".")
    End If

    If cboStatusTransacao.ListIndex > 0 Then
        Filtro = Filtro & " AND Status_Transacao = '" & Replace(cboStatusTransacao.Text, "'", "''") & "'"
    End If

    MontaFiltro = Filtro
End Function

Private Function FiltrosValidos() As Boolean
    FiltrosValidos = False

    If Trim$(txtDataInicial.Text) <> "" Then
        If Not IsDate(txtDataInicial.Text) Then
            MsgBox "Data inicial inválida.", vbExclamation, "Atenção"
            txtDataInicial.SetFocus
            Exit Function
        End If
    End If

    If Trim$(txtDataFinal.Text) <> "" Then
        If Not IsDate(txtDataFinal.Text) Then
            MsgBox "Data final inválida.", vbExclamation, "Atenção"
            txtDataFinal.SetFocus
            Exit Function
        End If
    End If

    If Trim$(txtDataInicial.Text) <> "" And Trim$(txtDataFinal.Text) <> "" Then
        If CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
            MsgBox "Data final deve ser maior ou igual à data inicial.", vbExclamation, "Atenção"
            txtDataFinal.SetFocus
            Exit Function
        End If
    End If

    If Trim$(txtValorTransacao.Text) <> "" Then
        If Not IsNumeric(txtValorTransacao.Text) Then
            MsgBox "Valor da transação inválido.", vbExclamation, "Atenção"
            txtValorTransacao.SetFocus
            Exit Function
        End If
    End If

    FiltrosValidos = True
End Function

Private Sub AtualizaPaginacao()
    lblPagina.Caption = "Página " & CStr(mPaginaAtual) & " de " & CStr(mTotalPaginas)
    lblContador.Caption = Format$(mTotalRegistros, "#,##0") & " registro(s) encontrado(s)"

    cmdAnterior.Enabled = (mPaginaAtual > 1)
    cmdProximo.Enabled = (mPaginaAtual < mTotalPaginas)
End Sub

Private Sub FormataGrid()
    On Error Resume Next

    grdBusca.Columns(0).Width = 800
    grdBusca.Columns(0).Caption = "Código"
    grdBusca.Columns(1).Width = 1800
    grdBusca.Columns(1).Caption = "Cartão"
    grdBusca.Columns(2).Width = 1800
    grdBusca.Columns(2).Caption = "Data/Hora"
    grdBusca.Columns(3).Width = 900
    grdBusca.Columns(3).Caption = "Valor"
    grdBusca.Columns(4).Width = 1200
    grdBusca.Columns(4).Caption = "Status"
    grdBusca.Columns(5).Width = 3300
    grdBusca.Columns(5).Caption = "Descrição"
End Sub

Private Sub PosicionaRegistroAtual()
    On Error Resume Next

    If rsBusca Is Nothing Then Exit Sub
    If rsBusca.State <> adStateOpen Then Exit Sub
    If rsBusca.BOF Or rsBusca.EOF Then Exit Sub

    If RegAtual <> 0 Then
        rsBusca.MoveFirst
        rsBusca.Find "Id_Transacao = " & CStr(RegAtual)
        If rsBusca.EOF Then rsBusca.MoveFirst
    Else
        rsBusca.MoveFirst
    End If
End Sub

Private Sub SelecionaRegistro()
    On Error GoTo TrataErro

    If rsBusca Is Nothing Then Exit Sub
    If rsBusca.State <> adStateOpen Then Exit Sub
    If rsBusca.BOF Or rsBusca.EOF Then Exit Sub

    CodBusca = CDbl(rsBusca!Id_Transacao)
    Unload Me

    Exit Sub
TrataErro:
    MostraErro Err, Me.Caption
End Sub
Private Function PeriodoExportacaoValido() As Boolean
    PeriodoExportacaoValido = False

    If Trim$(txtDataInicial.Text) = "" Then
        MsgBox "Informe a data inicial para exportar.", vbExclamation, "Exportar Excel"
        txtDataInicial.SetFocus
        Exit Function
    End If

    If Trim$(txtDataFinal.Text) = "" Then
        MsgBox "Informe a data final para exportar.", vbExclamation, "Exportar Excel"
        txtDataFinal.SetFocus
        Exit Function
    End If

    If Not IsDate(txtDataInicial.Text) Then
        MsgBox "Data inicial inválida.", vbExclamation, "Exportar Excel"
        txtDataInicial.SetFocus
        Exit Function
    End If

    If Not IsDate(txtDataFinal.Text) Then
        MsgBox "Data final inválida.", vbExclamation, "Exportar Excel"
        txtDataFinal.SetFocus
        Exit Function
    End If

    If CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
        MsgBox "Data final deve ser maior ou igual à data inicial.", vbExclamation, "Exportar Excel"
        txtDataFinal.SetFocus
        Exit Function
    End If

    PeriodoExportacaoValido = True
End Function
