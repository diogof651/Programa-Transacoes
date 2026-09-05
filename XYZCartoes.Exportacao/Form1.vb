Imports System
Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    Friend Const ConnectionString As String = "Server=localhost;Database=XYZCartoes;Integrated Security=True;TrustServerCertificate=True;"
    Private ReadOnly dataInicialInformada As Nullable(Of Date)
    Private ReadOnly dataFinalInformada As Nullable(Of Date)

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(dataInicial As Nullable(Of Date), dataFinal As Nullable(Of Date))
        Me.New()
        dataInicialInformada = dataInicial
        dataFinalInformada = dataFinal
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Close()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpDataInicial.Value = If(dataInicialInformada.HasValue, dataInicialInformada.Value, Date.Today.AddMonths(-1))
        dtpDataFinal.Value = If(dataFinalInformada.HasValue, dataFinalInformada.Value, Date.Today)
        txtArquivo.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Transacoes.xlsx")
        AtualizaPeriodo()
        lblStatus.Text = "Pronto para exportar."
    End Sub

    Private Sub dtpPeriodo_ValueChanged(sender As Object, e As EventArgs) Handles dtpDataInicial.ValueChanged, dtpDataFinal.ValueChanged
        AtualizaPeriodo()
    End Sub

    Private Sub btnProcurar_Click(sender As Object, e As EventArgs) Handles btnProcurar.Click
        Using dialogo As New SaveFileDialog()
            dialogo.Filter = "Arquivo Excel (*.xlsx)|*.xlsx"
            dialogo.FileName = Path.GetFileName(txtArquivo.Text)
            dialogo.InitialDirectory = Path.GetDirectoryName(txtArquivo.Text)

            If dialogo.ShowDialog(Me) = DialogResult.OK Then
                txtArquivo.Text = dialogo.FileName
            End If
        End Using
    End Sub

    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        Try
            btnExportar.Enabled = False
            lblStatus.Text = "Exportando..."
            Application.DoEvents()

            Dim dataInicial As Date = dtpDataInicial.Value.Date
            Dim dataFinal As Date = dtpDataFinal.Value.Date

            If dataFinal < dataInicial Then
                MessageBox.Show(Me, "A data final deve ser maior ou igual à data inicial.", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                lblStatus.Text = "Período inválido."
                Return
            End If

            dataFinal = dataFinal.AddDays(1)
            Dim quantidade As Integer = TransacoesExcelExporter.Exportar(ConnectionString, dataInicial, dataFinal, txtArquivo.Text)

            lblStatus.Text = quantidade.ToString() & " transação(ões) exportada(s)."
            MessageBox.Show(Me, "Arquivo exportado com sucesso:" & Environment.NewLine & txtArquivo.Text, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            lblStatus.Text = "Falha na exportação."
            MessageBox.Show(Me, ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnExportar.Enabled = True
        End Try
    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Close()
    End Sub

    Private Sub AtualizaPeriodo()
        lblPeriodo.Text = "Período: " & dtpDataInicial.Value.ToString("dd/MM/yyyy") & " até " & dtpDataFinal.Value.ToString("dd/MM/yyyy")
    End Sub
End Class
