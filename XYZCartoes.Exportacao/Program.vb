Imports System
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Public Module Program
    <STAThread>
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Dim dataInicial As Nullable(Of Date) = Nothing
        Dim dataFinal As Nullable(Of Date) = Nothing
        Dim argumentos() As String = Environment.GetCommandLineArgs()
        Dim exportarAutomatico As Boolean = Array.Exists(argumentos, Function(item) String.Equals(item, "--auto", StringComparison.OrdinalIgnoreCase))

        If argumentos.Length >= 3 Then
            Dim dataInicialInformada As Date
            Dim dataFinalInformada As Date

            If TentaConverterData(argumentos(1), dataInicialInformada) AndAlso TentaConverterData(argumentos(2), dataFinalInformada) Then
                dataInicial = dataInicialInformada
                dataFinal = dataFinalInformada
            Else
                MessageBox.Show("Período informado inválido. Selecione as datas na tela.", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If

        If exportarAutomatico Then
            If Not dataInicial.HasValue OrElse Not dataFinal.HasValue Then
                MessageBox.Show("Informe data inicial e data final para exportar automaticamente.", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            ExportarAutomaticamente(dataInicial.Value, dataFinal.Value)
            Return
        End If

        Application.Run(New Form1(dataInicial, dataFinal))
    End Sub

    Private Function TentaConverterData(valor As String, ByRef data As Date) As Boolean
        Return Date.TryParseExact(valor,
                                  New String() {"dd/MM/yyyy", "yyyy-MM-dd", "yyyyMMdd"},
                                  CultureInfo.GetCultureInfo("pt-BR"),
                                  DateTimeStyles.None,
                                  data) OrElse Date.TryParse(valor, CultureInfo.GetCultureInfo("pt-BR"), DateTimeStyles.None, data)
    End Function

    Private Sub ExportarAutomaticamente(dataInicial As Date, dataFinal As Date)
        Try
            If dataFinal < dataInicial Then
                MessageBox.Show("A data final deve ser maior ou igual à data inicial.", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            Dim nomeArquivo As String = "Transacoes_" & dataInicial.ToString("yyyyMMdd") & "_" & dataFinal.ToString("yyyyMMdd") & ".xlsx"
            Dim caminhoArquivo As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), nomeArquivo)
            Dim quantidade As Integer = TransacoesExcelExporter.Exportar(Form1.ConnectionString, dataInicial.Date, dataFinal.Date.AddDays(1), caminhoArquivo)

            MessageBox.Show(quantidade.ToString() & " transação(ões) exportada(s)." & Environment.NewLine & caminhoArquivo, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
