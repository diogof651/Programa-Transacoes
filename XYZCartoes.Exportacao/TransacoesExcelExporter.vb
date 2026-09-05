Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Security
Imports System.Text

Public NotInheritable Class TransacoesExcelExporter
    Private Sub New()
    End Sub

    Public Shared Function Exportar(connectionString As String, dataInicial As Date, dataFinalExclusiva As Date, arquivo As String) As Integer
        Dim tabela As DataTable = BuscarTransacoes(connectionString, dataInicial, dataFinalExclusiva)
        GerarXlsx(tabela, arquivo)
        Return tabela.Rows.Count
    End Function

    Private Shared Function BuscarTransacoes(connectionString As String, dataInicial As Date, dataFinalExclusiva As Date) As DataTable
        Dim tabela As New DataTable("Transacoes")

        Const sql As String =
            "SELECT Id_Transacao, Numero_Cartao, Data_Transacao, Valor_Transacao, Status_Transacao, Descricao " &
            "FROM dbo.Transacoes " &
            "WHERE Data_Transacao >= @DataInicial " &
            "AND Data_Transacao < @DataFinal " &
            "ORDER BY Data_Transacao, Id_Transacao"

        Using conexao As New SqlConnection(connectionString)
            Using comando As New SqlCommand(sql, conexao)
                comando.Parameters.Add("@DataInicial", SqlDbType.DateTime).Value = dataInicial
                comando.Parameters.Add("@DataFinal", SqlDbType.DateTime).Value = dataFinalExclusiva

                Using adaptador As New SqlDataAdapter(comando)
                    adaptador.Fill(tabela)
                End Using
            End Using
        End Using

        Return tabela
    End Function

    Private Shared Sub GerarXlsx(tabela As DataTable, arquivo As String)
        Dim pasta As String = Path.GetDirectoryName(arquivo)
        If String.IsNullOrWhiteSpace(pasta) Then
            Throw New InvalidOperationException("Informe um caminho válido para o arquivo.")
        End If

        Directory.CreateDirectory(pasta)
        If File.Exists(arquivo) Then File.Delete(arquivo)

        Using pacote As ZipArchive = ZipFile.Open(arquivo, ZipArchiveMode.Create)
            AdicionarTexto(pacote, "[Content_Types].xml", ContentTypesXml())
            AdicionarTexto(pacote, "_rels/.rels", RelsXml())
            AdicionarTexto(pacote, "xl/workbook.xml", WorkbookXml())
            AdicionarTexto(pacote, "xl/_rels/workbook.xml.rels", WorkbookRelsXml())
            AdicionarTexto(pacote, "xl/styles.xml", StylesXml())
            AdicionarTexto(pacote, "xl/worksheets/sheet1.xml", WorksheetXml(tabela))
        End Using
    End Sub

    Private Shared Sub AdicionarTexto(pacote As ZipArchive, nome As String, conteudo As String)
        Dim entrada As ZipArchiveEntry = pacote.CreateEntry(nome, CompressionLevel.Optimal)
        Using writer As New StreamWriter(entrada.Open(), New UTF8Encoding(False))
            writer.Write(conteudo)
        End Using
    End Sub

    Private Shared Function WorksheetXml(tabela As DataTable) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
        sb.AppendLine("<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">")
        sb.AppendLine("<sheetViews><sheetView workbookViewId=""0""/></sheetViews>")
        sb.AppendLine("<sheetFormatPr defaultRowHeight=""15""/>")
        sb.AppendLine("<cols>")
        sb.AppendLine("<col min=""1"" max=""1"" width=""14"" customWidth=""1""/>")
        sb.AppendLine("<col min=""2"" max=""2"" width=""22"" customWidth=""1""/>")
        sb.AppendLine("<col min=""3"" max=""3"" width=""22"" customWidth=""1""/>")
        sb.AppendLine("<col min=""4"" max=""4"" width=""16"" customWidth=""1""/>")
        sb.AppendLine("<col min=""5"" max=""5"" width=""18"" customWidth=""1""/>")
        sb.AppendLine("<col min=""6"" max=""6"" width=""45"" customWidth=""1""/>")
        sb.AppendLine("</cols>")
        sb.AppendLine("<sheetData>")
        sb.AppendLine("<row r=""1"">")
        AdicionarCelulaTexto(sb, 1, 1, "Id_Transacao", 1)
        AdicionarCelulaTexto(sb, 1, 2, "Numero_Cartao", 1)
        AdicionarCelulaTexto(sb, 1, 3, "Data_Transacao", 1)
        AdicionarCelulaTexto(sb, 1, 4, "Valor_Transacao", 1)
        AdicionarCelulaTexto(sb, 1, 5, "Status_Transacao", 1)
        AdicionarCelulaTexto(sb, 1, 6, "Descricao", 1)
        sb.AppendLine("</row>")

        Dim linhaExcel As Integer = 2
        For Each row As DataRow In tabela.Rows
            sb.Append("<row r=""").Append(linhaExcel.ToString(CultureInfo.InvariantCulture)).AppendLine(""">")
            AdicionarCelulaNumero(sb, linhaExcel, 1, row("Id_Transacao"))
            AdicionarCelulaTexto(sb, linhaExcel, 2, row("Numero_Cartao"), 0)
            AdicionarCelulaDataTexto(sb, linhaExcel, 3, row("Data_Transacao"))
            AdicionarCelulaDecimal(sb, linhaExcel, 4, row("Valor_Transacao"))
            AdicionarCelulaTexto(sb, linhaExcel, 5, row("Status_Transacao"), 0)
            AdicionarCelulaTexto(sb, linhaExcel, 6, row("Descricao"), 0)
            sb.AppendLine("</row>")
            linhaExcel += 1
        Next

        sb.AppendLine("</sheetData>")
        sb.AppendLine("<autoFilter ref=""A1:F" & Math.Max(1, linhaExcel - 1).ToString(CultureInfo.InvariantCulture) & """/>")
        sb.AppendLine("<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>")
        sb.AppendLine("</worksheet>")
        Return sb.ToString()
    End Function

    Private Shared Sub AdicionarCelulaTexto(sb As StringBuilder, linha As Integer, coluna As Integer, valor As Object, estilo As Integer)
        Dim texto As String = If(valor Is Nothing OrElse Convert.IsDBNull(valor), "", Convert.ToString(valor, CultureInfo.CurrentCulture))
        sb.Append("<c r=""").Append(ReferenciaCelula(linha, coluna)).Append(""" t=""inlineStr"" s=""").Append(estilo.ToString(CultureInfo.InvariantCulture)).Append("""><is><t>")
        sb.Append(SecurityElement.Escape(texto))
        sb.AppendLine("</t></is></c>")
    End Sub

    Private Shared Sub AdicionarCelulaDataTexto(sb As StringBuilder, linha As Integer, coluna As Integer, valor As Object)
        If valor Is Nothing OrElse Convert.IsDBNull(valor) Then
            AdicionarCelulaTexto(sb, linha, coluna, "", 0)
            Return
        End If

        Dim data As Date = Convert.ToDateTime(valor, CultureInfo.CurrentCulture)
        AdicionarCelulaTexto(sb, linha, coluna, data.ToString("dd/MM/yyyy HH:mm:ss"), 0)
    End Sub

    Private Shared Sub AdicionarCelulaNumero(sb As StringBuilder, linha As Integer, coluna As Integer, valor As Object)
        If valor Is Nothing OrElse Convert.IsDBNull(valor) Then
            AdicionarCelulaTexto(sb, linha, coluna, "", 0)
            Return
        End If

        sb.Append("<c r=""").Append(ReferenciaCelula(linha, coluna)).Append(""" t=""n""><v>")
        sb.Append(Convert.ToInt64(valor, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture))
        sb.AppendLine("</v></c>")
    End Sub

    Private Shared Sub AdicionarCelulaDecimal(sb As StringBuilder, linha As Integer, coluna As Integer, valor As Object)
        If valor Is Nothing OrElse Convert.IsDBNull(valor) Then
            AdicionarCelulaTexto(sb, linha, coluna, "", 0)
            Return
        End If

        Dim numero As Decimal = Convert.ToDecimal(valor, CultureInfo.CurrentCulture)
        sb.Append("<c r=""").Append(ReferenciaCelula(linha, coluna)).Append(""" t=""n"" s=""2""><v>")
        sb.Append(numero.ToString(CultureInfo.InvariantCulture))
        sb.AppendLine("</v></c>")
    End Sub

    Private Shared Function ReferenciaCelula(linha As Integer, coluna As Integer) As String
        Dim nomeColuna As String = String.Empty
        Dim n As Integer = coluna
        While n > 0
            n -= 1
            nomeColuna = ChrW(65 + (n Mod 26)) & nomeColuna
            n = n \ 26
        End While
        Return nomeColuna & linha.ToString(CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ContentTypesXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8""?>" &
               "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">" &
               "<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>" &
               "<Default Extension=""xml"" ContentType=""application/xml""/>" &
               "<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>" &
               "<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>" &
               "<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>" &
               "</Types>"
    End Function

    Private Shared Function RelsXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8""?>" &
               "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" &
               "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>" &
               "</Relationships>"
    End Function

    Private Shared Function WorkbookXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8""?>" &
               "<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">" &
               "<sheets><sheet name=""Transações"" sheetId=""1"" r:id=""rId1""/></sheets>" &
               "</workbook>"
    End Function

    Private Shared Function WorkbookRelsXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8""?>" &
               "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" &
               "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>" &
               "<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>" &
               "</Relationships>"
    End Function

    Private Shared Function StylesXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8""?>" &
               "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">" &
               "<fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><sz val=""11""/><name val=""Calibri""/></font></fonts>" &
               "<fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills>" &
               "<borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>" &
               "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>" &
               "<cellXfs count=""3""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1""/><xf numFmtId=""4"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1""/></cellXfs>" &
               "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>" &
               "</styleSheet>"
    End Function
End Class
