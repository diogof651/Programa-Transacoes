<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer
    Friend WithEvents lblTitulo As System.Windows.Forms.Label
    Friend WithEvents lblPeriodo As System.Windows.Forms.Label
    Friend WithEvents lblArquivo As System.Windows.Forms.Label
    Friend WithEvents txtArquivo As System.Windows.Forms.TextBox
    Friend WithEvents btnProcurar As System.Windows.Forms.Button
    Friend WithEvents btnExportar As System.Windows.Forms.Button
    Friend WithEvents btnFechar As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblDataInicial As System.Windows.Forms.Label
    Friend WithEvents dtpDataInicial As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDataFinal As System.Windows.Forms.Label
    Friend WithEvents dtpDataFinal As System.Windows.Forms.DateTimePicker

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.lblPeriodo = New System.Windows.Forms.Label()
        Me.lblArquivo = New System.Windows.Forms.Label()
        Me.txtArquivo = New System.Windows.Forms.TextBox()
        Me.btnProcurar = New System.Windows.Forms.Button()
        Me.btnExportar = New System.Windows.Forms.Button()
        Me.btnFechar = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblDataInicial = New System.Windows.Forms.Label()
        Me.dtpDataInicial = New System.Windows.Forms.DateTimePicker()
        Me.lblDataFinal = New System.Windows.Forms.Label()
        Me.dtpDataFinal = New System.Windows.Forms.DateTimePicker()
        Me.SuspendLayout()
        '
        'lblTitulo
        '
        Me.lblTitulo.AutoSize = True
        Me.lblTitulo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitulo.Location = New System.Drawing.Point(18, 18)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(277, 20)
        Me.lblTitulo.TabIndex = 0
        Me.lblTitulo.Text = "Exportação de Transações Excel"
        '
        'lblPeriodo
        '
        Me.lblPeriodo.AutoSize = True
        Me.lblPeriodo.Location = New System.Drawing.Point(20, 56)
        Me.lblPeriodo.Name = "lblPeriodo"
        Me.lblPeriodo.Size = New System.Drawing.Size(104, 13)
        Me.lblPeriodo.TabIndex = 1
        Me.lblPeriodo.Text = "Período"
        '
        'lblDataInicial
        '
        Me.lblDataInicial.AutoSize = True
        Me.lblDataInicial.Location = New System.Drawing.Point(20, 88)
        Me.lblDataInicial.Name = "lblDataInicial"
        Me.lblDataInicial.Size = New System.Drawing.Size(60, 13)
        Me.lblDataInicial.TabIndex = 2
        Me.lblDataInicial.Text = "Data inicial"
        '
        'dtpDataInicial
        '
        Me.dtpDataInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDataInicial.Location = New System.Drawing.Point(23, 106)
        Me.dtpDataInicial.Name = "dtpDataInicial"
        Me.dtpDataInicial.Size = New System.Drawing.Size(128, 20)
        Me.dtpDataInicial.TabIndex = 3
        '
        'lblDataFinal
        '
        Me.lblDataFinal.AutoSize = True
        Me.lblDataFinal.Location = New System.Drawing.Point(177, 88)
        Me.lblDataFinal.Name = "lblDataFinal"
        Me.lblDataFinal.Size = New System.Drawing.Size(52, 13)
        Me.lblDataFinal.TabIndex = 4
        Me.lblDataFinal.Text = "Data final"
        '
        'dtpDataFinal
        '
        Me.dtpDataFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDataFinal.Location = New System.Drawing.Point(180, 106)
        Me.dtpDataFinal.Name = "dtpDataFinal"
        Me.dtpDataFinal.Size = New System.Drawing.Size(128, 20)
        Me.dtpDataFinal.TabIndex = 5
        '
        'lblArquivo
        '
        Me.lblArquivo.AutoSize = True
        Me.lblArquivo.Location = New System.Drawing.Point(20, 150)
        Me.lblArquivo.Name = "lblArquivo"
        Me.lblArquivo.Size = New System.Drawing.Size(46, 13)
        Me.lblArquivo.TabIndex = 6
        Me.lblArquivo.Text = "Arquivo"
        '
        'txtArquivo
        '
        Me.txtArquivo.Location = New System.Drawing.Point(23, 168)
        Me.txtArquivo.Name = "txtArquivo"
        Me.txtArquivo.Size = New System.Drawing.Size(445, 20)
        Me.txtArquivo.TabIndex = 7
        '
        'btnProcurar
        '
        Me.btnProcurar.Location = New System.Drawing.Point(482, 166)
        Me.btnProcurar.Name = "btnProcurar"
        Me.btnProcurar.Size = New System.Drawing.Size(90, 24)
        Me.btnProcurar.TabIndex = 8
        Me.btnProcurar.Text = "Procurar..."
        Me.btnProcurar.UseVisualStyleBackColor = True
        '
        'btnExportar
        '
        Me.btnExportar.Location = New System.Drawing.Point(372, 222)
        Me.btnExportar.Name = "btnExportar"
        Me.btnExportar.Size = New System.Drawing.Size(96, 30)
        Me.btnExportar.TabIndex = 9
        Me.btnExportar.Text = "Exportar"
        Me.btnExportar.UseVisualStyleBackColor = True
        '
        'btnFechar
        '
        Me.btnFechar.Location = New System.Drawing.Point(482, 222)
        Me.btnFechar.Name = "btnFechar"
        Me.btnFechar.Size = New System.Drawing.Size(90, 30)
        Me.btnFechar.TabIndex = 10
        Me.btnFechar.Text = "Fechar"
        Me.btnFechar.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(20, 231)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(40, 13)
        Me.lblStatus.TabIndex = 11
        Me.lblStatus.Text = "Status"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 277)
        Me.Controls.Add(Me.dtpDataFinal)
        Me.Controls.Add(Me.lblDataFinal)
        Me.Controls.Add(Me.dtpDataInicial)
        Me.Controls.Add(Me.lblDataInicial)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnFechar)
        Me.Controls.Add(Me.btnExportar)
        Me.Controls.Add(Me.btnProcurar)
        Me.Controls.Add(Me.txtArquivo)
        Me.Controls.Add(Me.lblArquivo)
        Me.Controls.Add(Me.lblPeriodo)
        Me.Controls.Add(Me.lblTitulo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "XYZCartoes - Exportação Excel"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
