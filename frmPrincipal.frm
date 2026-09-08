VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8805
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTransacoes 
      Caption         =   "Transações"
   End
   Begin VB.Menu mnuExportarExcel 
      Caption         =   "Exportar Excel"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Connection

Private Sub Form_Load()
    On Error GoTo LoadError
    
    
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.CommandTimeout = 600     '05/09/18 (c/ Marta): estava 300, Laércio pediu para dobrar.
    
    db.Open "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=XYZCartoes;Integrated Security=SSPI;Encrypt=True;TrustServerCertificate=True;"
    
    Dim rsTransacoes As New Recordset
    rsTransacoes.Open ("Select * from Transacoes"), frmPrincipal.db, adOpenForwardOnly, adLockReadOnly
    
    Exit Sub
LoadError:
    If Err.Number = -2147467259 Then
        MsgBox "Não é possível estabelecer a conexão com o servidor!" & vbCrLf & "Tente efetuar o Logon novamente.", vbExclamation, "Atenção"
        Unload Me
    End If
End Sub

Private Sub mnuTransacoes_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmTransacoes.Show vbModal
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuExportarExcel_Click()
    On Error GoTo TrataErro

    Dim CaminhoExportador As String

    CaminhoExportador = App.Path & "\XYZCartoes.Exportacao\bin\Debug\XYZCartoes.Exportacao.exe"

    If Dir(CaminhoExportador) = "" Then
        MsgBox "Exportador não encontrado:" & vbCrLf & CaminhoExportador, vbExclamation, "Exportar Excel"
        Exit Sub
    End If

    Shell """" & CaminhoExportador & """", vbNormalFocus
    Exit Sub

TrataErro:
    MsgBox "Não foi possível abrir o exportador." & vbCrLf & Err.Description, vbCritical, "Exportar Excel"
End Sub
