VERSION 5.00
Begin VB.Form FrmRelatorio 
   Caption         =   "Relatório Mensal"
   ClientHeight    =   3015
   ClientLeft      =   10230
   ClientTop       =   4260
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdGerarRelatorio 
      Caption         =   "Gerar Relatório Mensal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label 
      Caption         =   "Transações do último mês"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "FrmRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub AbrirConexao()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-SUARC27\SQLEXPRESS;Initial Catalog=dtbTransacao;Integrated Security=SSPI;"
    conn.Open
End Sub

Private Sub FecharConexao()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

Private Sub cmdGerarRelatorio_Click()
    Call ExportarTransacoesParaExcel
End Sub

Private Sub ExportarTransacoesParaExcel()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim row As Integer


    Call AbrirConexao
    sql = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, dbo.CategorizarTransacao(Valor_Transacao) AS Categoria " & _
          "FROM tbdTransacoes WHERE Data_Transacao >= DATEADD(MONTH, -1, GETDATE())"

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly


    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)


    For Col = 0 To rs.Fields.Count - 1
        xlSheet.Cells(1, Col + 1).Value = rs.Fields(Col).Name
    Next Col


    row = 2
    Do While Not rs.EOF
        For Col = 0 To rs.Fields.Count - 1
            xlSheet.Cells(row, Col + 1).Value = rs.Fields(Col).Value
        Next Col
        rs.MoveNext
        row = row + 1
    Loop


    Dim savePath As String
    savePath = ShowSaveFileDialogExcel()

    If savePath <> "" Then
        xlBook.SaveAs savePath
        MsgBox "Relatório exportado com sucesso!", vbInformation
    Else
        MsgBox "Operação cancelada.", vbExclamation
    End If

    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    rs.Close
    FecharConexao
End Sub

Private Function ShowSaveFileDialogExcel() As String
    Dim xlApp As Object
    Dim savePath As String


    Set xlApp = CreateObject("Excel.Application")


    savePath = xlApp.GetSaveAsFilename(InitialFileName:="Relatório.xlsx", _
                                       FileFilter:="Excel Files (*.xlsx), *.xlsx", _
                                       Title:="Salvar Relatório")


    If savePath = "False" Then
        ShowSaveFileDialogExcel = ""
    Else
        ShowSaveFileDialogExcel = savePath
    End If


    xlApp.Quit
    Set xlApp = Nothing
End Function


