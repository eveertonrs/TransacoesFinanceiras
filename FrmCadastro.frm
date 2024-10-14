VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCadastro 
   Caption         =   " Cadastro"
   ClientHeight    =   9525
   ClientLeft      =   7905
   ClientTop       =   1935
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   7365
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Tag             =   "asdsa"
      Top             =   4800
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4048
         _Version        =   393216
      End
      Begin VB.TextBox txtFiltroValor 
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtFiltroData 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtFiltroCartao 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton btnConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton btnExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblValorTransacao 
         Caption         =   "Valor Transacão:"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblNumerocartao 
         Caption         =   "Data Transação:"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblTransacao 
         Caption         =   "Nr Cartão:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Cadastro 
      Caption         =   "Cadastro"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   3240
         Width           =   4095
      End
      Begin VB.TextBox txtDataTransacao 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtValorTransacao 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtNumeroCartao 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton btnInserir 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label LblDescricao 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblDataTansacao 
         Caption         =   "Data Transação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblValorTransacao 
         Caption         =   "Valor Transação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblNumerocartao 
         Caption         =   "Numero Cartão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmCadastro"
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
Private Sub ErrorHandler(msg As String)
    MsgBox "Erro: " & msg, vbCritical
End Sub

Private Sub btnInserir_Click()
    On Error GoTo erro
    
    If txtNumeroCartao = "" Then
        MsgBox "Número do cartão é obrigatório!", vbInformation
        Exit Sub
    End If
    
    If txtValorTransacao = "" Then
        MsgBox "Valor da Transação é obrigatório!", vbInformation
        Exit Sub
    End If
    
    If txtDataTransacao = "" Then
        MsgBox "Data da Transação é obrigatória!", vbInformation
        Exit Sub
    End If
    
    If txtDescricao = "" Then
        MsgBox "Descrição da Transação é obrigatória!", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(txtNumeroCartao.Text) Or Len(txtNumeroCartao.Text) < 4 Then
        MsgBox "O Número do cartão deve conter apenas números e ter pelo menos 4 dígitos.", vbCritical
        Exit Sub
    End If

    Dim valorTransacao As String
    valorTransacao = Replace(txtValorTransacao.Text, ",", ".")
    
    If Not IsNumeric(valorTransacao) Then
        MsgBox "Valor da Transação inválido! Use o formato decimal com ponto (ex: 15.99).", vbCritical
        Exit Sub
    End If

    Dim valorDecimal As Double
    valorDecimal = FormatNumber(CDbl(valorTransacao), 2)
    
    Call AbrirConexao
    
    Dim sql As String
    sql = "INSERT INTO tbdTransacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao) VALUES ('" & Replace(txtNumeroCartao.Text, "'", "''") & "', " & valorDecimal & ", '" & txtDataTransacao.Text & "', '" & Replace(txtDescricao.Text, "'", "''") & "')"
    

    conn.Execute sql
    MsgBox "Transação inserida com sucesso!", vbInformation
    
    FecharConexao
    Call PreencherGrid
    Call LimparCampos
    Exit Sub
    
erro:
    ErrorHandler Err.Description
    FecharConexao
End Sub
Private Sub btnEditar_Click()
    On Error GoTo erro
    

    If MSFlexGrid.row <= 0 Then
        MsgBox "Selecione uma transação para editar!", vbInformation
        Exit Sub
    End If
    
    Dim idTransacao As Long
    idTransacao = MSFlexGrid.TextMatrix(MSFlexGrid.row, 0)
    
    Dim campoAlterado As Boolean
    campoAlterado = False
    
    If txtNumeroCartao.Text <> MSFlexGrid.TextMatrix(MSFlexGrid.row, 1) Then campoAlterado = True
    If CDbl(txtValorTransacao.Text) <> CDbl(MSFlexGrid.TextMatrix(MSFlexGrid.row, 2)) Then campoAlterado = True
    If txtDataTransacao.Text <> MSFlexGrid.TextMatrix(MSFlexGrid.row, 3) Then campoAlterado = True
    If txtDescricao.Text <> MSFlexGrid.TextMatrix(MSFlexGrid.row, 4) Then campoAlterado = True

    If Not campoAlterado Then
        MsgBox "Para editar, é necessário alterar pelo menos um campo!", vbInformation
        Exit Sub
    End If

    AbrirConexao
    
    Dim sql As String
    sql = "UPDATE tbdTransacoes SET Numero_Cartao = '" & Replace(txtNumeroCartao.Text, "'", "''") & _
          "', Valor_Transacao = " & CDbl(txtValorTransacao.Text) & _
          ", Data_Transacao = '" & txtDataTransacao.Text & _
          "', Descricao = '" & Replace(txtDescricao.Text, "'", "''") & _
          "' WHERE Id_Transacao = " & idTransacao
    
    conn.Execute sql
    MsgBox "Transação atualizada com sucesso!", vbInformation
    FecharConexao
    
    Call PreencherGrid
    Call LimparCampos
    Exit Sub

erro:
    ErrorHandler Err.Description
    FecharConexao
End Sub

Private Sub btnExcluir_Click()
    On Error GoTo erro
    
    If MSFlexGrid.row <= 0 Then
        MsgBox "Selecione uma transação para excluir!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir esta transação?", vbYesNo + vbQuestion, "Confirmação de Exclusão") = vbNo Then
        Exit Sub
    End If
    
    Dim idTransacao As Long
    idTransacao = MSFlexGrid.TextMatrix(MSFlexGrid.row, 0)
    
    AbrirConexao
    
    Dim sql As String
    sql = "DELETE FROM tbdTransacoes WHERE Id_Transacao = " & idTransacao
    
    conn.Execute sql
    MsgBox "Transação excluída com sucesso!", vbInformation
    FecharConexao
    
    Call PreencherGrid
    Call LimparCampos
    Exit Sub
    
erro:
    ErrorHandler Err.Description
    FecharConexao
End Sub



Private Sub btnConsultar_Click()
    On Error GoTo erro
    AbrirConexao
    
    Dim sql As String
    sql = "SELECT * FROM tbdTransacoes WHERE 1=1"
    
    If txtFiltroCartao.Text <> "" Then
        sql = sql & " AND Numero_Cartao = '" & Replace(txtFiltroCartao.Text, "'", "''") & "'"
    End If
    
    If txtFiltroData.Text <> "" Then
        sql = sql & " AND Data_Transacao = '" & txtFiltroData.Text & "'"
    End If
    
    If txtFiltroValor.Text <> "" Then
        sql = sql & " AND Valor_Transacao = " & CDbl(txtFiltroValor.Text)
    End If

    Set rs = New ADODB.Recordset

    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    MSFlexGrid.Clear
    MSFlexGrid.Rows = 1

    MSFlexGrid.Cols = rs.Fields.Count

    For i = 0 To rs.Fields.Count - 1
        MSFlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
        
        Select Case i
            Case 0
                MSFlexGrid.ColWidth(i) = 800
            Case 1
                MSFlexGrid.ColWidth(i) = 1500
            Case 2
                MSFlexGrid.ColWidth(i) = 1200
            Case 3
                MSFlexGrid.ColWidth(i) = 1200
            Case Else
                MSFlexGrid.ColWidth(i) = 2500
        End Select
    Next i

    If Not rs.EOF Then
        Dim row As Integer
        row = 1

        Do While Not rs.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            
            For i = 0 To rs.Fields.Count - 1
                MSFlexGrid.TextMatrix(row, i) = rs.Fields(i).Value
            Next i

            rs.MoveNext

            row = row + 1
        Loop
    Else
        MsgBox "Nenhum registro encontrado."
    End If

    rs.Close
    FecharConexao
    Exit Sub

erro:
    MsgBox "Erro: " & Err.Description
    FecharConexao
End Sub


Private Sub PreencherGrid()
    On Error GoTo erro
    AbrirConexao

    Dim sql As String
    sql = "SELECT * FROM tbdTransacoes"

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    MSFlexGrid.Clear
    MSFlexGrid.Rows = 1
    MSFlexGrid.Cols = rs.Fields.Count

    For i = 0 To rs.Fields.Count - 1
        MSFlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
        
        Select Case i
            Case 0
                MSFlexGrid.ColWidth(i) = 800
            Case 1
                MSFlexGrid.ColWidth(i) = 1500
            Case 2
                MSFlexGrid.ColWidth(i) = 1200
            Case 3
                MSFlexGrid.ColWidth(i) = 1200
            Case Else
                MSFlexGrid.ColWidth(i) = 2500
        End Select
    Next i

    If Not rs.EOF Then
        Dim row As Integer
        row = 1

        Do While Not rs.EOF

            MSFlexGrid.Rows = MSFlexGrid.Rows + 1

            For i = 0 To rs.Fields.Count - 1
                MSFlexGrid.TextMatrix(row, i) = rs.Fields(i).Value & ""
            Next i
            rs.MoveNext
            row = row + 1
        Loop
    Else
        MsgBox "Nenhum registro encontrado."
    End If

    rs.Close
    FecharConexao
    Exit Sub

erro:
    MsgBox "Erro: " & Err.Description
    FecharConexao
End Sub

Private Sub Form_Load()
txtDataTransacao.Text = Format(Date, "dd/MM/yyyy")
End Sub

Private Sub MSFlexGrid_Click()
    If MSFlexGrid.row > 0 Then
        txtNumeroCartao.Text = MSFlexGrid.TextMatrix(MSFlexGrid.row, 1)
        txtValorTransacao.Text = MSFlexGrid.TextMatrix(MSFlexGrid.row, 2)
        txtDataTransacao.Text = MSFlexGrid.TextMatrix(MSFlexGrid.row, 3)
        txtDescricao.Text = MSFlexGrid.TextMatrix(MSFlexGrid.row, 4)
    End If
End Sub

Private Sub txtDataTransacao_LostFocus()
    On Error Resume Next
    If Len(txtDataTransacao.Text) > 0 Then
        Dim dt As Date
        dt = CDate(txtDataTransacao.Text)
        txtDataTransacao.Text = Format(dt, "dd/mm/yyyy")
    End If
End Sub

Private Sub txtDataTransacao_Change()
    On Error Resume Next

    Dim texto As String
    texto = Replace(txtDataTransacao.Text, "/", "")
    
    Select Case Len(texto)
        Case 1 To 2
            txtDataTransacao.Text = texto
        Case 3 To 4
            txtDataTransacao.Text = Left(texto, 2) & "/" & Mid(texto, 3)
        Case 5 To 8
            txtDataTransacao.Text = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
    End Select
    
    txtDataTransacao.SelStart = Len(txtDataTransacao.Text)
End Sub

Private Sub txtFiltroData_Change()
   On Error Resume Next

    Dim texto As String
    texto = Replace(txtFiltroData.Text, "/", "")

    Select Case Len(texto)
        Case 1 To 2
            txtFiltroData.Text = texto
        Case 3 To 4
            txtFiltroData.Text = Left(texto, 2) & "/" & Mid(texto, 3)
        Case 5 To 8
            txtFiltroData.Text = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
    End Select
    
    txtFiltroData.SelStart = Len(txtDataTransacao.Text)
End Sub
Private Sub LimparCampos()
    On Error Resume Next
    
    txtNumeroCartao.Text = ""
    txtValorTransacao.Text = ""
    txtDataTransacao.Text = ""
    txtDescricao.Text = ""
    'txtFiltroCartao.Text = ""
    'txtFiltroData.Text = ""
    'txtFiltroValor.Text = ""
End Sub

