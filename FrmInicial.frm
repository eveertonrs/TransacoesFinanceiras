VERSION 5.00
Begin VB.Form FrmInicial 
   Caption         =   "Tela Inicial"
   ClientHeight    =   7215
   ClientLeft      =   4755
   ClientTop       =   2235
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   14940
   Begin VB.Menu mnuViewToolbar 
      Caption         =   "&Cadastro"
      Begin VB.Menu Menu1 
         Caption         =   "Transações"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "&Relatorios"
      Begin VB.Menu Menu3 
         Caption         =   "Relatório Mensal"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   ""
   End
End
Attribute VB_Name = "FrmInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Transação_Click()

End Sub

Private Sub Menu1_Click()
FrmCadastro.Show
End Sub

Private Sub Menu3_Click()
FrmRelatorio.Show
End Sub
