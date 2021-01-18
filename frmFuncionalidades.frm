VERSION 5.00
Begin VB.Form frmFuncionalidades 
   Caption         =   "Mais funcionalidades"
   ClientHeight    =   5520
   ClientLeft      =   6585
   ClientTop       =   1965
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6615
   Begin VB.Frame Frame4 
      Caption         =   "NS NFe API"
      Height          =   2535
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5895
      Begin VB.Frame Frame1 
         Caption         =   "Envia Email"
         Height          =   1455
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdTofrmEnviaEmail 
            Caption         =   "Ir para Fomulario"
            Height          =   495
            Left            =   480
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consultar Situação NFe"
         Height          =   1455
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdTofrmConsultaSituacaoNFe 
            Caption         =   "Ir para Formulario"
            Height          =   495
            Left            =   480
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NS PAINEL"
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   5895
      Begin VB.Frame Frame5 
         Caption         =   "Cadastrar Licença Painel"
         Height          =   1575
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   2655
         Begin VB.CommandButton cmdTofrmCadLic 
            Caption         =   "Add Licença Formulario"
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   600
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmFuncionalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTofrmEnviaEmail_Click()
    frmEnviaEmail.Show
End Sub

Private Sub cmdTofrmCadLic_Click()
    MsgBox ("Formulario ainda não disponivel, por favor fale com nossa equipe")
    'frmCadLic.Show
End Sub

Private Sub cmdTofrmConsultaSituacaoNFe_Click()
    frmConsultaSituacaoNFe.Show
End Sub
