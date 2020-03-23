VERSION 5.00
Begin VB.Form frmNFeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de NF- e"
   ClientHeight    =   9300
   ClientLeft      =   4785
   ClientTop       =   1455
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10500
   Begin VB.CommandButton cmdTofrmFuncionalidades 
      Caption         =   "Mais Funcinalidades"
      Height          =   495
      Left            =   8640
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrevia 
      Caption         =   "Visualizar Previa"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtCaminho 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "C:\Notas\"
      Top             =   360
      Width           =   5535
   End
   Begin VB.ComboBox cbTpConteudo 
      Height          =   315
      ItemData        =   "frmNFeAPI.frx":0000
      Left            =   8400
      List            =   "frmNFeAPI.frx":000D
      TabIndex        =   12
      Text            =   "txt"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtTpAmb 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Text            =   "2"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   2535
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox cbTpDown 
      Height          =   315
      ItemData        =   "frmNFeAPI.frx":0021
      Left            =   120
      List            =   "frmNFeAPI.frx":0034
      TabIndex        =   5
      Text            =   "XP"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtResult 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6120
      Width           =   10215
   End
   Begin VB.TextBox txtConteudo 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   10215
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Documento para Processamento >>>>>>"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente:"
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   4800
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   690
   End
End
Attribute VB_Name = "frmNFeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    On Error GoTo SAI
    Dim retorno As String
    
    If (txtCaminho.Text <> "") And (txtConteudo.Text <> "") And (cbTpConteudo.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
        
        'Faz a emissão síncrona
        retorno = emitirNFeSincrono(txtConteudo.Text, cbTpConteudo.Text, txtCNPJ.Text, cbTpDown.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value)
        txtResult.Text = retorno
        
        'Abaixo, confira um exemplo de tratamento de retorno da função emitirNFeSincrono
        
        Dim statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros As String
        
        'Lê o statusEnvio
        statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
        'Lê o statusConsulta
        statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
        'Lê o statusDownload
        statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
        'Lê o cStat
        cStat = LerDadosJSON(retorno, "cStat", "", "")
        'Lê a chNFe
        cStat = LerDadosJSON(retorno, "chNFe", "", "")
        'Lê o nProt
        nProt = LerDadosJSON(retorno, "nProt", "", "")
        'Lê o motivo
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        'Lê o nsNRec
        nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
        'Lê os erros
        erros = LerDadosJSON(retorno, "erros", "", "")
        
        'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
        'eles no seu banco de dados antes de prosseguir para o teste abaixo
                 
        'Testa se houve sucesso na emissão
        If (statusEnvio = 200) Or (statusEnvio = -6) Then
            'Testa se houve sucesso na consulta
            If (statusConsulta = 200) Then
                'Testa se a nota foi autorizada
                If (cStat = 100) Then
                    'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
                    MsgBox (motivo)
                     
                    'Testa se o download teve problemas
                    If (statusDownload <> 200) Then
                        MsgBox (motivo)
                    End If
                Else
                    'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
                    MsgBox (motivo)
                End If
            'Caso tenha dado erro na consulta
            Else
                'Aqui você pode mostrar uma mensagem ao usuário
                MsgBox (motivo + Chr(13) + erros)
            End If
        Else
            'Aqui você pode exibir para o usuário o erro que ocorreu no envio
            MsgBox (motivo + Chr(13) + erros)
        End If
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleNFeAPI

End Sub

Private Sub cmdPrevia_Click()
    Dim retorno As String
    Dim i As Integer
    If (txtCaminho.Text <> "") And (txtConteudo.Text <> "") And (cbTpConteudo.Text <> "") Then
        retorno = previaNFeESalvar(txtConteudo.Text, cbTpConteudo.Text, txtCaminho.Text, "PreviaTeste", True)
    Else
        MsgBox ("Todos campos necessarios devem ser preenchidos ['caminho', 'tipo de conteudo', 'conteudo']")
    End If
    
End Sub

Private Sub cmdTofrmFuncionalidades_Click()
    frmFuncionalidades.Show
End Sub
