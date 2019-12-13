VERSION 5.00
Begin VB.Form frmNFeAPIRetorno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Busca Retorno de Processamento Documento"
   ClientHeight    =   9630
   ClientLeft      =   5370
   ClientTop       =   1545
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   27412.13
   ScaleMode       =   0  'User
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Nota"
      Height          =   1575
      Left            =   120
      TabIndex        =   29
      Top             =   2160
      Width           =   8415
      Begin VB.TextBox txtNatOperacao 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   8055
      End
      Begin VB.TextBox txtTotalNota 
         Height          =   315
         Left            =   4200
         TabIndex        =   32
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtCNPJ_CPF_Dest 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Operação"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1620
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total da Nota"
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF Destinatário"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.TextBox txtTpAmb 
      Height          =   315
      Left            =   5880
      TabIndex        =   27
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox comboTpDown 
      Height          =   315
      ItemData        =   "frmNFeAPIRetorno.frx":0000
      Left            =   120
      List            =   "frmNFeAPIRetorno.frx":0013
      TabIndex        =   26
      Text            =   "X"
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimirDoc 
      Caption         =   "Imprimir Documento Autorizado"
      Height          =   615
      Left            =   2400
      TabIndex        =   23
      Top             =   8760
      Width           =   6135
   End
   Begin VB.TextBox txtdhRecbto 
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox txtnProt 
      Height          =   315
      Left            =   6360
      TabIndex        =   19
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox txtStatusSefaz 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox txtMotivoSefaz 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   8160
      Width           =   6135
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtMotivo 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   6960
      Width           =   7215
   End
   Begin VB.TextBox txtChaveRetorno 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   7560
      Width           =   4935
   End
   Begin VB.TextBox txtnsNRec 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8415
   End
   Begin VB.CommandButton cmdConsultarStatus 
      Caption         =   "Verificar Retorno de Processamento do Documento"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   8415
   End
   Begin VB.TextBox txtResult 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Width           =   8415
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8760
      Y1              =   18445.54
      Y2              =   18445.54
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente (1 ou 2)"
      Height          =   195
      Left            =   5880
      TabIndex        =   28
      Top             =   720
      Width           =   1830
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo de Autorização"
      Height          =   195
      Left            =   6360
      TabIndex        =   22
      Top             =   7320
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora de Recebimento"
      Height          =   195
      Left            =   6360
      TabIndex        =   21
      Top             =   7920
      Width           =   2085
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Sefaz"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   7920
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "cStat"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Motivo"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   6720
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso Documento"
      Height          =   195
      Left            =   1320
      TabIndex        =   12
      Top             =   7320
      Width           =   2130
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "nRec do Envio para Sefaz"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ Emitente"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1530
   End
End
Attribute VB_Name = "frmNFeAPIRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultarStatus_Click()
On Error GoTo SAI
    Dim result As String
    result = consultaStatusProcessamento(txtToken.Text, txtCNPJ.Text, txtnsNRec.Text, txtTpAmb.Text)
    txtResult.Text = result
    Dim protocolo As String
    
    
    'lendo status do JSON recebido da API
    txtStatus.Text = LerDadosJSON(result, "status", "", "")
    'lendo motivo do JSON recebido da API
    txtMotivo.Text = LerDadosJSON(result, "motivo", "", "")
    'lendo chave de acesso do JSON recebido da API
    txtChaveRetorno.Text = LerDadosJSON(result, "chNFe", "", "")
    'lendo Data e Hora de Recebimento na Sefaz, retornado no JSON recebido da API
    txtdhRecbto.Text = LerDadosJSON(result, "dhRecbto", "", "")
    'lendo cSat da Sefaz retornado no JSON recebido da API
    txtStatusSefaz.Text = LerDadosJSON(result, "cStat", "", "")
    'lendo xMotivo da Sefaz retornado no JSON recebido da API
    txtMotivoSefaz.Text = LerDadosJSON(result, "xMotivo", "", "")
    'lendo nProt(Protocolo de Autorização) retornado no JSON recebido da API
    protocolo = LerDadosJSON(result, "nProt", "", "")
    If txtStatusSefaz <> "100" Then
        txtnProt.Text = "Não Possui"
    Else
        txtnProt.Text = protocolo
        
        'Lê dados do XML
        Dim xml As String
        xml = LerDadosJSON(result, "xml", "", "")
        
        'Lê CNPJ
        txtCNPJ_CPF_Dest.Text = LerDadosXML(xml, "dest", "CNPJ")
        'Se CNPJ estiver em branco, então lê CPF
        If (txtCNPJ_CPF_Dest.Text = "") Then
            txtCNPJ_CPF_Dest.Text = LerDadosXML(xml, "dest", "CPF")
        End If
        'Lê valor nota
        txtTotalNota.Text = LerDadosXML(xml, "total", "vNF")
        'Lê natureza operação
        txtNatOperacao.Text = LerDadosXML(xml, "ide", "natOp")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub

Private Sub cmdImprimirDoc_Click()
On Error GoTo SAI
    'Requisitando download para a API
    Dim result As String
    Dim isShow As Boolean
    isShow = checkExibir.Value
    
    'lendo o responsetext, que é onde está ou estarão o xml, pdf, JSON conforme o tipo informado
    result = downloadNFeAndSave(txtToken.Text, txtChaveRetorno.Text, comboTpDown.Text, txtTpAmb.Text, "C:\Documentos", isShow)
    'result = downloadNFe(txtToken.Text, txtChaveRetorno.Text, comboTpDown.Text)
    txtResult.Text = result
    
    Exit Sub
SAI:
    txtResult.Text = ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description)
End Sub
