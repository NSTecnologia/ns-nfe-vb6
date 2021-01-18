VERSION 5.00
Begin VB.Form frmCadLic 
   Caption         =   "Cadastro de Licenca para DF-es"
   ClientHeight    =   8535
   ClientLeft      =   5085
   ClientTop       =   1110
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10425
   Begin VB.Frame frameLicenca 
      Caption         =   "Licença"
      Height          =   2895
      Left            =   360
      TabIndex        =   18
      Top             =   4920
      Width           =   9735
   End
   Begin VB.Frame frameEndereco 
      Caption         =   "Endereço"
      Height          =   1815
      Left            =   4680
      TabIndex        =   17
      Top             =   3000
      Width           =   5415
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3000
         TabIndex        =   39
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4200
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "cIBGE"
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "CEP"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Numero"
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Rua"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frameServidorEmail 
      Caption         =   "Servidor de Email"
      Height          =   2175
      Left            =   4680
      TabIndex        =   16
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cbSSL 
         Height          =   315
         ItemData        =   "frmCadLic.frx":0000
         Left            =   3720
         List            =   "frmCadLic.frx":000A
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox ckConfLeitura 
         Caption         =   "Confirmar Leitura"
         Height          =   315
         Left            =   3720
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Senha"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "SSL"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Porta"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Servidor SMTP"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSalvarLicenca 
      Caption         =   "Adicionar Licenca"
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame framePessoa 
      Caption         =   "Pessoa"
      Height          =   4095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtTelefones 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtEmails 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cbTipoICMS 
         Height          =   315
         Index           =   0
         ItemData        =   "frmCadLic.frx":0020
         Left            =   1200
         List            =   "frmCadLic.frx":002D
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtIE 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFantasiaCadLic 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtRazaoCadLic 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCNPJCadLic 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Telefones"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Emails"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label label7 
         Caption         =   "Fantasia"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de ICMS"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "IE"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lLabel3 
         Caption         =   "Razao Social"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "CNPJ/CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCadLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSalvarLicenca_Click()
    Dim retorno As String
    Dim enviaParaOriginal As String
    Dim base As String
    Dim fileNum As Integer
    Dim bytes() As Byte

    fileNum = FreeFile
    Open "D:\certTeste.pfx" For Binary As fileNum
    ReDim bytes(LOF(fileNum) - 1)
    Get fileNum, , bytes
    Close fileNum
    
    base = Base64Encode(bytes)
    ' cIBGE, situacao, idprojeto, usarAssinaturaLocal, certificado, senhaCert As String

    retorno = NFeAPI.cadastrarLicenca("27187851000141", "V A DE LIMA ALIMENTOS - ME", "LIMA DISTRIBUIDORA", "225276368118", "1", "lima.distribuidora.2017@gmail.com", "R CARVALHO MOTTA", "663", "VILA MOTA", "12903170", "3507605", "11944738654", "0", "1", "false", base, "v1202a65l", "smtp.gmail.com", "587", "1", "1", "matheusdiasmazzoni@gmail.com", "mazzoni123")

    
End Sub
