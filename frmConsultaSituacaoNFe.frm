VERSION 5.00
Begin VB.Form frmConsultaSituacaoNFe 
   Caption         =   "Consultar Situação de NFe"
   ClientHeight    =   3195
   ClientLeft      =   7890
   ClientTop       =   2850
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4200
   Begin VB.ComboBox cbTpAmb 
      Height          =   315
      ItemData        =   "frmConsultaSituacaoNFe.frx":0000
      Left            =   600
      List            =   "frmConsultaSituacaoNFe.frx":000A
      TabIndex        =   5
      Text            =   "2"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados NFe"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txtchNFe 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "35200333642842000104550010000000401027735546"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "chNFe"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdConsSit 
      Caption         =   "Consultar Situção"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Licença"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtLicencaCNPJ 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "33642842000104"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "CNPJ"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Ambiente"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmConsultaSituacaoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsSit_Click()
    Dim resposta As String
    Dim status As String
        Dim cStat As String
    Dim auxInfProt As Variant
    Dim respInfProt As String
    Dim infProt As String
    Dim chNFe As String
    Dim nProt As String
    Dim dhEmissao As String
    
    If (frmConsultaSituacaoNFe.txtLicencaCNPJ.Text <> "") And (frmConsultaSituacaoNFe.txtchNFe.Text <> "") And (frmConsultaSituacaoNFe.cbTpAmb.Text <> "") Then
        resposta = consultarSituacao(frmConsultaSituacaoNFe.txtLicencaCNPJ.Text, frmConsultaSituacaoNFe.txtchNFe.Text, frmConsultaSituacaoNFe.cbTpAmb.Text)
    Else
        MsgBox ("Todos campos necessarios devem ser preenchidos")
    End If
       
    status = LerDadosJSON(resposta, "status", "", "")
    If (status = 200) Then
        cStat = LerDadosJSON(resposta, "retConsSitNFe", "cStat", "")
        If (cStat = 101 Or cStat = 100) Then
            auxInfProt = Split(resposta, """protNFe"":[")
            respInfProt = auxInfProt(1)
            auxInfProt = Split(respInfProt, "]")
            infProt = auxInfProt(0)
            chNFe = LerDadosJSON(infProt, "infProt", "chNFe", "")
            nProt = LerDadosJSON(infProt, "infProt", "nProt", "")
            dhEmissao = LerDadosJSON(infProt, "infProt", "dhRecbto", "")
            Dim msg As String
            msg = "Chave da Nota: " & chNFe & vbCrLf & "nProt:" & nProt & vbCrLf & "Data de Emissão:" & dhEmissao
            MsgBox (msg)
        End If
    End If
End Sub
