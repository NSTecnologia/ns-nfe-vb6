VERSION 5.00
Begin VB.Form frmEnviaEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar Email"
   ClientHeight    =   4440
   ClientLeft      =   6510
   ClientTop       =   2565
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6690
   Begin VB.CommandButton cmdEnviarEmail 
      Caption         =   "Enviar!"
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtEmails 
      Height          =   1335
      Left            =   720
      TabIndex        =   3
      Text            =   "exemplo@exemplo.com, exemplo1@exemplo1.com"
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CheckBox cbEnviaOriginal 
      Caption         =   "cbEnviaEmailDoc"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   255
   End
   Begin VB.TextBox txtChaveDocEmail 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "ENVIAR EMAIL PARA DOCUMENTO"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Enviar para e-maisl originais do documento (emails do destinataio)"
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Destinatario(separados por virgula)"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Chave da NFe:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmEnviaEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviarEmail_Click()
    On Error GoTo SAI
    Dim retorno As String
    Dim enviaParaOriginal As String
    
    If (txtChaveDocEmail.Text <> "") And (txtEmails.Text <> "") Then
    
        If (cbEnviaOriginal.Value) Then
            enviaParaOriginal = "true"
        Else
            enviaParaOriginal = "false"
        End If
        
        retorno = NFeAPI.enviarEmail(txtChaveDocEmail.Text, enviaParaOriginal, txtEmails.Text)
              
        Dim status, motivo, cStat, erro As String
        status = LerDadosJSON(retorno, "status", "", "")
        If (status <> 200) Then
            If (status <> -2) Then
                If (status = -3) Then
                    MsgBox (retorno)
                End If
            Else
                MsgBox ("enviaEmailDoc deve ser true ou os enderecos de destino devem ser informados no campo email")
            End If
        Else
            motivo = LerDadosJSON(retorno, "motivo", "", "")
            MsgBox (motivo)
        End If
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description)
End Sub

