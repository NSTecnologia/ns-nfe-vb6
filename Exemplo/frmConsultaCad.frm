VERSION 5.00
Begin VB.Form frmConsultaCad 
   Caption         =   "frmConsultaCad"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRetornoCNPJ 
      Height          =   375
      Left            =   1050
      TabIndex        =   16
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtRetornoIE 
      Height          =   405
      Left            =   3735
      TabIndex        =   15
      Top             =   2745
      Width           =   1935
   End
   Begin VB.TextBox txtCNPJ_CPF 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtUF 
      Height          =   375
      Left            =   105
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCNPJCont 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   345
      Width           =   2535
   End
   Begin VB.ComboBox cbTtipoContrib 
      Height          =   315
      ItemData        =   "frmConsultaCad.frx":0000
      Left            =   1440
      List            =   "frmConsultaCad.frx":000D
      TabIndex        =   11
      Text            =   "CNPJ"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtRetornoCep 
      Height          =   375
      Left            =   1020
      TabIndex        =   4
      Top             =   4155
      Width           =   1575
   End
   Begin VB.TextBox txtRetornoXLgr 
      Height          =   375
      Left            =   1035
      TabIndex        =   3
      Top             =   3690
      Width           =   4695
   End
   Begin VB.TextBox txtRetornoNome 
      Height          =   375
      Left            =   1035
      TabIndex        =   2
      Top             =   3225
      Width           =   4695
   End
   Begin VB.CommandButton cmdConsultarCad 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.E."
      Height          =   255
      Left            =   3135
      TabIndex        =   18
      Top             =   2835
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNPJ a Consultar"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   885
      Width           =   1305
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Razão"
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   3300
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CEP"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Endereço"
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   3735
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNPJ"
      Height          =   255
      Left            =   165
      TabIndex        =   7
      Top             =   2820
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retorno Consulta"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2415
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UF"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNPJ"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmConsultaCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultarCad_Click()
    On Error GoTo SAI
    Dim retorno As String
     
    Dim status As String
    Dim restContCad As String
    Dim infCons As String
    Dim auxInfCad As Variant
    Dim respInfCad As String
    Dim infCad As String
    
    Dim ie As String
    Dim cnpj As String
    Dim xNome As String
    Dim xLgr As String
    Dim nro As String
    Dim xCpl As String
    Dim xBairro As String
    Dim cMun As String
    Dim CEP As String
            
    retorno = consultarCadastroContribuinte(txtCNPJCont.Text, txtUF.Text, txtCNPJ_CPF.Text, cbTtipoContrib.Text)
    
    status = LerDadosJSON(retorno, "status", "", "")
    
    If (status = 200) Then
        cStat = LerDadosJSON(retorno, "retConsCad", "infCons", "cStat")

        If (cStat = "111") Or (cStat = "112") Then
            motivo = LerDadosJSON(retorno, "motivo", "", "")
            MsgBox (motivo)
            auxInfCad = Split(retorno, """infCad"":[")
            auxInfCad = Split(auxInfCad(1), "]")
            auxInfCad = Split(auxInfCad(0), "},")

            If (UBound(auxInfCad) = 0) Then
                infCad = auxInfCad(0)

                ie = LerDadosJSON(infCad, "IE", "", "")
                cnpj = LerDadosJSON(infCad, "CNPJ", "", "")
                UF = LerDadosJSON(infCad, "UF", "", "")
                xNome = LerDadosJSON(infCad, "xNome", "", "")
                xLgr = LerDadosJSON(infCad, "ender", "xLgr", "")
                CEP = LerDadosJSON(infCad, "ender", "CEP", "")
            
                txtRetornoCNPJ.Text = cnpj
                txtRetornoIE.Text = ie
                txtRetornoNome.Text = xNome
                txtRetornoXLgr.Text = xLgr
                txtRetornoCep.Text = CEP

            Else
                Dim i As Integer
                For i = 0 To UBound(auxInfCad)
                    infCad = auxInfCad(i)

                    If (i <> UBound(auxInfCad)) Then
                        infCad = infCad & "}"
                    End If

                    ie = LerDadosJSON(infCad, "IE", "", "")
                    cnpj = LerDadosJSON(infCad, "CNPJ", "", "")
                    UF = LerDadosJSON(infCad, "UF", "", "")
                    xNome = LerDadosJSON(infCad, "xNome", "", "")
                    xLgr = LerDadosJSON(infCad, "ender", "xLgr", "")
                    CEP = LerDadosJSON(infCad, "ender", "CEP", "")
                Next
            End If
        End If
        
        If (cStat <> "111") Then
            xMotivo = LerDadosJSON(retorno, "retConsCad", "infCons", "xMotivo")
            MsgBox (xMotivo)
        End If

        
    End If
     
    If (status <> 200) Then
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        MsgBox (xMotivo)
    End If
    
Exit Sub
    
SAI:
    MsgBox (vbNewLine & Err.Description), vbInformation, titleNFeAPI
End Sub
