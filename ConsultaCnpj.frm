VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInscricaoEstadual 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   6000
      Width           =   4935
   End
   Begin VB.TextBox txtTelefone 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox txtComplemento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   5160
      Width           =   5175
   End
   Begin VB.TextBox txtDdd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtBairro 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtMunicipio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox txtCodigoIbge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtUf 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCep 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtLogradouro 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3480
      Width           =   5175
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtCnpj 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton btn_get_cnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar CNPJ"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o CNPJ abaixo para consultar os dados"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2640
      TabIndex        =   29
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Inscrição estadual"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7320
      TabIndex        =   27
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Telefone"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DDD"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Município"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Código IBGE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "UF"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CEP"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Logradouro"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblCnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CNPJ"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_get_cnpj_Click()
   
   If txtCnpj.Text = "" Or txtCnpj.Text = "0" Then
        MsgBox "Campo CPF/CNPJ Obrigatório!", vbCritical
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    Dim Documento As String
    Documento = Replace(Replace(Replace(txtCnpj.Text, ".", ""), "-", ""), "/", "")
    
    If Len(Documento) = 11 Then
        MsgBox "Não é possível consultar CPF, apenas CNPJ.", vbCritical, "Alerta"
        txtCnpj.SetFocus
        Exit Sub
    End If

    Dim Cnpj As String
    Dim DadosCnpj As T_DadosCnpjResponse
        
    Cnpj = txtCnpj.Text
    
    DadosCnpj = PreencherDadosCnpj(Cnpj)
    
    txtNome.Text = DadosCnpj.Nome
    txtCodigoIbge.Text = DadosCnpj.CodigoIbge
    txtCep.Text = DadosCnpj.Cep
    txtLogradouro.Text = DadosCnpj.Logradouro
    txtNumero.Text = DadosCnpj.Numero
    txtComplemento.Text = DadosCnpj.Complemento
    txtBairro.Text = DadosCnpj.Bairro
    txtMunicipio.Text = DadosCnpj.Cidade
    txtUf.Text = DadosCnpj.Uf
    txtInscricaoEstadual.Text = DadosCnpj.InscricaoEstadual
    txtDdd.Text = DadosCnpj.Ddd
    txtTelefone.Text = DadosCnpj.Telefone
    txtEmail.Text = DadosCnpj.Email
    
End Sub
