VERSION 5.00
Begin VB.Form frmDV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo do DV"
   ClientHeight    =   4995
   ClientLeft      =   10530
   ClientTop       =   4620
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   2310
   Begin VB.TextBox txtRIC 
      Height          =   345
      Left            =   600
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "1234567890"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalcularRIC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Calcular DV do RIC"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   4560
      Width           =   1995
   End
   Begin VB.CommandButton cmdCalcularRG 
      Caption         =   "&Calcular DV do RG"
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   540
      Width           =   1695
   End
   Begin VB.TextBox txtRG 
      Height          =   345
      Left            =   720
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "42943412"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblDUVIDA 
      Caption         =   "Dúvida! Como tratamos os restos ""0"", ""1"" e ""10""? Exemplo: No RG, resto ""10"" é dígito ""X""."
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   180
      TabIndex        =   9
      Top             =   3180
      Width           =   1935
   End
   Begin VB.Label lblATENCAO 
      Caption         =   $"frmDV.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   180
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblRIC 
      Caption         =   "RIC:"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblDVRIC 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   6
      Top             =   4140
      Width           =   315
   End
   Begin VB.Label lblDVRG 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblRG 
      Caption         =   "RG:"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
End
Attribute VB_Name = "frmDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcularRG_Click()
    lblDVRG.Caption = Trim(mdlCalculo.CalculoDV11(txtRG.Text))
End Sub

Private Sub cmdCalcularRIC_Click()
    lblDVRIC.Caption = Trim(mdlCalculo.Mod_dig11(txtRIC.Text))
End Sub
