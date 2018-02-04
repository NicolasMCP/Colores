VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Códigos de Cores"
   ClientHeight    =   2535
   ClientLeft      =   1695
   ClientTop       =   1935
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4020
   Begin VB.CommandButton cmdAutor 
      Caption         =   "&Autor"
      Height          =   312
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   312
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEscolhe 
      Caption         =   "&Escolher Cor"
      Height          =   312
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   443
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Left            =   -60
      Top             =   540
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAutor_Click()
    MsgBox "Autor" & vbCrLf & "Herley Nicolas Ramos Sanchez" & vbCrLf & _
            "e-mail: nicolas@infodata.xyz" & vbCrLf & vbCrLf & _
            "Licença GNU GPL (Software Livre)", vbOKOnly, "Autor"
        
End Sub

Private Sub cmdEscolhe_Click()

    '---------------------------------------------------------------------------------------
    ' Software  : Cores
    ' Data Hora : 08/07/2002 13:51
    ' Autor     : Nicolás Ramos
    ' Propósito : Generar o código de uma cor
    ' Sintaxe   : Executar
    '---------------------------------------------------------------------------------------
    '
    CommonDialog1.ShowColor
    Label1.Caption = Hex(CommonDialog1.Color)
    
    Text1.Text = Label1.Caption
    
    Label1.Caption = "#" & Label1.Caption
    
    Shape1.BackColor = "&H" & Text1.Text

End Sub
