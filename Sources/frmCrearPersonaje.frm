VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCaptcha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   9480
      MaxLength       =   4
      TabIndex        =   35
      Top             =   7560
      Width           =   1935
   End
   Begin VB.PictureBox pCaptcha 
      Appearance      =   0  'Flat
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9480
      ScaleHeight     =   345
      ScaleWidth      =   1905
      TabIndex        =   34
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   7080
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   31
      Top             =   6330
      Width           =   615
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   990
      Left            =   7800
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   30
      Top             =   6330
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   6150
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Timer tAnimacion 
      Left            =   840
      Top             =   1080
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":3FC3E
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":3FC40
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4035
      Width           =   2625
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":3FC42
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":3FC4C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4560
      Width           =   2625
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":3FC5F
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":3FC61
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3495
      Width           =   2625
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":3FC63
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":3FC65
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2985
      Width           =   2625
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1320
      Width           =   4815
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   6795
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   7200
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   7605
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   8010
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   6390
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin ClientGSZAO.uAOButton cVolver 
      Height          =   495
      Left            =   2880
      TabIndex        =   32
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      TX              =   "Volver"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearPersonaje.frx":3FC67
      PICF            =   "frmCrearPersonaje.frx":3FC83
      PICH            =   "frmCrearPersonaje.frx":3FC9F
      PICV            =   "frmCrearPersonaje.frx":3FCBB
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ClientGSZAO.uAOButton cCrear 
      Height          =   495
      Left            =   4800
      TabIndex        =   33
      Top             =   8160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "Crear Personaje"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearPersonaje.frx":3FCD7
      PICF            =   "frmCrearPersonaje.frx":3FCF3
      PICH            =   "frmCrearPersonaje.frx":3FD0F
      PICV            =   "frmCrearPersonaje.frx":3FD2B
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "C�digo de confirmaci�n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   37
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lCodigo 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el c�digo aqu�:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   36
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   5685
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   29
      Top             =   7395
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   479
      X2              =   505
      Y1              =   417
      Y2              =   417
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   479
      X2              =   505
      Y1              =   391
      Y2              =   391
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   505
      X2              =   505
      Y1              =   392
      Y2              =   416
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   479
      X2              =   479
      Y1              =   392
      Y2              =   416
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5445
      TabIndex        =   23
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5445
      TabIndex        =   22
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5445
      TabIndex        =   21
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5445
      TabIndex        =   20
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5445
      TabIndex        =   19
      Top             =   3450
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4950
      TabIndex        =   18
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4950
      TabIndex        =   17
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4950
      TabIndex        =   16
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4950
      TabIndex        =   15
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4950
      TabIndex        =   14
      Top             =   3450
      Width           =   225
   End
   Begin VB.Image imgAtributos 
      Height          =   270
      Left            =   3840
      Top             =   2745
      Width           =   1245
   End
   Begin VB.Label lblHelp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   9480
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   6720
      Top             =   4335
      Width           =   1245
   End
   Begin VB.Image imgClase 
      Height          =   255
      Left            =   6720
      Top             =   3795
      Width           =   1245
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   6720
      Top             =   3270
      Width           =   1245
   End
   Begin VB.Image imgPuebloOrigen 
      Height          =   225
      Left            =   6360
      Top             =   2760
      Width           =   2025
   End
   Begin VB.Image imgEspecialidad 
      Height          =   240
      Left            =   3120
      Top             =   7410
      Width           =   1245
   End
   Begin VB.Image imgArcos 
      Height          =   225
      Left            =   3120
      Top             =   7140
      Width           =   1245
   End
   Begin VB.Image imgArmas 
      Height          =   240
      Left            =   3120
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Image imgEscudos 
      Height          =   255
      Left            =   3120
      Top             =   6540
      Width           =   1245
   End
   Begin VB.Image imgVida 
      Height          =   225
      Left            =   3120
      Top             =   6270
      Width           =   1245
   End
   Begin VB.Image imgMagia 
      Height          =   255
      Left            =   3120
      Top             =   5955
      Width           =   1245
   End
   Begin VB.Image imgEvasion 
      Height          =   255
      Left            =   3120
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   3120
      Top             =   4785
      Width           =   1245
   End
   Begin VB.Image imgCarisma 
      Height          =   240
      Left            =   3120
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Image imgInteligencia 
      Height          =   240
      Left            =   3120
      Top             =   4110
      Width           =   1245
   End
   Begin VB.Image imgAgilidad 
      Height          =   240
      Left            =   3120
      Top             =   3750
      Width           =   1245
   End
   Begin VB.Image imgFuerza 
      Height          =   240
      Left            =   3120
      Top             =   3420
      Width           =   1245
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   5415
      Top             =   3075
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   4950
      Top             =   3075
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   4485
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgTirarDados 
      Height          =   765
      Left            =   1200
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Image DirPJ 
      Height          =   300
      Index           =   1
      Left            =   7380
      Top             =   7275
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image DirPJ 
      Height          =   300
      Index           =   0
      Left            =   7080
      Top             =   7275
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image HeadPJ 
      Height          =   345
      Index           =   1
      Left            =   8460
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   345
      Index           =   0
      Left            =   6075
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgDados 
      Height          =   1245
      Left            =   0
      MouseIcon       =   "frmCrearPersonaje.frx":3FD47
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   5640
      Picture         =   "frmCrearPersonaje.frx":3FE99
      Top             =   9120
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4500
      TabIndex        =   12
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4500
      TabIndex        =   11
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4500
      TabIndex        =   10
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4500
      TabIndex        =   9
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4500
      TabIndex        =   8
      Top             =   3450
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieTirarDados
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieD
    ieM
    ieF
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
End Enum

Private vHelp(24) As String
Private vEspecialidades() As String
Private MacroDados As Boolean

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    Da�oArmas As Double
    Da�oProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza() As tModRaza
Private ModClase() As tModClase

Private NroRazas As Integer
Private NroClases As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private Sub cCrear_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    Dim i As Integer
    Dim CharAscii As Byte
    UserName = txtNombre.Text
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
        txtNombre.Text = UserName
        MsgBox "Nombre invalido, se han removido los espacios al final del nombre."
        Exit Sub
    End If
    
    If Left(UserName, 1) = " " Then
        UserName = LTrim$(UserName)
        txtNombre.Text = UserName
        MsgBox "Nombre invalido, se han removido los espacios al comienzo del nombre."
        Exit Sub
    End If
    
    If LCase$(txtCaptcha.Text) <> (Chr$(CaptchaCode(0)) & Chr$(CaptchaCode(1)) & Chr$(CaptchaCode(2)) & Chr$(CaptchaCode(3))) Then
        MsgBox "Los c�digos de confirmaci�n no coinciden. Vuelva a ingresarlo."
        txtCaptcha.Text = vbNullString
        txtCaptcha.SetFocus
        Exit Sub
    End If
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i
    
    UserHogar = lstHogar.ListIndex + 1
    
    If Not CheckData Then Exit Sub

    Dim eMD5 As New clsMD5
    UserPassword = eMD5.DigestStrToHexStr(txtPasswd.Text) ' GSZ
    
    UserEmail = txtMail.Text
    
    frmMain.Socket1.HostAddress = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
    If Not frmMain.Socket1.Connected Then
        MsgBox "ERROR" & vbCrLf & "Se ha perdido la conexi�n con el servidor.", vbCritical
        frmConnect.Visible = True
        Unload Me
    Else
        Call Login
    End If
    
    bShowTutorial = TutorialDefualt
End Sub

Private Sub cVolver_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call Audio.PlayMIDI("2.mid")
    bShowTutorial = False
    frmConnect.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
    frmConnect.Visible = False
    
    Cargando = True
    MacroDados = False
    
    'UserClase = 0
    UserSexo = eGenero.hombre
    UserRaza = eRaza.Humano
    UserHogar = 1
    UserEmail = vbNullString
    UserHead = 0
    Dir = SOUTH
    
    Me.Picture = LoadPicture(pathGUI & "frmCrearPersonaje.jpg")
    
    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(pathButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(pathButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(pathButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(pathButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(pathButtons & sty_cCheckbox2))
        End If
    Next
    
    Call LoadCharInfo
    Call CargarEspecialidades
    
    Call IniciarGraficos
    Call CargarCombos
    
    Call LoadHelp
    
    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    Call DarCuerpoYCabeza
        
    Cargando = False
End Sub

Private Sub CargarEspecialidades()

    ReDim vEspecialidades(1 To NroClases)
    
    vEspecialidades(eClass.Hunter) = "Ocultarse"
    vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
    vEspecialidades(eClass.Assasin) = "Apu�alar"
    vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
    vEspecialidades(eClass.Druid) = "Domar"
    vEspecialidades(eClass.Pirat) = "Navegar"
    vEspecialidades(eClass.Worker) = "Extracci�n y Construcci�n"
End Sub

Private Sub IniciarGraficos()
    
    Set picFullStar = LoadPicture(pathButtons & "SetStars_EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(pathButtons & "SetStars_EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(pathButtons & "SetStars_EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
    Dim i As Integer
    
    lstProfesion.Clear
    
    For i = LBound(ListaClases) To NroClases
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    For i = LBound(Ciudades()) To UBound(Ciudades())
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.ListIndex = 1
    lstHogar.ListIndex = 0
    lstRaza.ListIndex = 0
    lstGenero.ListIndex = 0
End Sub

Function CheckData() As Boolean
'Last Modification: 05/11/2011 - ^[GS]^
'**************************************

    If txtPasswd.Text <> txtConfirmPasswd.Text Then
        MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
        Exit Function
    End If
    
    If Not CheckMailString(txtMail.Text) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If
    
    If UserSexo = 0 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If
    
    If UserClase = 0 Then
        MsgBox "Seleccione la clase del personaje."
        Exit Function
    End If
    
    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If Val(lblAtributos(i).Caption) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    CheckData = True

End Function


Private Sub DirPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select
    
    Call UpdateHeadSelection
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearLabel
    MacroDados = False
End Sub

Private Sub Form_Resize()
    Call UpdateCaptcha
End Sub

Private Sub HeadPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            UserHead = CheckCabeza(UserHead + 1)
        Case 1
            UserHead = CheckCabeza(UserHead - 1)
    End Select
    
    Call UpdateHeadSelection
    
End Sub

Private Sub UpdateHeadSelection()
    Dim Head As Integer
    
    Head = UserHead
    Call DrawHead(Head, 2)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 3)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 4)
    
    Head = UserHead
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 1)
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 0)
End Sub

Private Sub imgDados_Click()
    If MacroDados = False Then
        Call Audio.PlayWave(SND_DICE)
        Call TirarDados
        MacroDados = True
    End If
End Sub

Private Sub imgDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MacroDados = False
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgTirarDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieTirarDados)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
    Call DarCuerpoYCabeza
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
'    Image1.Picture = LoadPicture(pathGraphics & lstProfesion.Text & ".jpg")
'
    UserClase = lstProfesion.ListIndex + 1
    
    Call UpdateStats
    Call UpdateEspecialidad(UserClase)
End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
    lblEspecialidad.Caption = vEspecialidades(eClase)
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza
    
    Call UpdateStats
End Sub

Private Sub pCaptcha_Click()
    Call UpdateCaptcha
End Sub

Private Sub pCaptcha_Paint()
    Call UpdateCaptcha
End Sub

Private Sub picHead_Click(Index As Integer)
    ' No se mueve si clickea al medio
    If Index = 2 Then Exit Sub
    
    Dim Counter As Integer
    Dim Head As Integer
    
    Head = UserHead
    
    If Index > 2 Then
        For Counter = Index - 2 To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter
    Else
        For Counter = 2 - Index To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter
    End If
    
    UserHead = Head
    
    Call UpdateHeadSelection
    
End Sub

Private Sub tAnimacion_Timer()
On Error Resume Next
    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long
    Static Frame As Byte

    If currentGrh = 0 Then Exit Sub
    UserHead = CheckCabeza(UserHead)
    
    Frame = Frame + 1
    If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1
    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    
    Grh = GrhData(currentGrh).Frames(Frame)
    
    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = ((picPJ.Height - .pixelHeight) \ 2) + 5
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = DR.Bottom + BodyData(UserBody).HeadOffset.Y - .pixelHeight
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long

    Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picHead(0).Width - .pixelWidth) \ 2 + 1
        DR.Top = 5
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
End Sub

Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim$(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
    Select Case UserSexo
        Case eGenero.hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
    bVisible = UserHead <> 0 And UserBody <> 0
    
    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible
    
    For PicIndex = 0 To 4
        picHead(PicIndex).Visible = bVisible
    Next PicIndex
    
    For LineIndex = 0 To 3
        Line1(LineIndex).Visible = bVisible
    Next LineIndex
    
    If bVisible Then Call UpdateHeadSelection
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

Select Case UserSexo
    Case eGenero.hombre
        Select Case UserRaza
            Case eRaza.Humano
                If Head > HUMANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_H_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_H_PRIMER_CABEZA Then
                    CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_H_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
        
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                If Head > HUMANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_M_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_M_PRIMER_CABEZA Then
                    CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_M_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
    Case Else
        UserSexo = lstGenero.ListIndex + 1
        CheckCabeza = CheckCabeza(Head)
End Select
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
    vHelp(eHelp.iePasswd) = "La contrase�a que utilizar�s para conectar tu personaje al juego."
    vHelp(eHelp.ieTirarDados) = "Presionando sobre la Esfera Roja, se modificar�n al azar los atributos de tu personaje, de esta manera puedes elegir los que m�s te parezcan para definir a tu personaje."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una direcci�n de correo electr�nico v�lida, ya que en el caso de perder la contrase�a de tu personaje, se te enviar� cuando lo requieras, a esa direcci�n."
    vHelp(eHelp.ieNombre) = "S� cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo m�gico y fant�stico, y si seleccion�s un nombre obsceno o con connotaci�n pol�tica, los administradores borrar�n tu personaje y no habr� ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contrase�a que utilizar�s para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
    vHelp(eHelp.ieD) = "Son los atributos que obtuviste al azar. Presion� la esfera roja para volver a tirarlos."
    vHelp(eHelp.ieM) = "Son los modificadores por raza que influyen en los atributos de tu personaje."
    vHelp(eHelp.ieF) = "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
    vHelp(eHelp.ieFuerza) = "De ella depender� qu� tan potentes ser�n tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendr� en qu� tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influir� de manera directa en cu�nto man� ganar�s por nivel."
    vHelp(eHelp.ieCarisma) = "Ser� necesario tanto para la relaci�n con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectar� a la cantidad de vida que podr�s ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Eval�a la habilidad esquivando ataques f�sicos."
    vHelp(eHelp.ieMagia) = "Punt�a la cantidad de man� que se tendr�."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podr� llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Eval�a la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Eval�a la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = vbNullString
    vHelp(eHelp.iePuebloOrigen) = "Define el hogar de tu personaje. Sin embargo, el personaje nacer� en Nemahuak, la ciudad de los novatos."
    vHelp(eHelp.ieRaza) = "De la raza que elijas depender� c�mo se modifiquen los dados que saques. Pod�s cambiar de raza para poder visualizar c�mo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influir� en las caracter�sticas principales que tenga tu personaje, asi como en las magias e items que podr� utilizar. Las estrellas que ves abajo te mostrar�n en qu� habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje ser� masculino o femenino. Esto influye en los items que podr� equipar."
End Sub

Private Sub ClearLabel()
    lblHelp = vbNullString
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()
    
    Call UpdateRazaMod
    Call UpdateStars
End Sub

Private Sub UpdateRazaMod()
    Dim SelRaza As Integer
    Dim i As Integer
    
    If lstRaza.ListIndex > -1 Then
    
        SelRaza = lstRaza.ListIndex + 1
        
        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", "") & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", "") & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion
        End With
    End If
    
    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i
    
End Sub

' GSZAO
Public Sub UpdateCaptcha()
    pCaptcha.Cls
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 10, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 20, pCaptcha.ScaleWidth), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(300, 400)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(0))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-60, 100)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(2))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-100, -200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(3))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(150, 200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(150, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(1))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    'Set picCaptcha = pCaptcha.Image
    'pCaptcha.PaintPicture picCaptcha, 0, 0
End Sub

Private Sub UpdateStars()
    Dim NumStars As Double
    
    If UserClase = 0 Then Exit Sub
    
    ' Estrellas de evasion
    NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(UserClase).Evasion
    Call SetStars(imgEvasionStar, NumStars * 2)
    
    ' Estrellas de magia
    NumStars = ModClase(UserClase).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
    Call SetStars(imgMagiaStar, NumStars * 2)
    
    ' Estrellas de vida
    NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(UserClase).Vida) * 0.475
    Call SetStars(imgVidaStar, NumStars * 2)
    
    ' Estrellas de escudo
    NumStars = 4 * ModClase(UserClase).Escudo
    Call SetStars(imgEscudosStar, NumStars * 2)
    
    ' Estrellas de armas
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * _
                ModClase(UserClase).Da�oArmas + 0.119 * ModClase(UserClase).AtaqueArmas * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)
    
    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
                ModClase(UserClase).Da�oProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
    Dim FullStars As Integer
    Dim HasHalfStar As Boolean
    Dim Index As Integer
    Dim Counter As Integer

    If NumStars > 0 Then
        
        If NumStars > 10 Then NumStars = 10
        
        FullStars = Int(NumStars / 2)
        
        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then
            For Index = 1 To FullStars
                ImgContainer(Index).Picture = picGlowStar
            Next Index
        Else
            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
            ' Muestro las estrellas enteras
            If FullStars > 0 Then
                For Index = 1 To FullStars
                    ImgContainer(Index).Picture = picFullStar
                Next Index
                
                Counter = FullStars
            End If
            
            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1
                
                ImgContainer(Counter).Picture = picHalfStar
            End If
            
            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                ' Limpio las que queden vacias
                For Index = Counter + 1 To 5
                    Set ImgContainer(Index).Picture = Nothing
                Next Index
            End If
            
        End If
    Else
        ' Limpio todo
        For Index = 1 To 5
            Set ImgContainer(Index).Picture = Nothing
        Next Index
    End If

End Sub

Private Sub LoadCharInfo()
    Dim SearchVar As String
    Dim i As Integer
    
    NroRazas = UBound(ListaRazas())
    NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    ReDim ModClase(1 To NroClases)
    
    'Modificadores de Clase
    For i = 1 To NroClases
        With ModClase(i)
            SearchVar = stringSinTildes(ListaClases(i))
            
            .Evasion = Val(GetVar(pathInits & "CharInfo.dat", "MODEVASION", SearchVar))
            .AtaqueArmas = Val(GetVar(pathInits & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = Val(GetVar(pathInits & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
            .Da�oArmas = Val(GetVar(pathInits & "CharInfo.dat", "MODDA�OARMAS", SearchVar))
            .Da�oProyectiles = Val(GetVar(pathInits & "CharInfo.dat", "MODDA�OPROYECTILES", SearchVar))
            .Escudo = Val(GetVar(pathInits & "CharInfo.dat", "MODESCUDO", SearchVar))
            .Hit = Val(GetVar(pathInits & "CharInfo.dat", "HIT", SearchVar))
            .Magia = Val(GetVar(pathInits & "CharInfo.dat", "MODMAGIA", SearchVar))
            .Vida = Val(GetVar(pathInits & "CharInfo.dat", "MODVIDA", SearchVar))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NroRazas
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", "")
        
            .Fuerza = Val(GetVar(pathInits & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(pathInits & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(pathInits & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(pathInits & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(pathInits & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i

End Sub
