VERSION 5.00
Begin VB.Form frmConnectToken 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnectToken.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   221
   ScaleMode       =   0  'User
   ScaleWidth      =   387
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtToken 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin ClientGSZAO.uAOButton cCerrar 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      tx              =   "Cerrar"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnectToken.frx":030A
      picf            =   "frmConnectToken.frx":0326
      pich            =   "frmConnectToken.frx":0342
      picv            =   "frmConnectToken.frx":035E
      font            =   "frmConnectToken.frx":037A
   End
   Begin ClientGSZAO.uAOButton cAcceder 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      tx              =   "Acceder"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnectToken.frx":039E
      picf            =   "frmConnectToken.frx":03BA
      pich            =   "frmConnectToken.frx":03D6
      picv            =   "frmConnectToken.frx":03F2
      font            =   "frmConnectToken.frx":040E
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce el Token"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   570
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4530
   End
End
Attribute VB_Name = "frmConnectToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAcceder_Click()
    If Not ValidJWT(txtToken.Text) Then
        MsgBox "El token está incompleto, prueba volverlo a ingresar.", vbInformation + vbOKOnly
        Exit Sub
    End If
    ClientConfigInit.Token = txtToken.Text
    Call modGameIni.SaveConfigInit
    Call frmConnect.TryConnectToken
    Unload Me
End Sub

Private Sub cCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    
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
End Sub
