VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":0682
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tAccediendo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10200
      Top             =   3360
   End
   Begin VB.Frame fAccediendo 
      BackColor       =   &H001F1B12&
      Height          =   1935
      Left            =   8640
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
      Begin ClientGSZAO.uAOButton cManualToken 
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         TX              =   "Token Manual"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":45744
         PICF            =   "frmConnect.frx":45760
         PICH            =   "frmConnect.frx":4577C
         PICV            =   "frmConnect.frx":45798
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
      Begin VB.Label lblAccediendo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Si tiene problemas, utilice el token manualmente."
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   2340
      End
      Begin VB.Label lblAccediendo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accediendo..."
         ForeColor       =   &H0080FF80&
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Timer tEfectos 
      Left            =   7320
      Top             =   3240
   End
   Begin ClientGSZAO.uAOCheckbox chkRecordar 
      Height          =   195
      Left            =   4560
      TabIndex        =   12
      Top             =   5640
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmConnect.frx":457B4
   End
   Begin ClientGSZAO.uAOButton cSalir 
      Height          =   495
      Left            =   9960
      TabIndex        =   11
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45812
      PICF            =   "frmConnect.frx":4582E
      PICH            =   "frmConnect.frx":4584A
      PICV            =   "frmConnect.frx":45866
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
   Begin ClientGSZAO.uAOButton cCreditos 
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Creditos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45882
      PICF            =   "frmConnect.frx":4589E
      PICH            =   "frmConnect.frx":458BA
      PICV            =   "frmConnect.frx":458D6
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
   Begin ClientGSZAO.uAOButton cSitioOficial 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Sitio Oficial"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":458F2
      PICF            =   "frmConnect.frx":4590E
      PICH            =   "frmConnect.frx":4592A
      PICV            =   "frmConnect.frx":45946
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
   Begin ClientGSZAO.uAOButton cCrearPJ 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   8280
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      TX              =   "Nuevo Personaje"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45962
      PICF            =   "frmConnect.frx":4597E
      PICH            =   "frmConnect.frx":4599A
      PICV            =   "frmConnect.frx":459B6
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
   Begin ClientGSZAO.uAOButton cConectar 
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   6000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Conectarse"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":459D2
      PICF            =   "frmConnect.frx":459EE
      PICH            =   "frmConnect.frx":45A0A
      PICV            =   "frmConnect.frx":45A26
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
   Begin ClientGSZAO.uAOButton cTeclas 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Teclas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45A42
      PICF            =   "frmConnect.frx":45A5E
      PICH            =   "frmConnect.frx":45A7A
      PICV            =   "frmConnect.frx":45A96
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
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   10320
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   9330
      TabIndex        =   3
      Text            =   "7666"
      Top             =   240
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "l"
      TabIndex        =   1
      Text            =   "gs"
      Top             =   5160
      Width           =   2940
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4560
      TabIndex        =   0
      Text            =   "GS"
      Top             =   4380
      Width           =   2940
   End
   Begin ClientGSZAO.uAOButton cOpciones 
      Height          =   495
      Left            =   7920
      TabIndex        =   13
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Opciones"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45AB2
      PICF            =   "frmConnect.frx":45ACE
      PICH            =   "frmConnect.frx":45AEA
      PICV            =   "frmConnect.frx":45B06
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
   Begin ClientGSZAO.uAOButton cAcceder 
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   4200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Acceder"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45B22
      PICF            =   "frmConnect.frx":45B3E
      PICH            =   "frmConnect.frx":45B5A
      PICV            =   "frmConnect.frx":45B76
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
   Begin ClientGSZAO.uAOButton cCerrar 
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Cerrar sesi�n"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":45B92
      PICF            =   "frmConnect.frx":45BAE
      PICH            =   "frmConnect.frx":45BCA
      PICV            =   "frmConnect.frx":45BE6
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
   Begin VB.Label lblAccountName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   60
   End
   Begin VB.Label lRemember 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar contrase�a"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Mat�as Fernando Peque�o
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'C�digo Postal 1405

Option Explicit

' Animaci�n de los Controles...
Private Type tAnimControl
    Activo As Boolean
    Velocidad As Double
    Top As Integer
End Type
Private AnimControl(1 To 7) As tAnimControl
Private Fuerza As Double

Private clsFormulario       As clsFormMovementManager


Private Sub cAcceder_Click()
    'Abrir la web y minimizar el cliente
    Dim r As Long
    r = ShellExecute(0, "open", "https://www.gs-zone.org/login_server.php?code=gszoneao", 0, 0, 1)
    Me.WindowState = vbMinimized
    fAccediendo.Visible = True
    tAccediendo.Enabled = True
    cAcceder.Visible = False
End Sub

Private Sub cCerrar_Click()

    Call CleanToken
    
End Sub

Private Sub cManualToken_Click()
    Dim sToken As String
    sToken = InputBox("Ingresa el Token obtenido en GS-Zone", "Token")
    ClientConfigInit.Token = sToken
    Call modGameIni.SaveConfigInit
End Sub

Private Sub cOpciones_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmOpciones.Show vbModal
End Sub

Private Sub Form_Activate()
    Call Audio.PlayMIDI("0.mid")
End Sub

Private Sub tEfectos_Timer()
    Dim oTop As Integer
    Dim i As Integer
    For i = 1 To 7
        If AnimControl(i).Activo = True Then
            Select Case i
                Case 1: oTop = cTeclas.Top
                Case 2: oTop = cConectar.Top
                Case 3: oTop = cCrearPJ.Top
                Case 4: oTop = cSitioOficial.Top
                Case 5: oTop = cCreditos.Top
                Case 6: oTop = cSalir.Top
                Case 7: oTop = cOpciones.Top
            End Select
            If oTop > AnimControl(i).Top Then
                oTop = AnimControl(i).Top
                AnimControl(i).Velocidad = AnimControl(i).Velocidad * -0.6
            End If
            If AnimControl(i).Velocidad >= -0.6 And AnimControl(i).Velocidad <= -0.5 Then
                AnimControl(i).Activo = False
            Else
                AnimControl(i).Velocidad = AnimControl(i).Velocidad + Fuerza
                oTop = oTop + AnimControl(i).Velocidad
            End If
            Select Case i
                Case 1: cTeclas.Top = oTop
                Case 2: cConectar.Top = oTop
                Case 3: cCrearPJ.Top = oTop
                Case 4: cSitioOficial.Top = oTop
                Case 5: cCreditos.Top = oTop
                Case 6: cSalir.Top = oTop
                Case 7: cOpciones.Top = oTop
            End Select
        End If
    Next
    If AnimControl(1).Activo = False And AnimControl(2).Activo = False And AnimControl(3).Activo = False And _
       AnimControl(4).Activo = False And AnimControl(5).Activo = False And AnimControl(6).Activo = False And _
       AnimControl(7).Activo = False Then
        tEfectos.Enabled = False
        cTeclas.Top = AnimControl(1).Top
        cConectar.Top = AnimControl(2).Top
        cCrearPJ.Top = AnimControl(3).Top
        cSitioOficial.Top = AnimControl(4).Top
        cCreditos.Top = AnimControl(5).Top
        cSalir.Top = AnimControl(6).Top
        cOpciones.Top = AnimControl(7).Top
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
           
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

    CargarServidores ' Cargamos
    
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision

    Me.Picture = LoadPicture(ImgRequest(pathGUI & "frmConnect.jpg"))
    
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
    
    ' GSZAO - Animaci�n...
    cTeclas.Top = 10
    AnimControl(1).Activo = True
    AnimControl(1).Velocidad = 0
    AnimControl(1).Top = 400
    cConectar.Top = 10
    AnimControl(2).Activo = True
    AnimControl(2).Velocidad = 0
    AnimControl(2).Top = 400
    cCrearPJ.Top = 10
    AnimControl(3).Activo = True
    AnimControl(3).Velocidad = 0
    AnimControl(3).Top = 552
    cSitioOficial.Top = 10
    AnimControl(4).Activo = True
    AnimControl(4).Velocidad = 0
    AnimControl(4).Top = 552
    cCreditos.Top = 10
    AnimControl(5).Activo = True
    AnimControl(5).Velocidad = 0
    AnimControl(5).Top = 552
    cSalir.Top = 10
    AnimControl(6).Activo = True
    AnimControl(6).Velocidad = 0
    AnimControl(6).Top = 552
    cOpciones.Top = 10
    AnimControl(7).Activo = True
    AnimControl(7).Velocidad = 0
    AnimControl(7).Top = 552
    
    Fuerza = 3.7 ' Gravedad... 1.7
    tEfectos.Interval = 10
    tEfectos.Enabled = True
     
    Call Audio.PlayMIDI("0.mid")
     
End Sub

Public Sub CleanToken()

    If Len(ClientConfigInit.Token) > 0 Then
        lblAccountName.Caption = vbNullString
        cCerrar.Visible = False
        ClientConfigInit.Token = vbNullString
        Call modGameIni.SaveConfigInit
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        cAcceder.Visible = True
    End If

End Sub

Public Sub TryConnectToken()

    If Len(ClientConfigInit.Token) > 0 And modGSZ.ValidJWT(ClientConfigInit.Token) Then
        lblAccountName.Caption = vbNullString
        fAccediendo.Visible = False
        cCerrar.Visible = True
        cAcceder.Visible = False
        tAccediendo.Enabled = False
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        EstadoLogin = E_MODO.Token
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        DoEvents
    End If

End Sub

Private Sub tAccediendo_Timer()
    ' Revisar si config.init cambio
    Dim fileStamp As Date
    fileStamp = FileDateTime(pathInits & fConfigInit)
    If fileStamp <> LastConfigInit Then ' Cambio!
        ClientConfigInit = modGameIni.LoadConfigInit
        If Len(ClientConfigInit.Token) > 0 Then
            Me.WindowState = vbNormal
            DoEvents
            Call TryConnectToken
        End If
    End If
End Sub

Private Sub cCrearPJ_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If Not frmMain.Socket1.Connected Then
        Call TryConnectToken
        DoEvents
    End If
    
    Call WriteLoginNewChar
    
'    EstadoLogin = E_MODO.Dados
'    CaptchaKey = RandomNumber(1, 255) ' GSZAO
'    If frmMain.Socket1.Connected Then
'        frmMain.Socket1.Disconnect
'        frmMain.Socket1.Cleanup
'        DoEvents
'    End If
'    frmMain.Socket1.HostAddress = CurServerIp
'    frmMain.Socket1.RemotePort = CurServerPort
'    frmMain.Socket1.Connect
End Sub

Private Sub cCreditos_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCreditos.Show vbModal
End Sub

Private Sub chkRecordar_Click()
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub cSalir_Click()
    prgRun = False
    End
End Sub

Private Sub cSitioOficial_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://" & SitioOficial, "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub cTeclas_Click()
    Call Audio.PlayWave(SND_CLICK)
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Make Server IP and Port box visible
'Hacer CONTROL+I
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    PortTxt.Text = "7666"
    IPTxt.Text = "127.0.0.1"
    PortTxt.Visible = True
    IPTxt.Visible = True
    CurServer = 1
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub lRemember_Click()
    chkRecordar.Checked = Not chkRecordar.Checked
    Call chkRecordar_Click
End Sub

Private Sub lRemember_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call chkRecordar.SetFocus
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ' GSZAO
        If LenB(txtPasswd.Text) <> 0 Then
            Call cConectar_Click
        End If
    End If
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cConectar_Click
End Sub

Private Sub cConectar_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    If Not frmMain.Socket1.Connected Then
        Call TryConnectToken
        DoEvents
    End If
    
    Dim eMD5 As New clsMD5
    UserName = txtNombre.Text
    UserPassword = eMD5.DigestStrToHexStr(txtPasswd.Text) ' GSZ
    
    Call WriteLoginExistingChar


'    If frmMain.Socket1.Connected Then
'        frmMain.Socket1.Disconnect
'        frmMain.Socket1.Cleanup
'        DoEvents
'    End If
'
'    Dim eMD5 As New clsMD5
'    'update user info
'    UserName = txtNombre.Text
'    UserPassword = eMD5.DigestStrToHexStr(txtPasswd.Text) ' GSZ
'    If CheckUserData(False) = True Then
'        EstadoLogin = Normal
'        frmMain.Socket1.HostAddress = CurServerIp
'        frmMain.Socket1.RemotePort = CurServerPort
'        frmMain.Socket1.Connect
'        DoEvents
'    End If
    

End Sub
