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
         _extentx        =   3836
         _extenty        =   873
         tx              =   "Token Manual"
         enab            =   -1  'True
         fcol            =   7314354
         ocol            =   16777215
         pice            =   "frmConnect.frx":45744
         picf            =   "frmConnect.frx":45760
         pich            =   "frmConnect.frx":4577C
         picv            =   "frmConnect.frx":45798
         font            =   "frmConnect.frx":457B4
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
      _extentx        =   344
      _extenty        =   344
      chck            =   0   'False
      enab            =   -1  'True
      picc            =   "frmConnect.frx":457D8
   End
   Begin ClientGSZAO.uAOButton cSalir 
      Height          =   495
      Left            =   9960
      TabIndex        =   11
      Top             =   8280
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      tx              =   "Salir"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45838
      picf            =   "frmConnect.frx":45854
      pich            =   "frmConnect.frx":45870
      picv            =   "frmConnect.frx":4588C
      font            =   "frmConnect.frx":458A8
   End
   Begin ClientGSZAO.uAOButton cCreditos 
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   8280
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      tx              =   "Creditos"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":458CC
      picf            =   "frmConnect.frx":458E8
      pich            =   "frmConnect.frx":45904
      picv            =   "frmConnect.frx":45920
      font            =   "frmConnect.frx":4593C
   End
   Begin ClientGSZAO.uAOButton cSitioOficial 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   8280
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      tx              =   "Sitio Oficial"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45960
      picf            =   "frmConnect.frx":4597C
      pich            =   "frmConnect.frx":45998
      picv            =   "frmConnect.frx":459B4
      font            =   "frmConnect.frx":459D0
   End
   Begin ClientGSZAO.uAOButton cCrearPJ 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   8280
      Width           =   4575
      _extentx        =   8070
      _extenty        =   873
      tx              =   "Nuevo Personaje"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":459F4
      picf            =   "frmConnect.frx":45A10
      pich            =   "frmConnect.frx":45A2C
      picv            =   "frmConnect.frx":45A48
      font            =   "frmConnect.frx":45A64
   End
   Begin ClientGSZAO.uAOButton cConectar 
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   6000
      Width           =   2775
      _extentx        =   4895
      _extenty        =   873
      tx              =   "Conectarse"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45A88
      picf            =   "frmConnect.frx":45AA4
      pich            =   "frmConnect.frx":45AC0
      picv            =   "frmConnect.frx":45ADC
      font            =   "frmConnect.frx":45AF8
   End
   Begin ClientGSZAO.uAOButton cTeclas 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      tx              =   "Teclas"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45B1C
      picf            =   "frmConnect.frx":45B38
      pich            =   "frmConnect.frx":45B54
      picv            =   "frmConnect.frx":45B70
      font            =   "frmConnect.frx":45B8C
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
      _extentx        =   3413
      _extenty        =   873
      tx              =   "Opciones"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45BB0
      picf            =   "frmConnect.frx":45BCC
      pich            =   "frmConnect.frx":45BE8
      picv            =   "frmConnect.frx":45C04
      font            =   "frmConnect.frx":45C20
   End
   Begin ClientGSZAO.uAOButton cAcceder 
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   4200
      Width           =   2775
      _extentx        =   4895
      _extenty        =   873
      tx              =   "Acceder"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45C44
      picf            =   "frmConnect.frx":45C60
      pich            =   "frmConnect.frx":45C7C
      picv            =   "frmConnect.frx":45C98
      font            =   "frmConnect.frx":45CB4
   End
   Begin ClientGSZAO.uAOButton cCerrar 
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
      _extentx        =   4895
      _extenty        =   873
      tx              =   "Cerrar sesión"
      enab            =   -1  'True
      fcol            =   7314354
      ocol            =   16777215
      pice            =   "frmConnect.frx":45CD8
      picf            =   "frmConnect.frx":45CF4
      pich            =   "frmConnect.frx":45D10
      picv            =   "frmConnect.frx":45D2C
      font            =   "frmConnect.frx":45D48
   End
   Begin VB.Label lRemember 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar contraseña"
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
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

' Animación de los Controles...
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

    cCerrar.Visible = False
    cAcceder.Visible = True
    ClientConfigInit.Token = vbNullString
    Call modGameIni.SaveConfigInit
    
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

Private Sub tAccediendo_Timer()
    ' Revisar si config.init cambio
    Dim fileStamp As Date
    fileStamp = FileDateTime(pathInits & fConfigInit)
    If fileStamp <> LastConfigInit Then ' Cambio!
        ClientConfigInit = modGameIni.LoadConfigInit
        If Len(ClientConfigInit.Token) > 0 Then
            MsgBox "GSZAO Accedio con " & ClientConfigInit.Token
            Me.WindowState = vbNormal
            fAccediendo.Visible = False
            cCerrar.Visible = True
            tAccediendo.Enabled = False
        End If
    End If
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
    
    If Len(ClientConfigInit.Token) > 0 Then
        cCerrar.Visible = True
        cAcceder.Visible = False
        MsgBox "GSZAO Falta programar. Tiene token, es valido? " & ClientConfigInit.Token
    End If

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
    
    ' GSZAO - Animación...
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

Public Sub EstadoSocket()
    If frmMain.Socket1.Connected Then
        txtNombre.Enabled = False
        txtPasswd.Enabled = False
        Me.MousePointer = 11
    Else
        txtNombre.Enabled = True
        txtPasswd.Enabled = True
        Me.MousePointer = 0
    End If
End Sub

Private Sub cCrearPJ_Click()
    Call Audio.PlayWave(SND_CLICK)
    EstadoLogin = E_MODO.Dados
    CaptchaKey = RandomNumber(1, 255) ' GSZAO
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostAddress = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
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

Private Sub lRemember_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    Dim eMD5 As New clsMD5
    'update user info
    UserName = txtNombre.Text
    UserPassword = eMD5.DigestStrToHexStr(txtPasswd.Text) ' GSZ
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        DoEvents
    End If
    

End Sub
