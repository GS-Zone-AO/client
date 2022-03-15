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
   ScaleMode       =   0  'User
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3720
      TabIndex        =   18
      Top             =   3960
      Width           =   4455
      Begin VB.Label lblLastConnect 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Fecha conectado]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   2025
         TabIndex        =   25
         Top             =   2280
         Width           =   2250
      End
      Begin VB.Label lblGold 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro: ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRace 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   3645
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblClass 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   3645
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCharName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin ClientGSZAO.uAOButton cCrearPJ 
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "&Crear Nuevo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E450
      PICF            =   "frmConnect.frx":4E46C
      PICH            =   "frmConnect.frx":4E488
      PICV            =   "frmConnect.frx":4E4A4
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
   Begin VB.Timer tAccediendo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10320
      Top             =   1440
   End
   Begin VB.ListBox lPersonajes 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3540
      ItemData        =   "frmConnect.frx":4E4C0
      Left            =   8760
      List            =   "frmConnect.frx":4E4C2
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame fAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      Begin ClientGSZAO.uAOButton cCerrar 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         TX              =   "Cerrar sesión"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":4E4C4
         PICF            =   "frmConnect.frx":4E4E0
         PICH            =   "frmConnect.frx":4E4FC
         PICV            =   "frmConnect.frx":4E518
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClientGSZAO.uAOButton cOpciones 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         TX              =   "Opciones"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":4E534
         PICF            =   "frmConnect.frx":4E550
         PICH            =   "frmConnect.frx":4E56C
         PICV            =   "frmConnect.frx":4E588
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         TX              =   "Creditos"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":4E5A4
         PICF            =   "frmConnect.frx":4E5C0
         PICH            =   "frmConnect.frx":4E5DC
         PICV            =   "frmConnect.frx":4E5F8
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
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         TX              =   "Sitio Oficial"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":4E614
         PICF            =   "frmConnect.frx":4E630
         PICH            =   "frmConnect.frx":4E64C
         PICV            =   "frmConnect.frx":4E668
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
   End
   Begin ClientGSZAO.uAOButton cManualToken 
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   5640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      TX              =   "Ingreso manual"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E684
      PICF            =   "frmConnect.frx":4E6A0
      PICH            =   "frmConnect.frx":4E6BC
      PICV            =   "frmConnect.frx":4E6D8
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tEfectos 
      Enabled         =   0   'False
      Left            =   9720
      Top             =   1440
   End
   Begin ClientGSZAO.uAOButton cSalir 
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E6F4
      PICF            =   "frmConnect.frx":4E710
      PICH            =   "frmConnect.frx":4E72C
      PICV            =   "frmConnect.frx":4E748
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
      Left            =   8760
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Conectarse"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E764
      PICF            =   "frmConnect.frx":4E780
      PICH            =   "frmConnect.frx":4E79C
      PICV            =   "frmConnect.frx":4E7B8
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
      Left            =   6000
      TabIndex        =   3
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
      Left            =   5010
      TabIndex        =   2
      Text            =   "7666"
      Top             =   240
      Visible         =   0   'False
      Width           =   825
   End
   Begin ClientGSZAO.uAOButton cAcceder 
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   4920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      TX              =   "Acceder"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E7D4
      PICF            =   "frmConnect.frx":4E7F0
      PICH            =   "frmConnect.frx":4E80C
      PICV            =   "frmConnect.frx":4E828
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ClientGSZAO.uAOButton cBorrar 
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      TX              =   "Borrar personaje"
      ENAB            =   -1  'True
      FCOL            =   8421631
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":4E844
      PICF            =   "frmConnect.frx":4E860
      PICH            =   "frmConnect.frx":4E87C
      PICV            =   "frmConnect.frx":4E898
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPersonajes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mis personajes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   9240
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Shape sAccount 
      BackColor       =   &H00202020&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image iAccount 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblAccessing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accede para continuar."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   4080
      TabIndex        =   11
      Top             =   4560
      Width           =   2805
   End
   Begin VB.Label lblConnecting 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   4080
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblAccountName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Invitado]"
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   855
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
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
   Begin VB.Shape sPersonajes 
      BackColor       =   &H00202020&
      BackStyle       =   1  'Opaque
      Height          =   4935
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
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
    tAccediendo.Enabled = True
    
End Sub

Private Sub cBorrar_Click()

    Dim Name As String
    Name = InputBox("Vuelve a ingresar tal como está escrito el nombre del personaje que deseas borrar.", "Borrar personaje", "")
    If lblCharName.Caption <> Name Then
       Call MsgBox("El personaje " & lblCharName.Caption & " no será eliminado.", vbExclamation + vbOKOnly)
       Exit Sub
    End If
    Dim confirm As Long
    confirm = MsgBox("¿Está completamente seguro de que desea eliminar el personaje de " & Name & "?" & vbCrLf & vbCrLf & _
                "ADVERTENCIA: Está acción es irreversible, no tiene recuperación.", vbCritical + vbYesNoCancel, "ADVERTENCIA")
    If confirm = vbYes Then
        If frmMain.Socket1.Connected Then
            Call WriteDeleteChar(Name)
        Else
            Call Disconnected
        End If
    End If

End Sub

Private Sub cCerrar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call ShowMenu(False)
    Call CleanToken
    
End Sub

Private Sub cManualToken_Click()

    frmConnectToken.Show
    
End Sub

Private Sub cOpciones_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call ShowMenu(False)
    frmOpciones.Show vbModal
End Sub

Private Sub Form_Activate()

    Call Audio.PlayMIDI("0.mid")
    
    If frmMain.Socket1.Connected Then
        Call Connected
    End If
    
End Sub

Sub ShowMenu(ByVal Visible As Boolean)

    sAccount.Visible = Visible
    fAccount.Visible = sAccount.Visible
    
End Sub

Sub Connected()

    ' Header
    lblAccountName.Caption = AccountName

    ' Mensajes
    lblConnecting.Visible = False

    ' Acciones
    tAccediendo.Enabled = False
    cCerrar.Visible = True
    cAcceder.Visible = False
    cManualToken.Visible = False
    
    ' Personajes
    lblPersonajes.Visible = True
    sPersonajes.Visible = True
    lPersonajes.Visible = True
    cCrearPJ.Visible = True
    fChar.Visible = True
    cBorrar.Visible = False
    cConectar.Visible = False
    
    If NumberOfCharacters > 0 Then
        Call LoadCharacters
    End If
    
    lblCharName.Caption = "Seleciona un personaje"
    lblLevel.Visible = False
    lblGold.Visible = False
    lblLastConnect.Visible = False
    lblClass.Visible = False
    lblRace.Visible = False

End Sub

Sub Disconnected()

    ' Header
    lblAccountName.Caption = "[Invitado]"
    
    ' Acciones
    tAccediendo.Enabled = False
    cCerrar.Visible = False
    cAcceder.Visible = True
    cManualToken.Visible = True

    ' Mensajes
    lblConnecting.Visible = False
    lblAccessing.Visible = True
    
    ' Personajes
    lblPersonajes.Visible = False
    sPersonajes.Visible = False
    lPersonajes.Visible = False
    cCrearPJ.Visible = False
    cBorrar.Visible = False
    cConectar.Visible = False
    fChar.Visible = False
    
    If Not frmConnect.Visible Then
        frmConnect.Visible = True
    End If

End Sub


Public Sub CleanToken()

    ' Token
    TokenConnected = vbNullString
    ClientConfigInit.Token = vbNullString
    Call modGameIni.SaveConfigInit
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    Call Disconnected

End Sub

Private Sub iAccount_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call ShowMenu(Not sAccount.Visible)
    
End Sub

Private Sub lPersonajes_Click()
    
    CharName = Split(lPersonajes.List(lPersonajes.ListIndex), " [")(0)
    Dim i As Byte
    For i = 1 To NumberOfCharacters
        If Characters(i).Name = CharName Then
            If Characters(i).Level = 0 Then Exit For
        
            lblCharName.Caption = CharName
            lblLevel.Caption = "Nivel " & Characters(i).Level
            lblClass.Caption = ListaClases(Characters(i).Class)
            lblRace.Caption = ListaRazas(Characters(i).Race)
            lblGold.Caption = "Oro: " & Characters(i).Gold
            lblLastConnect.Caption = Characters(i).LastConnect
            
            ' Campos
            lblCharName.Visible = True
            lblLevel.Visible = True
            lblGold.Visible = True
            lblLastConnect.Visible = True
            lblClass.Visible = True
            lblRace.Visible = True
            
            ' Acciones
            cBorrar.Visible = True
            cConectar.Visible = True
            
            Exit For
        End If
    Next
    fChar.Visible = True
    
End Sub

Private Sub tEfectos_Timer()
'    Dim oTop As Integer
'    Dim i As Integer
'    For i = 1 To 7
'        If AnimControl(i).Activo = True Then
'            Select Case i
'                Case 1: oTop = cTeclas.Top
'                Case 2: oTop = cConectar.Top
'                Case 3: oTop = cCrearPJ.Top
'                Case 4: oTop = cSitioOficial.Top
'                Case 5: oTop = cCreditos.Top
'                Case 6: oTop = cSalir.Top
'                Case 7: oTop = cOpciones.Top
'            End Select
'            If oTop > AnimControl(i).Top Then
'                oTop = AnimControl(i).Top
'                AnimControl(i).Velocidad = AnimControl(i).Velocidad * -0.6
'            End If
'            If AnimControl(i).Velocidad >= -0.6 And AnimControl(i).Velocidad <= -0.5 Then
'                AnimControl(i).Activo = False
'            Else
'                AnimControl(i).Velocidad = AnimControl(i).Velocidad + Fuerza
'                oTop = oTop + AnimControl(i).Velocidad
'            End If
'            Select Case i
'                Case 1: cTeclas.Top = oTop
'                Case 2: cConectar.Top = oTop
'                Case 3: cCrearPJ.Top = oTop
'                Case 4: cSitioOficial.Top = oTop
'                Case 5: cCreditos.Top = oTop
'                Case 6: cSalir.Top = oTop
'                Case 7: cOpciones.Top = oTop
'            End Select
'        End If
'    Next
'    If AnimControl(1).Activo = False And AnimControl(2).Activo = False And AnimControl(3).Activo = False And _
'       AnimControl(4).Activo = False And AnimControl(5).Activo = False And AnimControl(6).Activo = False And _
'       AnimControl(7).Activo = False Then
'        tEfectos.Enabled = False
'        cTeclas.Top = AnimControl(1).Top
'        cConectar.Top = AnimControl(2).Top
'        cCrearPJ.Top = AnimControl(3).Top
'        cSitioOficial.Top = AnimControl(4).Top
'        cCreditos.Top = AnimControl(5).Top
'        cSalir.Top = AnimControl(6).Top
'        cOpciones.Top = AnimControl(7).Top
'    End If
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
    
    ' GSZAO - Animación...
'    cTeclas.Top = 10
'    AnimControl(1).Activo = True
'    AnimControl(1).Velocidad = 0
'    AnimControl(1).Top = 400
'    cConectar.Top = 10
'    AnimControl(2).Activo = True
'    AnimControl(2).Velocidad = 0
'    AnimControl(2).Top = 400
'    cCrearPJ.Top = 10
'    AnimControl(3).Activo = True
'    AnimControl(3).Velocidad = 0
'    AnimControl(3).Top = 552
'    cSitioOficial.Top = 10
'    AnimControl(4).Activo = True
'    AnimControl(4).Velocidad = 0
'    AnimControl(4).Top = 552
'    cCreditos.Top = 10
'    AnimControl(5).Activo = True
'    AnimControl(5).Velocidad = 0
'    AnimControl(5).Top = 552
'    cSalir.Top = 10
'    AnimControl(6).Activo = True
'    AnimControl(6).Velocidad = 0
'    AnimControl(6).Top = 552
'    cOpciones.Top = 10
'    AnimControl(7).Activo = True
'    AnimControl(7).Velocidad = 0
'    AnimControl(7).Top = 552
'
'    Fuerza = 3.7 ' Gravedad... 1.7
'    tEfectos.Interval = 10
'    tEfectos.Enabled = True

    If Len(ClientConfigInit.Token) = 0 Then
        Call Disconnected
    End If
     
    Call Audio.PlayMIDI("0.mid")
     
End Sub

Public Sub TryConnectToken()

    If Len(ClientConfigInit.Token) > 0 And modGSZ.ValidJWT(ClientConfigInit.Token) Then
        lblAccountName.Caption = "[Accediendo...]"
        tAccediendo.Enabled = False
        lblConnecting.Visible = True
        lblAccessing.Visible = False
        cAcceder.Visible = False
        cManualToken.Visible = False

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        
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
    If fileStamp <> LastConfigInit And _
        TokenConnected <> ClientConfigInit.Token Then ' Cambio!
        
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
    Me.Visible = False
    frmCrearPersonaje.Visible = True
    
End Sub

Private Sub cCreditos_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call ShowMenu(False)
    frmCreditos.Show vbModal
    
End Sub

Private Sub chkRecordar_Click()

    Call Audio.PlayWave(SND_CLICK)
    
End Sub

Private Sub cSalir_Click()

    Call Audio.PlayWave(SND_CLICK)
    prgRun = False
    End
    
End Sub

Private Sub cSitioOficial_Click()

    Call Audio.PlayWave(SND_CLICK)
    Call ShowMenu(False)
    Call ShellExecute(0, "Open", "http://" & SitioOficial, "", App.Path, SW_SHOWNORMAL)
    
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

Private Sub cConectar_Click()

    Call Audio.PlayWave(SND_CLICK)
    
    If Not frmMain.Socket1.Connected Then
        Call TryConnectToken
        DoEvents
    End If
    
    Call WriteExistingChar

End Sub
