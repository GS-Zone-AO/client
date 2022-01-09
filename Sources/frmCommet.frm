VERSION 5.00
Begin VB.Form frmCommet 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCommet.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin ClientGSZAO.uAOButton cEnviar 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Enviar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCommet.frx":D7B4
      PICF            =   "frmCommet.frx":D7D0
      PICH            =   "frmCommet.frx":D7EC
      PICV            =   "frmCommet.frx":D808
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
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCommet.frx":D824
      PICF            =   "frmCommet.frx":D840
      PICH            =   "frmCommet.frx":D85C
      PICV            =   "frmCommet.frx":D878
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
Attribute VB_Name = "frmCommet"
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

Option Explicit

Private clsFormulario As clsFormMovementManager
Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Public nombre As String
Public T As TIPO
Public Enum TIPO
    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3
End Enum

Private Sub cCerrar_Click()

    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    
End Sub

Private Sub cEnviar_Click()

    If nombre = vbNullString Then Exit Sub ' GSZAO

    Call Audio.PlayWave(SND_CLICK)
    If LenB(text1) = 0 Then
        If T = PAZ Or T = ALIANZA Then
            MsgBox "Debes redactar un mensaje solicitando la paz o alianza al líder de " & nombre
        Else
            MsgBox "Debes indicar el motivo por el cual rechazas la membresía de " & nombre
        End If
        
        Exit Sub
    End If
    
    If T = PAZ Then
        Call WriteGuildOfferPeace(nombre, Replace(text1, vbCrLf, "º"))
        
    ElseIf T = ALIANZA Then
        Call WriteGuildOfferAlliance(nombre, Replace(text1, vbCrLf, "º"))
        
    ElseIf T = RECHAZOPJ Then
        Call WriteGuildRejectNewMember(nombre, Replace(Replace(text1.Text, ",", " "), vbCrLf, " "))
        'Sacamos el char de la lista de aspirantes
        Dim i As Long
        
        For i = 0 To frmGuildLeader.solicitudes.ListCount - 1
            If frmGuildLeader.solicitudes.List(i) = nombre Then
                frmGuildLeader.solicitudes.RemoveItem i
                Exit For
            End If
        Next i
        
        Me.Hide
        Unload frmCharInfo
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    ' Fondo
    Select Case T
        Case TIPO.ALIANZA
            Me.Picture = LoadPicture(pathGUI & "frmCommetAlianza.jpg")
        Case TIPO.PAZ
            Me.Picture = LoadPicture(pathGUI & "frmCommetPaz.jpg")
        Case TIPO.RECHAZOPJ
            Me.Picture = LoadPicture(pathGUI & "frmCommetRechazo.jpg")
    End Select

    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(pathButtons & sty_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(pathButtons & sty_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(pathButtons & sty_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(pathButtons & sty_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(pathButtons & sty_cCheckbox))
        End If
    Next
    
End Sub

Private Sub Text1_Change()

    If Len(text1.Text) > MAX_PROPOSAL_LENGTH Then _
        text1.Text = Left$(text1.Text, MAX_PROPOSAL_LENGTH)
        
End Sub

