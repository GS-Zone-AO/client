VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForo.frx":0000
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ClientGSZAO.uAOButton cNuevoAnuncio 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   6060
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      TX              =   "Anuncio"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmForo.frx":14197
      PICF            =   "frmForo.frx":141B3
      PICH            =   "frmForo.frx":141CF
      PICV            =   "frmForo.frx":141EB
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
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
      Left            =   1140
      MaxLength       =   35
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   4620
   End
   Begin VB.TextBox txtPost 
      Appearance      =   0  'Flat
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
      Height          =   3960
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":14207
      Top             =   1935
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.ListBox lstTitulos 
      Appearance      =   0  'Flat
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
      Height          =   5100
      Left            =   765
      TabIndex        =   0
      Top             =   825
      Width           =   4785
   End
   Begin ClientGSZAO.uAOButton cCerrar 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   6060
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmForo.frx":1420D
      PICF            =   "frmForo.frx":14229
      PICH            =   "frmForo.frx":14245
      PICV            =   "frmForo.frx":14261
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ClientGSZAO.uAOButton cNuevoMensaje 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   6060
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      TX              =   "Nuevo Mensaje"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmForo.frx":1427D
      PICF            =   "frmForo.frx":14299
      PICH            =   "frmForo.frx":142B5
      PICV            =   "frmForo.frx":142D1
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ClientGSZAO.uAOButton cListaMsg 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6060
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      TX              =   "Listado"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmForo.frx":142ED
      PICF            =   "frmForo.frx":14309
      PICH            =   "frmForo.frx":14325
      PICV            =   "frmForo.frx":14341
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Morpheus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   255
      Left            =   1125
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image imgMarcoTexto 
      Height          =   465
      Left            =   1095
      Top             =   840
      Width           =   4725
   End
   Begin VB.Label lblAutor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Height          =   255
      Left            =   1125
      TabIndex        =   3
      Top             =   1455
      Width           =   4650
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   2
      Left            =   4320
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   1
      Left            =   2520
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   0
      Left            =   960
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmForo"
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

Private clsFormulario As clsFormMovementManager

' Para controlar las imagenes de fondo y el envio de posteos
Private ForoActual As eForumType
Private VerListaMsg As Boolean
Private Lectura As Boolean

Public ForoLimpio As Boolean
Private Sticky As Boolean

' Para restringir la visibilidad de los foros
Public Privilegios As Byte
Public ForosVisibles As eForumType
Public CanPostSticky As Byte

' Imagenes de fondo
Private FondosDejarMsg(0 To 2) As Picture
Private FondosListaMsg(0 To 2) As Picture

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
End Sub

Private Sub cListaMsg_Click()
    Call Audio.PlayWave(SND_CLICK)
    VerListaMsg = True
    ToogleScreen
End Sub

Private Sub cNuevoAnuncio_Click()
    Lectura = False
    VerListaMsg = False
    Sticky = True
    Call Audio.PlayWave(SND_CLICK)
    
    'Switch to proper background
    ToogleScreen
End Sub

Private Sub cNuevoMensaje_Click()
    Dim PostStyle As Byte
    
    Call Audio.PlayWave(SND_CLICK)
    
    If Not VerListaMsg Then
        If Not Lectura Then
        
            If Sticky Then
                PostStyle = GetStickyPost
            Else
                PostStyle = GetNormalPost
            End If

            Call WriteForumPost(txtTitulo.Text, txtPost.Text, PostStyle)
            
            ' Actualizo localmente
            Call clsForos.AddPost(ForoActual, txtTitulo.Text, CharName, txtPost.Text, Sticky)
            Call UpdateList
            
            VerListaMsg = True
        End If
    Else
        VerListaMsg = False
        Sticky = False
    End If
    
    Lectura = False
    
    'Switch to proper background
    ToogleScreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoForo = False
    Privilegios = 0
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
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

    ' Load pictures
    Set FondosListaMsg(eForumType.ieGeneral) = LoadPicture(pathGUI & "frmForoGeneral.jpg")
    Set FondosListaMsg(eForumType.ieREAL) = LoadPicture(pathGUI & "frmForoReal.jpg")
    Set FondosListaMsg(eForumType.ieCAOS) = LoadPicture(pathGUI & "frmForoCaos.jpg")
    
    Set FondosDejarMsg(eForumType.ieGeneral) = LoadPicture(pathGUI & "frmForoMsgGeneral.jpg")
    Set FondosDejarMsg(eForumType.ieREAL) = LoadPicture(pathGUI & "frmForoMsgReal.jpg")
    Set FondosDejarMsg(eForumType.ieCAOS) = LoadPicture(pathGUI & "frmForoMsgCaos.jpg")
    
    imgMarcoTexto.Picture = LoadPicture(pathGUI & "frmForoMarco.jpg")
    
    ' Initial config
    ForoActual = eForumType.ieGeneral
    VerListaMsg = True
    UpdateList
    
    ' Default background
    ToogleScreen
    
    ForoLimpio = False
    MirandoForo = True
    
    ' Si no es caos o gms, no puede ver el tab de caos.
    If (Privilegios And eForumVisibility.ieCAOS_MEMBER) = 0 Then imgTab(2).Visible = False
    
    ' Si no es armada o gm, no puede ver el tab de armadas.
    If (Privilegios And eForumVisibility.ieREAL_MEMBER) = 0 Then imgTab(1).Visible = False
    
End Sub

Private Sub imgTab_Click(Index As Integer)

    Call Audio.PlayWave(SND_CLICK)
    
    If Index <> ForoActual Then
        ForoActual = Index
        VerListaMsg = True
        Lectura = False
        UpdateList
        ToogleScreen
    Else
        If Not VerListaMsg Then
            VerListaMsg = True
            Lectura = False
            ToogleScreen
        End If
    End If
    
End Sub

Private Sub ToogleScreen()
    
    Dim PostOffset As Integer
    
    imgMarcoTexto.Visible = Not VerListaMsg And Not Lectura
    txtTitulo.Visible = Not VerListaMsg And Not Lectura
    lblTitulo.Visible = Not VerListaMsg And Lectura
    
    cNuevoMensaje.Enabled = VerListaMsg Or Lectura
    
    txtPost.Visible = Not VerListaMsg
    
    cNuevoAnuncio.Visible = VerListaMsg And PuedeDejarAnuncios
    cListaMsg.Visible = Not VerListaMsg
    lstTitulos.Visible = VerListaMsg
    
    If VerListaMsg Then
        Me.Picture = FondosListaMsg(ForoActual)
    Else
        If Lectura Then
            With lstTitulos
                PostOffset = .ItemData(.ListIndex)
                ' Normal post?
                If PostOffset < STICKY_FORUM_OFFSET Then
                    lblTitulo.Caption = Foros(ForoActual).GeneralTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).GeneralPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).GeneralAuthor(PostOffset)
                
                ' Sticky post
                Else
                    PostOffset = PostOffset - STICKY_FORUM_OFFSET
                    lblTitulo.Caption = Foros(ForoActual).StickyTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).StickyPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).StickyAuthor(PostOffset)
                End If
            End With
        Else
            lblAutor.Caption = CharName
            txtTitulo.Text = vbNullString
            txtPost.Text = vbNullString
            
            txtTitulo.SetFocus
        End If
        
        txtPost.Locked = Lectura
        Me.Picture = FondosDejarMsg(ForoActual)
    End If
    
End Sub

Private Function PuedeDejarAnuncios() As Boolean
    
    ' No puede
    If CanPostSticky = 0 Then Exit Function

    If ForoActual = eForumType.ieGeneral Then
        ' Solo puede dejar en el general si es gm
        If CanPostSticky <> 2 Then Exit Function
    End If
    
    PuedeDejarAnuncios = True
    
End Function

Private Sub lstTitulos_Click()
    VerListaMsg = False
    Lectura = True
    ToogleScreen
End Sub

Private Sub txtPost_Change()
    If Lectura Then Exit Sub
    cNuevoMensaje.Enabled = (Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0)
End Sub

Private Sub txtTitulo_Change()
    If Lectura Then Exit Sub
    cNuevoMensaje.Enabled = (Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0)
End Sub

Private Sub UpdateList()
    Dim PostIndex As Long
    
    lstTitulos.Clear
    
    With lstTitulos
        ' Sticky first
        For PostIndex = 1 To clsForos.GetNroSticky(ForoActual)
            .AddItem "[ANUNCIO] " & Foros(ForoActual).StickyTitle(PostIndex) & " (" & Foros(ForoActual).StickyAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = STICKY_FORUM_OFFSET + PostIndex
        Next PostIndex
    
        ' Then normal posts
        For PostIndex = 1 To clsForos.GetNroPost(ForoActual)
            .AddItem Foros(ForoActual).GeneralTitle(PostIndex) & " (" & Foros(ForoActual).GeneralAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = PostIndex
        Next PostIndex
    End With
    
End Sub

Private Function GetStickyPost() As Byte
    Select Case ForoActual
        Case 0
            GetStickyPost = eForumMsgType.ieGENERAL_STICKY
            
        Case 1
            GetStickyPost = eForumMsgType.ieREAL_STICKY
            
        Case 2
            GetStickyPost = eForumMsgType.ieCAOS_STICKY
            
    End Select
    
End Function

Private Function GetNormalPost() As Byte
    Select Case ForoActual
        Case 0
            GetNormalPost = eForumMsgType.ieGeneral
            
        Case 1
            GetNormalPost = eForumMsgType.ieREAL
            
        Case 2
            GetNormalPost = eForumMsgType.ieCAOS
            
    End Select
    
End Function
