Attribute VB_Name = "modGameIni"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M?rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat?as Fernando Peque?o
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
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

' GSZAO - Archivos de configuraci?n!
Public Const fAOSetup = "AOSetup.init"
Public Const fConfigInit = "Config.init"

' GSZAO - Las variables de path se definen una sola vez (Ver Sub InitFilePaths())
Public pathGraphics As String
Public pathInterface As String
Public pathExtras As String
Public pathCursors As String
Public pathGUI As String
Public pathButtons As String
Public pathParticles As String
Public pathSound As String
Public pathMusic As String
Public pathMaps As String

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tConfigInit
    ' Opciones
    MostrarTips As Byte         ' Activa o desactiva la muestra de tips
    NumParticulas As Integer    ' Numero de particulas
    IndiceGraficos As String    ' Archivo de Indices de Graficos
    
    ' Usuario
    Token As String             ' Token del usuario
    
    ' Directorio
    pathGraphics As String       ' Directorio de graficos
    pathInterface As String      ' Directorio de interfaces
    DirCursores As String       ' Directorio de cursores
    DirGUI As String            ' Directorio del GUI
    DirBotones As String        ' Directorio de botones
    pathParticles As String     ' Directorio de particulas
    DirSonidos As String        ' Directorio de sonidos
    pathMusicas As String        ' Directorio de musicas
    pathMaps As String          ' Directorio de mapas
    pathExtras As String         ' Directorio de extras
    DirFotos As String          ' Directorio de fotos
    DirFrags As String          ' Directorio de frags
    DirMuertes As String        ' Directorio de muertes
End Type

Public Type tAOSetup
    ' VIDEO
    bVertex     As Byte     ' GSZAO - Cambia el Vortex de dibujado
    bVSync      As Boolean  ' GSZAO - Utiliza Sincronizaci?n Vertical (VSync)
    bDinamic    As Boolean  ' Utilizar carga Dinamica de Graficos o Estatica
    byMemory    As Byte     ' Uso maximo de memoria para la carga Dinamica (exclusivamente)

    ' SONIDO
    bNoMusic    As Boolean  ' Jugar sin Musica
    bNoSound    As Boolean  ' Jugar sin Sonidos
    bNoSoundEffects As Boolean  ' Jugar sin Efectos de sonido (basicamente, sonido que viene de la izquierda y de la derecha)
    lMusicVolume As Long ' Volumen de la Musica
    lSoundVolume As Long ' Volumen de los Sonidos

    ' SCREENSHOTS
    bActive     As Boolean  ' Activa el modo de screenshots
    bDie        As Boolean  ' Obtiene una screenshot al morir (si bActive = True)
    bKill       As Boolean  ' Obtiene una screenshot al matar (si bActive = True)
    byMurderedLevel As Byte ' La screenshot al matar depende del nivel de la victima (si bActive = True)
    
    ' CLAN
    bGuildNews  As Boolean      ' Mostrar Noticias del Clan al inicio
    bGldMsgConsole As Boolean   ' Activa los Dialogos de Clan
    bCantMsgs   As Byte         ' Establece el maximo de mensajes de Clan en pantalla
    
    ' GENERALEs
    bCursores   As Boolean      ' Utilizar Cursores Personalizados
End Type

Public MiCabecera As tCabecera
Public TokenConnected As String
Public ClientConfigInit As tConfigInit
Public ClientAOSetup As tAOSetup
Public LastConfigInit As Date

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'**************************************************************
'Author: Unknown
'Last Modify Date: 19/01/2022 - ^[GS]^
'**************************************************************
    Cabecera.Desc = "GS-Zone Argentum Online MOD - Copyright GS-Zone 2022 - info@gs-zone.org - Original by Pablo Marquez " ' GSZAO
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
    
End Sub

Public Function LoadConfigInit() As tConfigInit
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 19/01/2022 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    Dim N As Integer
    Dim ConfigInit As tConfigInit
    N = FreeFile
    Open pathInits & fConfigInit For Binary As #N
        Get #N, , MiCabecera
        Get #N, , ConfigInit
    Close #N
    
    LastConfigInit = FileDateTime(pathInits & fConfigInit)
    LoadConfigInit = ConfigInit
    
End Function

Public Sub SaveConfigInit()
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 19/01/2022 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    Dim ImaConfigInit As tConfigInit
    ImaConfigInit = ClientConfigInit

    Dim N As Integer
    N = FreeFile
    Open pathInits & fConfigInit For Binary As #N
    Put #N, , MiCabecera
    Put #N, , ImaConfigInit
    Close #N
    
End Sub

Public Sub InitFilePaths()
'*************************************************
'Author: ^[GS]^
'Last modified: 09/01/2022 - ^[GS]^
'*************************************************
    If InStr(1, ClientConfigInit.IndiceGraficos, "Graficos") Then ' indice de graficos
        GraphicsFile = ClientConfigInit.IndiceGraficos
    Else
        GraphicsFile = "Graficos1.ind"
    End If
    'Requeridos
    Call FileRequired(pathInits & GraphicsFile)
    pathGraphics = ValidDirectory(pathClient & ClientConfigInit.pathGraphics)
    pathInterface = ValidDirectory(pathClient & ClientConfigInit.pathInterface)
    pathCursors = ValidDirectory(pathClient & ClientConfigInit.DirCursores)
    pathGUI = ValidDirectory(pathClient & ClientConfigInit.DirGUI)
    pathButtons = ValidDirectory(pathClient & ClientConfigInit.DirBotones)
    pathParticles = ValidDirectory(pathClient & ClientConfigInit.pathParticles)
    pathSound = ValidDirectory(pathClient & ClientConfigInit.DirSonidos)
    pathMusic = ValidDirectory(pathClient & ClientConfigInit.pathMusicas)
    pathMaps = ValidDirectory(pathClient & ClientConfigInit.pathMaps)
    ' Opcionales
    pathExtras = ValidDirectory(pathClient & ClientConfigInit.pathExtras)
    If Not FileExist(pathExtras, vbDirectory) Then
        MkDir ValidDirectory(pathExtras)
    End If
    If Not FileExist(pathClient & ClientConfigInit.DirFotos, vbDirectory) Then
        MkDir ValidDirectory(pathClient & ClientConfigInit.DirFotos)
    End If
    If Not FileExist(pathClient & ClientConfigInit.DirFrags, vbDirectory) Then
        MkDir ValidDirectory(pathClient & ClientConfigInit.DirFrags)
    End If
    If Not FileExist(pathClient & ClientConfigInit.DirMuertes, vbDirectory) Then
        MkDir ValidDirectory(pathClient & ClientConfigInit.DirMuertes)
    End If
End Sub

Public Sub LoadClientAOSetup()
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modification: 22/08/2013 - ^[GS]^
'**************************************************************
    Dim fHandle As Integer
    
    ' Por default
    ClientAOSetup.bDinamic = True
    ClientAOSetup.bVertex = 0 ' software
    ClientAOSetup.bVSync = False
    
    If FileExist(pathInits & fAOSetup, vbArchive) Then
        fHandle = FreeFile
        Open pathInits & fAOSetup For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientAOSetup
        Close fHandle
    End If
    
    ClientAOSetup.bGuildNews = Not ClientAOSetup.bGuildNews
    Set DialogosClanes = New clsGuildDlg ' 0.13.3
    DialogosClanes.Activo = Not ClientAOSetup.bGldMsgConsole
    DialogosClanes.CantidadDialogos = ClientAOSetup.bCantMsgs
End Sub

Public Sub SaveClientAOSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 22/08/2013 - ^[GS]^
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    ClientAOSetup.bNoMusic = Not Audio.MusicActivated
    ClientAOSetup.bNoSound = Not Audio.SoundActivated
    ClientAOSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    ClientAOSetup.bGuildNews = Not ClientAOSetup.bGuildNews
    ClientAOSetup.bGldMsgConsole = Not DialogosClanes.Activo
    ClientAOSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    ClientAOSetup.lMusicVolume = Audio.MusicVolume
    ClientAOSetup.lSoundVolume = Audio.SoundVolume
    
    Open pathInits & fAOSetup For Binary As fHandle
        Put fHandle, , ClientAOSetup
    Close fHandle
    
End Sub

Public Function SEncriptar(ByVal Cadena As String) As String
' GSZ-AO - Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    DoEvents

End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
' GSZ-AO - Desencripta una cadena de texto
    Dim i As Long, NumDesencriptar As String
    
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    SDesencriptar = Cadena
    DoEvents

End Function

Public Function SXor(ByVal Cadena As String, ByVal Clave As String) As String
' GSZ-AO - Aplicamos un XOR por la clave indicada
    Dim i As Long, c As Integer

    If Len(Cadena) > 0 Then
        c = 1
        For i = 1 To Len(Cadena)
            If c > Len(Clave) Then c = 1
            Mid$(Cadena, i, 1) = Chr(Asc(mid$(Cadena, i, 1)) Xor Asc(mid$(Clave, c, 1)))
            c = c + 1
        Next i
    End If
    SXor = Cadena
    DoEvents

End Function
