Attribute VB_Name = "modDeclaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

'Code by Shige
Public Type TGUID
        data1 As Long
        data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
End Type

Public Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long


'TORNEO
Public OpcionTorneo As Byte
Public Participantes() As String
'TORNEO

' *********************************************************
' ******************** CONFIGURACI�N **********************
' *********************************************************

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef destination As Any, _
                                      ByVal length As Long)
                                      
' ObtenerIP:
Public Const ObtenerIP As Byte = 0
' 0 = Utiliza Sinfo.dat (MODO ORIGINAL)
' 1 = Utiliza una fuente "fija", configurada en modGeneral, la funci�n "CargarServidores()".
' 2 = Utiliza una fuente obtenida de internet, configurada en modGeneral, la funci�n "CargarServidores()".
Public Const NombreCliente As String = "GS-Zone AO"
' NombreCliente:
' Define el nombre del cliente para todos los mensajes ;)
Public Const SitioOficial As String = "www.gs-zone.org"
' SitioOficial:
' Define cual es la web oficial del juego. Colocar sin "http://"!!!
Public Const TutorialDefualt As Boolean = False
' Ciudades:
Public Const NUMCIUDADES As Byte = 3 ' GSZAO
Public Ciudades(1 To NUMCIUDADES) As String ' Ciudades, configurarlas en Private Sub InicializarNombres()

' *********************************************************
' ******************** CONFIGURACI�N **********************
' *********************************************************

' * Configuraci�n de estilos de controles
Public Const sty_bEsquina As String = "uAOButton_bEsquina.bmp"
Public Const sty_bFondo As String = "uAOButton_bFondo.bmp"
Public Const sty_bHorizontal As String = "uAOButton_bHorizontal.bmp"
Public Const sty_bVertical As String = "uAOButton_bVertical.bmp"
Public Const sty_cCheckbox As String = "uAOCheckbox_cCheckbox.bmp" ' Grande
Public Const sty_cCheckbox2 As String = "uAOCheckbox_cCheckbox2.bmp" ' Chico
' * Configuraci�n de estilo de controles


Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'Public LastTexture As Long
Public PixelOffsetX As Integer
Public PixelOffsetY As Integer

'Objetos p�blicos
Public Light_Cuadrada As New clsLightsCuadradas
Public Light_Redonda As New clsLightsRedondas
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New clsDialogs
Public Audio As New clsAudio
Public Inventario As New clsGraphicalInventory
Public Spells As New clsGraphicalSpells
Public InvBanco(1) As New clsGraphicalInventory

'Configuracion Especial
Public CfgSiempreNombres As Boolean
Public CfgDiaNoche       As Boolean
Public CfgSistemaLuces   As Boolean
Public ClMeditarRapido   As Boolean

'Inventarios de comercio con usuario
Public InvComUsu As New clsGraphicalInventory ' Inventario del usuario visible en el comercio
Public InvOroComUsu(2) As New clsGraphicalInventory ' Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As New clsGraphicalInventory ' Inventarios de ofertas (ambos usuarios)

Public InvComNpc As New clsGraphicalInventory ' Inventario con los items que ofrece el npc

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 4
Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As New clsGraphicalInventory
Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As New clsGraphicalInventory
                
Public SurfaceDB As clsSurfaceManager   'No va new porque es una interfaz, el new se pone al decidir que clase de objeto es
Public CustomKeys As New clsCustomKeys
Public CustomMessages As New clsCustomMessages

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

'Error code
Public Const TOO_FAST As Long = 24036
Public Const REFUSED As Long = 24061
Public Const TIME_OUT As Long = 24060
Public Const NO_AVAILIBLE As Long = 24049 ' GSZAO
Public Const ABORTED_OR_OTHER As Long = 24053 ' GSZAO

''
'The main timer of the game.
Public MainTimer As New clsTimer

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

'Mover el arma cuando ataca!
Public EstaAtacando As Boolean

' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2788
Public Const INT_MACRO_TRABAJO As Integer = 900
Public Const INT_SEND_PING As Integer = 200 ' GSZAO

Public Const INT_ATTACK As Integer = 1500
Public Const INT_ARROWS As Integer = 1400
Public Const INT_CAST_SPELL As Integer = 1400
Public Const INT_CAST_ATTACK As Integer = 1000
Public Const INT_WORK As Integer = 700
Public Const INT_USEITEMU As Integer = 450
Public Const INT_USEITEMDCK As Integer = 125
Public Const INT_SENTRPU As Integer = 2000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

Public Const HUMANO_H_PRIMER_CABEZA As Integer = 1
Public Const HUMANO_H_ULTIMA_CABEZA As Integer = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables
Public Const HUMANO_H_CUERPO_DESNUDO As Integer = 21

Public Const ELFO_H_PRIMER_CABEZA As Integer = 101
Public Const ELFO_H_ULTIMA_CABEZA As Integer = 122
Public Const ELFO_H_CUERPO_DESNUDO As Integer = 210

Public Const DROW_H_PRIMER_CABEZA As Integer = 201
Public Const DROW_H_ULTIMA_CABEZA As Integer = 221
Public Const DROW_H_CUERPO_DESNUDO As Integer = 32

Public Const ENANO_H_PRIMER_CABEZA As Integer = 301
Public Const ENANO_H_ULTIMA_CABEZA As Integer = 319
Public Const ENANO_H_CUERPO_DESNUDO As Integer = 53

Public Const GNOMO_H_PRIMER_CABEZA As Integer = 401
Public Const GNOMO_H_ULTIMA_CABEZA As Integer = 416
Public Const GNOMO_H_CUERPO_DESNUDO As Integer = 222
'**************************************************
Public Const HUMANO_M_PRIMER_CABEZA As Integer = 70
Public Const HUMANO_M_ULTIMA_CABEZA As Integer = 89
Public Const HUMANO_M_CUERPO_DESNUDO As Integer = 39

Public Const ELFO_M_PRIMER_CABEZA As Integer = 170
Public Const ELFO_M_ULTIMA_CABEZA As Integer = 188
Public Const ELFO_M_CUERPO_DESNUDO As Integer = 259

Public Const DROW_M_PRIMER_CABEZA As Integer = 270
Public Const DROW_M_ULTIMA_CABEZA As Integer = 288
Public Const DROW_M_CUERPO_DESNUDO As Integer = 40

Public Const ENANO_M_PRIMER_CABEZA As Integer = 370
Public Const ENANO_M_ULTIMA_CABEZA As Integer = 384
Public Const ENANO_M_CUERPO_DESNUDO As Integer = 60

Public Const GNOMO_M_PRIMER_CABEZA As Integer = 470
Public Const GNOMO_M_ULTIMA_CABEZA As Integer = 484
Public Const GNOMO_M_CUERPO_DESNUDO As Integer = 260

'Musica
'Public Const MP3_Inicio As Byte = 101

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 51) As tColor ' Colores de Nick

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion
Public GraphicsFile As String 'Que graficos.ind usamos

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles

Public UsaMacro As Boolean
Public CnTd As Byte

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 30
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAX_SPELL_SLOTS As Byte = 30

Public Const INV_OFFER_SLOTS As Byte = 20
Public Const INV_GOLD_SLOTS As Byte = 1

Public Const MAXSKILLPOINTS As Byte = 100

Public Const MAXATRIBUTOS As Byte = 38

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1
Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Const FOgata As Integer = 1521


Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Cl�rigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladr�n
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Palad�n
    Hunter      'Cazador
    Worker      'Trabajador
    Pirat       'Pirata
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apu�alar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
End Enum

' GSZAO
Public Enum eAccionClick
    Matrimonio = 21
    Divorcio = 22
End Enum
' GSZAO

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLe�a = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otMochilas = 37
    otCualquiera = 1000
End Enum

Public Enum eMochilas
    Mediana = 1
    GRANDE = 2
End Enum

Public MaxInventorySlots As Byte

Public Const FundirMetal As Integer = 88

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieNewbie = &H3 ' GSZAO
    ieAtacable = &H4
    ieMuerto = &H5 ' GSZAO
End Enum

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "���La criatura fall� el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "���La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "���Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "���El usuario rechaz� el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "���Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "��Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertir�s en uno de ellos y ser�s perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "�Est�s meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "��La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "��La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "��La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "��La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "��La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "��La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "��"
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_11 As String = "�"
Public Const MENSAJE_22 As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "��Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te atac� y fall�!!"

Public Const MENSAJE_RECIBE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIBE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIBE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIBE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIBE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIBE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "��Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la v�ctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el �rbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la v�ctima..."
Public Const MENSAJE_MATROMONIO As String = "Haz click sobre tu pareja para contraer matrimonio..." ' GSZAO
Public Const MENSAJE_DIVORCIO As String = "Haz click sobre tu pareja para divorciarte..." ' GSZAO

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO As String = "te ha matado!"
Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO As String = "Has matado a"
Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO As String = "Has ganado "
Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = "puntos de experiencia."

Public Const MENSAJE_NO_VES_NADA_INTERESANTE As String = "No ves nada interesante."
Public Const MENSAJE_HAS_MATADO_A As String = "Has matado a "
Public Const MENSAJE_HAS_GANADO_EXPE_1 As String = "Has ganado "
Public Const MENSAJE_HAS_GANADO_EXPE_2 As String = " puntos de experiencia."
Public Const MENSAJE_TE_HA_MATADO As String = " te ha matado!"

Public Const MENSAJE_HOGAR As String = "Has llegado a tu hogar. El viaje ha finalizado."
Public Const MENSAJE_HOGAR_CANCEL As String = "Tu viaje ha sido cancelado."

Public Enum eMessages
    DontSeeAnything
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
End Enum

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer 'Budi
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    amount As Integer
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tItemsConstruibles
    Name As String
    OBJIndex As Integer
    GrhIndex As Integer
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Integer
End Type

Public Nombres As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory
Public UserHechizos(1 To MAX_SPELL_SLOTS) As Spells

Type Spells
    Index As Integer
    Name As String
    GrhIndex As Integer
End Type

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As eCiudad

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
Public MirandoCarpinteria As Boolean ' 0.13.3
Public MirandoHerreria As Boolean ' 0.13.3
Public bFormYesNo As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMSKILLS As Byte = 20
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 12
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Byte
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
'Public logged As Boolean ' www.gs-zone.org

Public UsingSkill As Integer

Public MD5HushYo As String * 16

Public pingTime As Long

Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
End Enum

Public EstadoLogin As E_MODO
Public CaptchaKey As Byte       ' GSZAO
Public CaptchaCode(3) As Byte   ' GSZAO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum


''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param BAJOTECHOSINNPCS bajo techo y los NPCs no pueden pisar o respawn en tiles con este trigger
' @param SINNPCS los NPCs no pueden pisar o respawn en tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE los usuarios no pueden quedarse parados en este trigger demasiado tiempo o ser�n encarcelados
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    BAJOTECHOSINNPCS = 2
    SINNPCS = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Long ' GSZAO
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Long ' GSZAO
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long ' GSZAO
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

' 1 foro general y 2 faccionarios
Public Foros(0 To 2) As tForo

' Forum info handler
Public clsForos As New clsForum

Public isCapturePending As Boolean
Public Traveling As Boolean

Public bShowGuildNews As Boolean ' 0.13.3
Public GuildNames() As String
Public GuildMembers() As String

Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sResucitation
    sSafemode
    mSpells
    mWork
End Enum

'FragShooter variables ' 0.13.3
Public FragShooterCapturePending As Boolean
Public FragShooterNickname As String
Public FragShooterKilledSomeone As Boolean

Public Const SM_CANT As Byte = 4
Public SMStatus(SM_CANT) As Boolean

'Hardcoded grhs and items
Public Const GRH_INI_SM As Integer = 4978

Public Const ORO_INDEX As Integer = 12
Public Const ORO_GRH As Integer = 511

Public Const GRH_HALF_STAR As Integer = 5357
Public Const GRH_FULL_STAR As Integer = 5358
Public Const GRH_GLOW_STAR As Integer = 5359

Public Const LH_GRH As Integer = 724
Public Const LP_GRH As Integer = 725
Public Const LO_GRH As Integer = 723

Public Const MADERA_GRH As Integer = 550
Public Const MADERA_ELFICA_GRH As Integer = 1999

Public Const GrhFogata As Integer = 1521


Public picMouseIcon As Picture

' 0.13.3
Public Enum eMoveType
    Inventory = 1
    Bank = 2
    SpellsI = 3
End Enum

' 0.13.3
Public Const MP3_INITIAL_INDEX As Integer = 1000

' GSZAO
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessages Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

