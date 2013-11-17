Attribute VB_Name = "modVariables"
'                       ######################################
'                       ############  FRoG Creator 1.0   ###########
'                       ##  Module de stockage des variables globales  ##
'                       ##### Dernière modification : JJ/MM/AAAA  #####
'                       ######################################


' -- Réseau --
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public MyIndex As Long

' -- Configuration du jeu --
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public MAX_DX_PETS As Long
Public MAX_PETS As Long
Public MAX_METIER As Long
Public MAX_CLASSES As Byte
Public MAX_MAPX As Long
Public MAX_MAPY As Long

' -- Variables d'objets --
Public quete() As QueteRec
Public Map() As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public PlayerAnim() As Long
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public CoffreTmp(1 To 30) As CoffreTempRec
Public Pets() As PetsRec

' -- Graphique --
Public ExtraSheets As Long
Public dX As New DirectX7
Public DD As DirectDraw7
Public D3D As Direct3D7
Public Dev As Direct3DDevice7
Public DD_Clip As DirectDrawClipper
Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2
Public DD_SpriteSurf() As DirectDrawSurface7
Public DDSD_Character() As DDSURFACEDESC2
Public SpriteTimer() As Long
Public SpriteUsed() As Boolean
Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2
Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2
Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2
Public DD_SpellAnim() As DirectDrawSurface7
Public DDSD_SpellAnim() As DDSURFACEDESC2
Public SpellTimer() As Long
Public SpellUsed() As Boolean
Public DD_BigSpellAnim() As DirectDrawSurface7
Public DDSD_BigSpellAnim() As DDSURFACEDESC2
Public BigSpellTimer() As Long
Public BigSpellUsed() As Boolean
Public DD_TileSurf() As DirectDrawSurface7
Public DDSD_Tile() As DDSURFACEDESC2
Public TileFile() As Boolean
Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7
Public DDSD_Outil As DDSURFACEDESC2
Public DD_OutilSurf As DirectDrawSurface7
Public DD_PaperDollSurf() As DirectDrawSurface7
Public DDSD_PaperDoll() As DDSURFACEDESC2
Public PaperDollTimer() As Long
Public PaperDollUsed() As Boolean
Public DD_PetsSurf() As DirectDrawSurface7
Public DDSD_Pets() As DDSURFACEDESC2
Public PetTimer() As Long
Public PetUsed() As Boolean
Public DDSD_Blood As DDSURFACEDESC2
Public DD_Blood As DirectDrawSurface7
Public DDSD_PanoInf As DDSURFACEDESC2
Public DD_PanoInfSurf As DirectDrawSurface7
Public DDSD_PanoSup As DDSURFACEDESC2
Public DD_PanoSupSurf As DirectDrawSurface7
Public DDSD_Night As DDSURFACEDESC2
Public DD_NightSurf As DirectDrawSurface7
Public NightVerts(3) As D3DTLVERTEX
Public DDSD_Fog As DDSURFACEDESC2
Public DD_FogSurf As DirectDrawSurface7
Public FogVerts(3) As D3DTLVERTEX
Public DDSD_Tmp As DDSURFACEDESC2
Public DD_TmpSurf As DirectDrawSurface7
Public rec As RECT
Public rec_pos As RECT
Public AlphaBlendDXIsInit As Boolean
Public ABDXWidth As Integer
Public ABDXHeight As Integer
Public ABDXAlpha As Single
Public Const SurfaceTimerMax As Long = 30000

' -- Logique --
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean
Public CHECK_WAIT As Boolean
Public MyText As String
Public MapAnim As Boolean
Public MapAnimTimer As Long
Public GettingMap As Boolean
Public InToit As Boolean
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public GameFPS As Long
Public CurX As Single '/case
Public CurY As Single '/case
Public PotX As Single 'réel
Public PotY As Single 'réel
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long
Public NewPlayerPicX As Long
Public NewPlayerPicY As Long
Public NewPlayerPOffsetX As Long
Public NewPlayerPOffsetY As Long
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long
Public DmgAddRem As Long
Public NPCDmgAddRem As Long
Public ii As Long, iii As Long
Public sx As Long
Public sy As Long
Public MouseDownX As Long
Public MouseDownY As Long
Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long
Public SoundFileName As String
Public Connucted As Boolean
Public bankmsg As String
Public Accepter As Boolean
Public ConOff As Boolean
Public OldMap As Long
Public Rep_Theme As String
Public NumShop As Long
Public drx As Long
Public dry As Long
Public dr As Boolean
Public cychat As Integer
Public AccModo As Long
Public AccMapeur As Long
Public AccDevelopeur As Long
Public AccAdmin As Long
Public PNJAnim(1 To MAX_MAP_NPCS) As Byte
Public PicScWidth As Single
Public PicScHeight As Single
Public MaxSprite As Integer
Public MaxPaperdoll As Integer
Public MaxSpell As Integer
Public MaxBigSpell As Integer
Public MaxPet As Integer

' -- Divers --
Public RecetteSelect As Byte
Public Loading As Boolean
Public deco As Boolean
Public netbook As Boolean
