Attribute VB_Name = "modTypes"
Option Explicit

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Byte
Public MAX_SPELLS As Integer
Public MAX_MAPS As Integer
Public MAX_SHOPS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_MAP_ITEMS As Byte
Public MAX_EMOTICONS As Byte
Public MAX_SPELL_ANIM As Byte
Public MAX_BLT_LINE As Byte
Public MAX_LEVEL As Integer
Public MAX_QUETES As Integer
Public MAX_DX_PETS As Byte
Public MAX_PETS As Integer
Public MAX_METIER As Integer
Public MAX_RECETTE As Integer
Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public BLT_SNOW_DROPS As Long

Public Const MAX_ARROWS As Byte = 100
Public Const MAX_PLAYER_ARROWS As Byte = 100

Public MAX_INV As Integer
Public Const MAX_PARTY_MEMBERS As Byte = 20
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_DATA_METIER = 100

Public Const NO As Byte = 0
Public Const YES As Byte = 1

Public RecetteSelect As Byte

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 As String = "aqcashlhriyjjmbiklsqzzjdiazqgiawaivwvilzftnysppcvglemckghmqqzfhbnfqwtgnnpafrvnxatftqncgnbwbbfnjswgrtxqwnltdnertceivfcnqzbjt"
Public Const SEC_CODE2 As String = "digshuxirmautdxdsdtlmwckaalubgjmmauqhrmgxxtlgcbenzregecdawwviryxcpckckxbregphfaregjinrxanwmtdmhluhfrdivayqhpdmmaqkqjqaybpayct"
Public Const SEC_CODE3 As String = "thumqnewytvtctwktdnzsitkecsnlcwihrelzxnbsdluhucqspsjlmwbbpjabfwzjechdkskzsxzasdsxejytcudtfpyefrugwnhvvcfbkwigmsfeywjvpf"
Public Const SEC_CODE4 As String = "58389610143670529438361696763476787278903650107818303274347098703634903098149832927278741812909214565096961"

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Byte
Public MAX_MAPY As Byte
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32
Public PIC_PL As Byte
Public PIC_NPC1 As Byte
Public PIC_NPC2 As Byte

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_CBLOCK As Byte = 10
Public Const TILE_TYPE_ARENA As Byte = 11
Public Const TILE_TYPE_SOUND As Byte = 12
Public Const TILE_TYPE_SPRITE_CHANGE As Byte = 13
Public Const TILE_TYPE_SIGN As Byte = 14
Public Const TILE_TYPE_DOOR As Byte = 15
Public Const TILE_TYPE_NOTICE As Byte = 16
Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NPC_SPAWN As Byte = 20
Public Const TILE_TYPE_BANK As Byte = 21
Public Const TILE_TYPE_POISON As Byte = 26
Public Const TILE_TYPE_COFFRE As Byte = 22
Public Const TILE_TYPE_PORTE_CODE As Byte = 23
Public Const TILE_TYPE_BLOCK_MONTURE As Byte = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX As Byte = 25
Public Const TILE_TYPE_TOIT As Byte = 26
Public Const TILE_TYPE_BLOCK_GUILDE As Byte = 27
Public Const TILE_TYPE_BLOCK_TOIT As Byte = 28
Public Const TILE_TYPE_BLOCK_DIR As Byte = 29
Public Const TILE_TYPE_CRAFT As Byte = 30
Public Const TILE_TYPE_METIER As Byte = 31

' quetes constant
Public Const QUETE_TYPE_AUCUN As Byte = 0
Public Const QUETE_TYPE_RECUP As Byte = 1
Public Const QUETE_TYPE_APORT As Byte = 2
Public Const QUETE_TYPE_PARLER As Byte = 3
Public Const QUETE_TYPE_TUER As Byte = 4
Public Const QUETE_TYPE_FINIR As Byte = 5
Public Const QUETE_TYPE_GAGNE_XP As Byte = 6
Public Const QUETE_TYPE_SCRIPT As Byte = 7
Public Const QUETE_TYPE_MINIQUETE As Byte = 8

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13
Public Const ITEM_TYPE_MONTURE As Byte = 14
Public Const ITEM_TYPE_SCRIPT As Byte = 15
Public Const ITEM_TYPE_PET As Byte = 16

Public Const ITEM_TYPEARME_NONE As Byte = 0
Public Const ITEM_TYPEARME_EPEES As Byte = 1
Public Const ITEM_TYPEARME_HACHES As Byte = 2
Public Const ITEM_TYPEARME_DAGUES As Byte = 3
Public Const ITEM_TYPEARME_FAUX As Byte = 4
Public Const ITEM_TYPEARME_MARTEAUX As Byte = 5
Public Const ITEM_TYPEARME_PIOCHES As Byte = 6
Public Const ITEM_TYPEARME_PELLES As Byte = 7
Public Const ITEM_TYPEARME_BATONS As Byte = 8
Public Const ITEM_TYPEARME_BAGUETTES As Byte = 9
Public Const ITEM_TYPEARME_OUTILLAGE As Byte = 10
Public Const ITEM_TYPEARME_ARC As Byte = 11

' Metier
Public Const METIER_CHASSEUR As Byte = 0
Public Const METIER_CRAFT As Byte = 1

' Direction constants
Public Const DIR_UP As Byte = 3
Public Const DIR_DOWN As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2
Public Const MOVING_VEHICUL As Byte = 3

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3

' Time constants
Public Const TIME_DAY As Byte = 0
Public Const TIME_NIGHT As Byte = 1

' Admin constants
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4
Public Const NPC_BEHAVIOR_QUETEUR As Byte = 5

' Speach bubble constants
Public DISPLAY_BUBBLE_TIME As Long ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 16 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 20 ' In characters.
Public Const MAX_LINES As Byte = 5

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
Public Const SPELL_TYPE_SCRIPT As Byte = 6
Public Const SPELL_TYPE_AMELIO As Byte = 7
Public Const SPELL_TYPE_DECONC As Byte = 8
Public Const SPELL_TYPE_PARALY As Byte = 9
Public Const SPELL_TYPE_DEFENC As Byte = 10

Public Loading As Boolean
Public deco As Boolean

Public netbook As Boolean

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Integer
    Value As Integer
    dur As Long
End Type

Type CoffreTempRec
    Numeros As Integer
    Valeur As Integer
    Durabiliter As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    
    Target As Byte
    TargetType As Byte
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Byte
    ArrowAnim As Byte
    ArrowTime As Long
    ArrowVarX As Byte
    ArrowVarY As Byte
    ArrowX As Byte
    ArrowY As Byte
    ArrowPosition As Byte
End Type

Type IndRec
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
End Type

Type PlayerQueteRec
    Temps As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    indexe(1 To 15) As IndRec
End Type

Type PetPosRec
    x As Byte
    y As Byte
    Dir As Byte
    XOffset As Byte
    YOffset As Byte
    Anim As Byte
End Type

Type PlayerRec

    ' General
    name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Byte
    Sprite As Byte
    level As Integer
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Integer
    DEF As Integer
    speed As Integer
    MAGI As Integer
    POINTS As Integer
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    PetSlot As Byte
    
    ' Inventory
    Inv() As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    pet As PetPosRec
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte

    ' Client use only
    MaxHp As Long
    MaxMp As Long
    MaxSP As Long
    XOffset As Byte
    YOffset As Byte
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    PartyIndex As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec
    BloodAnim As SpellAnimRec

    EmoticonNum As Byte
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
    QueteEnCour As Integer
    Quetep As PlayerQueteRec
    
    Anim As Byte
    'PAPERDOLL
    Casque As Integer
    Armure As Integer
    Arme As Integer
    Bouclier As Integer
    'FIN PAPERDOLL

    Metier As Integer
    MetierLvl As Integer
    MetierExp As Integer
End Type
    
Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    M2Anim As Integer
    Mask3 As Integer
    M3Anim As Integer
    Fringe As Integer
    FAnim As Integer
    Fringe2 As Integer
    F2Anim As Integer
    Fringe3 As Integer
    F3Anim As Integer
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    Mask3Set As Byte
    M3AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
    Fringe3Set As Byte
    F3AnimSet As Byte
End Type

Type NpcMapRec
    x As Byte
    y As Byte
    x1 As Byte
    y1 As Byte
    x2 As Byte
    y2 As Byte
    x3 As Byte
    y3 As Byte
    x4 As Byte
    y4 As Byte
    x5 As Byte
    y5 As Byte
    x6 As Byte
    y6 As Byte
    boucle As Byte
    Hasardm As Byte
    Hasardp As Byte
    Imobile As Byte
    Axy As Boolean
    Axy1 As Boolean
    Axy2 As Boolean
End Type

Type MapRec
    name As String * 40
    Revision As Integer
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Integer
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Npcs(1 To MAX_MAP_NPCS) As NpcMapRec
    PanoInf As String * 50
    TranInf As Byte
    PanoSup As String * 50
    TranSup As Byte
    Fog As Integer
    FogAlpha As Byte
    guildSoloView As Byte
    petView As Byte
    traversable As Byte
    meteo As Byte
    frequenceMeteo As Byte
End Type

Type RecompRec
    exp As Long
    objn1 As Integer
    objn2 As Integer
    objn3 As Integer
    objq1 As Integer
    objq2 As Integer
    objq3 As Integer
End Type

Type QueteRec
    nom As String * 40
    Type As Byte
    description As String
    reponse As String
    Temps As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    Recompence As RecompRec
    indexe(1 To 15) As IndRec
    Case As Long
End Type

Type ClassRec
    name As String * NAME_LENGTH
    MaleSprite As Integer
    FemaleSprite As Integer
    
    Locked As Byte
    
    STR As Integer
    DEF As Integer
    speed As Integer
    MAGI As Integer
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    name As String * NAME_LENGTH
    desc As String * 150
    
    Pic As Integer
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Integer
    DefReq As Integer
    SpeedReq As Integer
    ClassReq As Byte
    AccessReq As Byte
    LevelReq As Integer
    
    paperdoll As Byte
    paperdollPic As Integer
    
    Empilable As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Integer
    AddDef As Integer
    AddMagi As Integer
    AddSpeed As Integer
    AddEXP As Long
    AttackSpeed As Integer
    
    NCoul As Long
    tArme As Long
End Type

Type MapItemRec
    num As Integer
    Value As Integer
    dur As Integer
    
    x As Byte
    y As Byte
End Type

Type NPCEditorRec
    Itemnum As Integer
    ItemValue As Integer
    Chance As Long
End Type

Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Integer
    DEF As Integer
    speed As Integer
    MAGI As Integer
    MaxHp As Long
    exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    QueteNum As Integer
    Inv As Integer
    Vol As Byte
End Type

Type MapNpcRec
    num As Integer
    
    Target As Byte
    
    HP As Long
    MaxHp As Long
    MP As Long
    MaxMp As Long
    SP As Long
    
    Map As Integer
    x As Byte
    y As Byte
    x2 As Byte
    y2 As Byte
    x3 As Byte
    y3 As Byte
    x4 As Byte
    y4 As Byte
    x5 As Byte
    y5 As Byte
    x6 As Byte
    y6 As Byte
    Dir As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Integer
    GiveValue As Integer
    GetItem As Integer
    getValue As Integer
End Type

Type TradeItemsRec
    Value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
    FixObjet As Long
End Type

Type SpellRec
    name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Integer
    Sound As Long
    MPCost As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    
    Big As Byte
    
    SpellAnim As Integer
    SpellTime As Long
    SpellDone As Long
    
    SpellIco As Integer
    
    AE As Long
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Byte
    InvName As String
    InvVal As Integer
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Integer
    Command As String
End Type

Type DropRainRec
    x As Byte
    y As Byte
    Randomized As Boolean
    speed As Byte
End Type

Type PetsRec
    nom As String
    Sprite As Integer
    addForce As Byte
    addDefence As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

Public Max_Classes As Byte
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
Public recette() As RecetteRec
Public DropRain() As DropRainRec
Public DropSnow() As DropRainRec

Type ItemTradeRec
    ItemGetNum As Integer
    ItemGiveNum As Integer
    ItemGetVal As Integer
    ItemGiveVal As Integer
End Type

Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Byte
    SelectedItem As Integer
End Type
Public Trade(1 To 6) As TradeRec

Type ArrowRec
    name As String
    Pic As Integer
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    Done As Byte
    y As Integer
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Integer
    dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long

Public Minu As Long
Public Seco As Long

'Type pour stocker le contenu de Account.ini
Type TpAccOpt
    InfName As String
    InfPass As String
    SpeechBubbles As Boolean
    NpcBar As Boolean
    NpcName As Boolean
    NpcDamage As Boolean
    PlayBar As Boolean
    PlayName As Boolean
    PlayDamage As Boolean
    MapGrid As Boolean
    Music As Boolean
    Sound As Boolean
    Autoscroll As Boolean
    NomObjet As Boolean
    LowEffect As Boolean
End Type

Public rac(0 To 13, 0 To 1) As String
Public dragAndDrop As Byte
Public dragAndDropT As Byte

Public AccOpt As TpAccOpt

Type MetierRec
    nom As String
    Type As Byte
    desc As String
    
    data(0 To MAX_DATA_METIER, 0 To 1) As Integer
End Type
Public Metier() As MetierRec

Type RecetteRec
    nom As String
    InCraft(0 To 9, 0 To 1) As Integer
    craft(0 To 1) As Integer
End Type

' Configuration Menu Option des touches
Type optToucheRec
    nom As String
    Value As Byte
End Type
Public nelvl As Byte
Public Const TCHMAX = 51
Public optTouche(0 To TCHMAX) As optToucheRec

Type charSelectRec
    name As String
    classe As String
    level As Integer
    sprt As Long
End Type
Public charSelect(1 To MAX_CHARS) As charSelectRec
Public charSelectNum As Byte

Sub iniOptTouche()
    optTouche(0).nom = "A"
    optTouche(0).Value = vbKeyA
    optTouche(1).nom = "B"
    optTouche(1).Value = vbKeyB
    optTouche(2).nom = "C"
    optTouche(2).Value = vbKeyC
    optTouche(3).nom = "D"
    optTouche(3).Value = vbKeyD
    optTouche(4).nom = "E"
    optTouche(4).Value = vbKeyE
    optTouche(5).nom = "F"
    optTouche(5).Value = vbKeyF
    optTouche(6).nom = "G"
    optTouche(6).Value = vbKeyG
    optTouche(7).nom = "H"
    optTouche(7).Value = vbKeyH
    optTouche(8).nom = "I"
    optTouche(8).Value = vbKeyI
    optTouche(9).nom = "J"
    optTouche(9).Value = vbKeyJ
    optTouche(10).nom = "K"
    optTouche(10).Value = vbKeyK
    optTouche(11).nom = "L"
    optTouche(11).Value = vbKeyL
    optTouche(12).nom = "M"
    optTouche(12).Value = vbKeyM
    optTouche(13).nom = "N"
    optTouche(13).Value = vbKeyN
    optTouche(14).nom = "O"
    optTouche(14).Value = vbKeyO
    optTouche(15).nom = "P"
    optTouche(15).Value = vbKeyP
    optTouche(16).nom = "Q"
    optTouche(16).Value = vbKeyQ
    optTouche(17).nom = "R"
    optTouche(17).Value = vbKeyR
    optTouche(18).nom = "S"
    optTouche(18).Value = vbKeyS
    optTouche(19).nom = "T"
    optTouche(19).Value = vbKeyT
    optTouche(20).nom = "U"
    optTouche(20).Value = vbKeyU
    optTouche(21).nom = "V"
    optTouche(21).Value = vbKeyV
    optTouche(22).nom = "W"
    optTouche(22).Value = vbKeyW
    optTouche(23).nom = "X"
    optTouche(23).Value = vbKeyX
    optTouche(24).nom = "Y"
    optTouche(24).Value = vbKeyY
    optTouche(25).nom = "Z"
    optTouche(25).Value = vbKeyZ
    optTouche(26).nom = "0"
    optTouche(26).Value = vbKey0
    optTouche(27).nom = "1"
    optTouche(27).Value = vbKey1
    optTouche(28).nom = "2"
    optTouche(28).Value = vbKey2
    optTouche(29).nom = "3"
    optTouche(29).Value = vbKey3
    optTouche(30).nom = "4"
    optTouche(30).Value = vbKey4
    optTouche(31).nom = "5"
    optTouche(31).Value = vbKey5
    optTouche(32).nom = "6"
    optTouche(32).Value = vbKey6
    optTouche(33).nom = "7"
    optTouche(33).Value = vbKey7
    optTouche(34).nom = "8"
    optTouche(34).Value = vbKey8
    optTouche(35).nom = "9"
    optTouche(35).Value = vbKey9
    optTouche(36).nom = "F1"
    optTouche(36).Value = vbKeyF1
    optTouche(37).nom = "F2"
    optTouche(37).Value = vbKeyF2
    optTouche(38).nom = "F3"
    optTouche(38).Value = vbKeyF3
    optTouche(39).nom = "F4"
    optTouche(39).Value = vbKeyF4
    optTouche(40).nom = "F5"
    optTouche(40).Value = vbKeyF5
    optTouche(41).nom = "F6"
    optTouche(41).Value = vbKeyF6
    optTouche(42).nom = "F7"
    optTouche(42).Value = vbKeyF7
    optTouche(43).nom = "F8"
    optTouche(43).Value = vbKeyF8
    optTouche(44).nom = "Haut"
    optTouche(44).Value = vbKeyUp
    optTouche(45).nom = "Bas"
    optTouche(45).Value = vbKeyDown
    optTouche(46).nom = "Gauche"
    optTouche(46).Value = vbKeyLeft
    optTouche(47).nom = "Droite"
    optTouche(47).Value = vbKeyRight
    optTouche(48).nom = "Ctrl"
    optTouche(48).Value = vbKeyControl
    optTouche(49).nom = "Alt"
    optTouche(49).Value = vbKeyMenu
    optTouche(50).nom = "Shift"
    optTouche(50).Value = vbKeyShift
    optTouche(51).nom = "Espace"
    optTouche(51).Value = vbKeySpace
    
    
End Sub

Sub ClearTempTile()
Dim x As Byte, y As Byte

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Byte)
Dim i As Byte, n As Byte

With Player(Index)
    .name = vbNullString
    .Guild = vbNullString
    .Guildaccess = 0
    .Class = 0
    .level = 0
    .Sprite = 0
    .exp = 0
    .Access = 0
    .PK = NO
        
    .HP = 0
    .MP = 0
    .SP = 0
        
    .STR = 0
    .DEF = 0
    .speed = 0
    .MAGI = 0
    
    .QueteEnCour = 0
    .Quetep.Data1 = 0
    .Quetep.Data2 = 0
    .Quetep.Data3 = 0
    .Quetep.String1 = vbNullString
      
    For n = 1 To 15
    .Quetep.indexe(n).Data1 = 0
    .Quetep.indexe(n).Data2 = 0
    .Quetep.indexe(n).Data3 = 0
    .Quetep.indexe(n).String1 = vbNullString
    Next n
        
    For n = 1 To MAX_INV
        .Inv(n).num = 0
        .Inv(n).Value = 0
        .Inv(n).dur = 0
    Next n
        
    .ArmorSlot = 0
    .WeaponSlot = 0
    .HelmetSlot = 0
    .ShieldSlot = 0
    .PetSlot = 0
    
    .Map = 0
    .x = 0
    .y = 0
    .Dir = 0
    
    .pet.Dir = DIR_DOWN
    .pet.y = 1
    .pet.y = 1
    
    ' Client use only
    .MaxHp = 0
    .MaxMp = 0
    .MaxSP = 0
    .XOffset = 0
    .YOffset = 0
    .Moving = 0
    .Attacking = 0
    .AttackTimer = 0
    .MapGetTimer = 0
    .CastedSpell = NO
    .EmoticonNum = -1
    .EmoticonTime = 0
    .EmoticonVar = 0
    
    For i = 1 To MAX_SPELL_ANIM
        .SpellAnim(i).CastedSpell = NO
        .SpellAnim(i).SpellTime = 0
        .SpellAnim(i).SpellVar = 0
        .SpellAnim(i).SpellDone = 0
        
        .SpellAnim(i).Target = 0
        .SpellAnim(i).TargetType = 0
    Next i
    
    .SpellNum = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    .QueteEnCour = 0
    
    Inventory = 1
End With
End Sub

Sub ClearPlayerQuete(ByVal Index As Byte)
Dim i As Byte
With Player(Index)
        .QueteEnCour = 0
        .Quetep.Data1 = 0
        .Quetep.Data2 = 0
        .Quetep.Data3 = 0
        .Quetep.String1 = vbNullString
        Accepter = False
        
        For i = 1 To 15
        .Quetep.indexe(i).Data1 = 0
        .Quetep.indexe(i).Data2 = 0
        .Quetep.indexe(i).Data3 = 0
        .Quetep.indexe(i).String1 = 0
        Next i
End With
End Sub

Sub ClearItem(ByVal Itemnum As Integer)
With Item(Itemnum)
    .name = vbNullString
    .desc = vbNullString
    
    .Type = 0
    .Data1 = 0
    .Data2 = 0
    .Data3 = 0
    .StrReq = 0
    .DefReq = 0
    .SpeedReq = 0
    .ClassReq = -1
    .AccessReq = 0
    
    .paperdoll = 0
    .paperdollPic = 0
    
    .Empilable = 0
    
    .AddHP = 0
    .AddMP = 0
    .AddSP = 0
    .AddStr = 0
    .AddDef = 0
    .AddMagi = 0
    .AddSpeed = 0
    .AddEXP = 0
    .AttackSpeed = 1000
    
    .NCoul = 0
    .tArme = 0
End With
End Sub

Sub ClearItems()
Dim i As Integer

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal MapNum As Integer)
With MapItem(MapNum)
    .num = 0
    .Value = 0
    .dur = 0
    .x = 0
    .y = 0
End With
End Sub

Sub ClearMaps()
Dim i As Integer
Dim x As Byte
Dim y As Byte

For i = 1 To MAX_MAPS
With Map(i)
    .name = vbNullString
    .Revision = 0
    .Moral = 0
    .Up = 0
    .Down = 0
    .Left = 0
    .Right = 0
    .Indoors = 0
    .meteo = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With .Tile(x, y)
            .Ground = 0
            .Mask = 0
            .Anim = 0
            .Mask2 = 0
            .M2Anim = 0
            .Mask3 = 0
            .M3Anim = 0
            .Fringe = 0
            .FAnim = 0
            .Fringe2 = 0
            .F2Anim = 0
            .Fringe3 = 0
            .F3Anim = 0
            .Type = 0
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            .String1 = vbNullString
            .String2 = vbNullString
            .String3 = vbNullString
            .Light = 0
            .GroundSet = 0
            .MaskSet = 0
            .AnimSet = 0
            .Mask2Set = 0
            .M2AnimSet = 0
            .Mask3Set = 0
            .M3AnimSet = 0
            .FringeSet = 0
            .FAnimSet = 0
            .Fringe2Set = 0
            .F2AnimSet = 0
            .Fringe3Set = 0
            .F3AnimSet = 0
            End With
        Next x
    Next y
    .PanoInf = vbNullString
    .TranInf = 0
    .PanoSup = vbNullString
    .TranSup = 0
    .Fog = 0
    .FogAlpha = 0
    .guildSoloView = 0
    .petView = 0
    .traversable = 0
End With
Next i
End Sub

Sub ClearMapItems()
Dim x As Byte

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapNpcs()
Dim i As Byte

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .num = 0
            .Target = 0
            .HP = 0
            .MP = 0
            .SP = 0
            .Map = 0
            .x = 0
            .y = 0
            .Dir = 0
    
             ' Client use only
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
            .Attacking = 0
            .AttackTimer = 0
        End With
        PNJAnim(i) = 1
    Next i
End Sub

Function GetPlayerName(ByVal Index As Byte) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Byte, ByVal name As String)
    Player(Index).name = name
End Sub

Function GetPlayerGuild(ByVal Index As Byte) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Byte, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Byte) As Byte
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Byte, ByVal Guildaccess As Byte)
    Player(Index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Byte) As Byte
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Byte, ByVal ClassNum As Byte)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Byte) As Integer
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Byte, ByVal Sprite As Integer)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Byte) As Integer
    GetPlayerLevel = Player(Index).level
End Function

Sub SetPlayerLevel(ByVal Index As Byte, ByVal level As Integer)
    Player(Index).level = level
End Sub

Function GetPlayerExp(ByVal Index As Byte) As Long
    GetPlayerExp = Player(Index).exp
End Function

Sub SetPlayerExp(ByVal Index As Byte, ByVal exp As Long)
    Player(Index).exp = exp
End Sub

Function GetPlayerAccess(ByVal Index As Byte) As Byte
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Byte, ByVal Access As Byte)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Byte) As Integer
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Integer)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Byte) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Byte, ByVal HP As Long)
    Player(Index).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerMP(ByVal Index As Byte) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Byte, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then Player(Index).MP = GetPlayerMaxMP(Index)
End Sub

Function GetPlayerSP(ByVal Index As Byte) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Byte, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then Player(Index).SP = GetPlayerMaxSP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Byte) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal Index As Byte) As Long
    GetPlayerMaxMP = Player(Index).MaxMp
End Function

Function GetPlayerMaxSP(ByVal Index As Byte) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerstr(ByVal Index As Byte) As Integer
    GetPlayerstr = Player(Index).STR
End Function

Sub SetPlayerstr(ByVal Index As Byte, ByVal STR As Integer)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Byte) As Integer
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Byte, ByVal DEF As Integer)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Byte) As Integer
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Byte, ByVal speed As Integer)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Byte) As Integer
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Byte, ByVal MAGI As Integer)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Byte) As Integer
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Byte, ByVal POINTS As Integer)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Byte) As Integer
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Byte, ByVal MapNum As Integer)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Byte) As Byte
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Byte, ByVal x As Byte)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Byte) As Byte
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Byte, ByVal y As Byte)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Byte) As Byte
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Byte, ByVal Dir As Byte)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Byte, ByVal InvSlot As Byte) As Integer
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Byte, ByVal InvSlot As Byte, ByVal Itemnum As Integer)
    Player(Index).Inv(InvSlot).num = Itemnum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Byte, ByVal InvSlot As Byte) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Byte, ByVal InvSlot As Byte, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Byte, ByVal InvSlot As Byte) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Byte, ByVal InvSlot As Byte, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Byte) As Byte
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Byte, InvNum As Byte)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Byte) As Byte
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Byte, InvNum As Byte)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Byte) As Byte
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Byte, InvNum As Byte)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Byte) As Byte
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Byte, InvNum As Byte)
    Player(Index).ShieldSlot = InvNum
End Sub

Sub ClearPet(ByVal Index As Integer)
    Pets(Index).nom = ""
    Pets(Index).Sprite = 0
    Pets(Index).addForce = 0
    Pets(Index).addDefence = 0
End Sub

Function GetPlayerPetSlot(ByVal Index As Byte) As Byte
    GetPlayerPetSlot = Player(Index).PetSlot
End Function

Sub SetPlayerPetSlot(ByVal Index As Byte, InvNum As Byte)
    Player(Index).PetSlot = InvNum
End Sub
