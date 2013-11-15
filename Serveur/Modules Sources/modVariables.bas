Attribute VB_Name = "modVariables"
' - Variables de réseau
Public GAME_PORT As Integer

' - Variables maximales
Public GAME_NAME As String
Public MAX_PLAYERS As Byte
Public MAX_SPELLS As Integer
Public MAX_MAPS As Integer
Public MAX_SHOPS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_MAP_ITEMS As Integer
Public MAX_GUILDS As Integer
Public MAX_GUILD_MEMBERS As Integer
Public MAX_EMOTICONS As Byte
Public MAX_LEVEL As Integer
Public MAX_QUETES As Integer
Public Scripting As Boolean
Public NOOB_LEVEL As Integer
Public PK_LEVEL As Integer
Public RATE_EXP As Byte
Public RATE_QUETE As Byte
Public RATE_MAX As Byte
Public MAX_PETS As Integer
Public MAX_METIER As Integer
Public MAX_RECETTE As Integer

' - Constantes maximales
Public Const MAX_PARTY_MEMBERS As Byte = 20
Public Const MAX_PARTYS As Byte = 20
Public Const MAX_HDV_TRADES As Byte = 5
Public Const MAX_ARROWS As Byte = 100
Public Const MAX_INV = 26
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_NPC_SPELLS = 10
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' - Clef de sécurité (necessaire à la connexion du client au serveur)
Public Const SEC_CODE1 = "aqcashlhriyjjmbiklsqzzjdiazqgiawaivwvilzftnysppcvglemckghmqqzfhbnfqwtgnnpafrvnxatftqncgnbwbbfnjswgrtxqwnltdnertceivfcnqzbjt"
Public Const SEC_CODE2 = "digshuxirmautdxdsdtlmwckaalubgjmmauqhrmgxxtlgcbenzregecdawwviryxcpckckxbregphfaregjinrxanwmtdmhluhfrdivayqhpdmmaqkqjqaybpayct"
Public Const SEC_CODE3 = "thumqnewytvtctwktdnzsitkecsnlcwihrelzxnbsdluhucqspsjlmwbbpjabfwzjechdkskzsxzasdsxejytcudtfpyefrugwnhvvcfbkwigmsfeywjvpf"
Public Const SEC_CODE4 = "58389610143670529438361696763476787278903650107818303274347098703634903098149832927278741812909214565096961"

' - Constantes des maps
Public MAX_MAPX As Byte
Public MAX_MAPY As Byte
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2

' - Attributs des maps
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_HEAL = 7
Public Const TILE_TYPE_KILL = 8
Public Const TILE_TYPE_SHOP = 9
Public Const TILE_TYPE_CBLOCK = 10
Public Const TILE_TYPE_ARENA = 11
Public Const TILE_TYPE_SOUND = 12
Public Const TILE_TYPE_SPRITE_CHANGE = 13
Public Const TILE_TYPE_SIGN = 14
Public Const TILE_TYPE_DOOR = 15
Public Const TILE_TYPE_NOTICE = 16
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
Public Const TILE_TYPE_NPC_SPAWN = 20
Public Const TILE_TYPE_BANK = 21
Public Const TILE_TYPE_COFFRE = 22
Public Const TILE_TYPE_PORTE_CODE = 23
Public Const TILE_TYPE_BLOCK_MONTURE = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX = 25
Public Const TILE_TYPE_TOIT = 26
Public Const TILE_TYPE_BLOCK_GUILDE = 27
Public Const TILE_TYPE_BLOCK_TOIT = 28
Public Const TILE_TYPE_BLOCK_DIR = 29
Public Const TILE_TYPE_CRAFT As Byte = 30
Public Const TILE_TYPE_METIER As Byte = 31

' - Types de quêtes
Public Const QUETE_TYPE_AUCUN = 0
Public Const QUETE_TYPE_RECUP = 1
Public Const QUETE_TYPE_APORT = 2
Public Const QUETE_TYPE_PARLER = 3
Public Const QUETE_TYPE_TUER = 4
Public Const QUETE_TYPE_FINIR = 5
Public Const QUETE_TYPE_GAGNE_XP = 6
Public Const QUETE_TYPE_SCRIPT = 7
Public Const QUETE_TYPE_MINIQUETE = 8

' - Type des objets
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

' - Type des métiers
Public Const METIER_CHASSEUR As Byte = 0
Public Const METIER_CRAFT As Byte = 1

' - Type des PNJs
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_QUETEUR = 5
Public Const NPC_BEHAVIOR_SCRIPT = 6

' - Types des sorts
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
'Public Const SPELL_TYPE_GIVEITEM = 7
Public Const SPELL_TYPE_SCRIPT = 6
Public Const SPELL_TYPE_AMELIO = 7
Public Const SPELL_TYPE_DECONC = 8
Public Const SPELL_TYPE_PARALY = 9
Public Const SPELL_TYPE_DEFENC = 10
Public Const SPELL_TYPE_TELE = 11

' - Type des "ciblages"
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_CASE = 2

' - Directions
Public Const DIR_UP = 3
Public Const DIR_DOWN = 0
Public Const DIR_LEFT = 1
Public Const DIR_RIGHT = 2

' - Vitesse des joueurs
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' - Météo
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' - Gestion du temps
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' - Accès des joueurs (0 = joueur normal)
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' - Sexes
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' - Images
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public PIC_PL As Byte
Public PIC_NPC1 As Byte
Public PIC_NPC2 As Byte

' - Divers
Public Const NO = 0
Public Const YES = 1

