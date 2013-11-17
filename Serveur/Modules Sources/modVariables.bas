Attribute VB_Name = "modVariables"
'                       ######################################
'                       ############  FRoG Creator 1.0   ###########
'                       ##  Module de stockage des variables globales  ##
'                       ##### Dernière modification : JJ/MM/AAAA  #####
'                       ######################################

' -- Réseau --
Public GAME_PORT As Long

'  -- Configuration du jeu --
Public GAME_NAME As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
Public MAX_EMOTICONS As Long
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public Scripting As Byte
Public NOOB_LEVEL As Long
Public PK_LEVEL As Long
Public RATE_EXP As Long
Public RATE_QUETE As Long
Public RATE_MAX As Long
Public MAX_PETS As Long
Public MAX_METIER As Long
Public MAX_RECETTE As Long

' -- Couleurs des messages (client) --
Public SayColor As Long
Public CouleurDesGuilde As Long
Public GlobalColor As Long
Public BroadcastColor As Long
Public TellColor As Long
Public EmoteColor As Long
Public AdminColor As Long
Public HelpColor As Long
Public WhoColor As Long
Public JoinLeftColor As Long
Public NpcColor As Long
Public AlertColor As Long
Public NewMapColor As Long

' -- Météo --
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long
Public RainIntensity As Long
Public InDestroy As Boolean

' -- Configuration du serveur --
Public KeyTimer As Long
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long
Public ServerLog As Boolean
Public CarteFTP As Boolean
Public SpawnSeconds As Long
Global MyScript As clsSadScript
Public clsScriptCommands As clsCommands
Public DetectScriptErr As Boolean
