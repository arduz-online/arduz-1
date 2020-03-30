Attribute VB_Name = "Protocol"
'
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'

'
'
'
'
'
'
'
'
'
'
'
'
'

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 245

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue


Private Enum ServerPacketID
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    UpdateNeeded            ' REAU
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMidi                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
End Enum

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENT
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTIN
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    RequestTCPStats         '/TCPESSTATS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
End Enum

''
' Handles incoming data.
'


Public Sub HandleIncomingData(ByVal UserIndex As Integer)
'

'01/09/07
'
'
On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(UserIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.ThrowDices _
      Or packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
    End If
    
    Select Case packetID
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.ThrowDices              'TIRDAD
            Call HandleThrowDices(UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
            Call HanldeCombatModeToggle(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(UserIndex)
        
        Case ClientPacketID.PartyCreate             '/CREARPARTY
            Call HandlePartyCreate(UserIndex)
        
        Case ClientPacketID.PartyJoin               '/PARTY
            Call HandlePartyJoin(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeAdminStat(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASEÑA
            Call HandleChangePassword(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(UserIndex)
        
        Case ClientPacketID.PartyKick               '/ECHARPARTY
            Call HandlePartyKick(UserIndex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(UserIndex)
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(UserIndex)
        
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
        
        
        'GM messages
        Case ClientPacketID.GMMessage               '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case ClientPacketID.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case ClientPacketID.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case ClientPacketID.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case ClientPacketID.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case ClientPacketID.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case ClientPacketID.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case ClientPacketID.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case ClientPacketID.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case ClientPacketID.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case ClientPacketID.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case ClientPacketID.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
        
        Case ClientPacketID.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case ClientPacketID.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
        
        Case ClientPacketID.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case ClientPacketID.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)
        
        Case ClientPacketID.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
            
        Case ClientPacketID.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case ClientPacketID.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case ClientPacketID.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
        
        Case ClientPacketID.Forgive                 '/PERDON
            Call HandleForgive(UserIndex)
            
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
            
        Case ClientPacketID.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
            
        Case ClientPacketID.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(UserIndex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
            
        Case ClientPacketID.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)
            
        Case ClientPacketID.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
            
        Case ClientPacketID.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
            
        Case ClientPacketID.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
            
        Case ClientPacketID.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(UserIndex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(UserIndex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(UserIndex)
        
        Case ClientPacketID.BanIP                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
        
        Case ClientPacketID.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)
        
        Case ClientPacketID.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case ClientPacketID.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(UserIndex)
        
        Case ClientPacketID.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(UserIndex)
        
        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case ClientPacketID.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(UserIndex)
        
        Case ClientPacketID.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(UserIndex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(UserIndex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case ClientPacketID.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)
        
        Case ClientPacketID.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case ClientPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)
        
        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(UserIndex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)
            
        Case ClientPacketID.night                   '/NOCHE
            Call HandleNight(UserIndex)
        
        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case ClientPacketID.RequestTCPStats         '/TCPESSTATS
            Call HandleRequestTCPStats(UserIndex)
        
        Case ClientPacketID.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case ClientPacketID.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case ClientPacketID.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case ClientPacketID.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case ClientPacketID.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)
        
        Case ClientPacketID.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(UserIndex)
        
        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
        
#If SeguridadAlkon Then
        Case Else
            Call HandleIncomingDataEx(UserIndex)
#Else
        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
#End If
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(UserIndex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.Source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
    End If
End Sub

''
'LoginExistingChar" message.
'


Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 22 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    
    UserName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    UserList(UserIndex).flags.NoActualizado = Not VersionesActuales(buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger())
        
        'If BANCheck(UserName) Then
        '    Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Argentum debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.com.ar")
        'Else
        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en http://ao.noicoder.com")
        Else
            Call ConnectUser(UserIndex, UserName, Password)
        End If

    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ThrowDices" message.
'


Private Sub HandleThrowDices(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = 9 + RandomNumber(0, 4) + RandomNumber(0, 5)
        .UserAtributos(eAtributos.Agilidad) = 9 + RandomNumber(0, 4) + RandomNumber(0, 5)
        .UserAtributos(eAtributos.Inteligencia) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
        .UserAtributos(eAtributos.Carisma) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
        .UserAtributos(eAtributos.Constitucion) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
    End With
    
    Call WriteDiceRoll(UserIndex)
End Sub

''
'LoginNewChar" message.
'


Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
'

'05/17/06
'
'
#If SeguridadAlkon Then
    If UserList(UserIndex).incomingData.length < 81 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(UserIndex).incomingData.length < 49 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim skills(NUMSKILLS - 1) As Byte
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    Dim mail As String
    
#If SeguridadAlkon Then
    Dim MD5 As String
#End If
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    UserName = buffer.ReadASCIIString()
    
#If SeguridadAlkon Then
    Password = buffer.ReadASCIIStringFixed(32)
#Else
    Password = buffer.ReadASCIIString()
#End If
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    UserList(UserIndex).flags.NoActualizado = Not VersionesActuales(buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger(), buffer.ReadInteger())
    

    
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Call buffer.ReadBlock(skills, NUMSKILLS)
    mail = buffer.ReadASCIIString()
    homeland = buffer.ReadByte()
    
        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            'Call ConnectNewUser(UserIndex, UserName, Password, race, gender, Class, skills, mail, homeland)
        End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Talk" message.
'


Private Sub HandleTalk(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        If LenB(chat) <> 0 Then
            Dim chars As String
            chars = IIf(.bando = eKip.eCUI, Chr(3), Chr(4))
            chars = IIf(.bando = eKip.eNone, Chr(5), chars)
            If .flags.Muerto = 1 Then
                If Len(RTrim(LTrim(chat))) > 0 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat(chars & .nick & " " & IIf(.bando = eKip.eNone, "(ESPECTADOR)", "(MUERTO)") & ": " & chat))
                End If
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor))
                If CInt(Len(RTrim(LTrim(chat)))) > 0 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat(chars & .nick & ": " & chat))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Yell" message.
'


Private Sub HandleYell(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
        Else
            '[Consejeros & GMs]
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If LenB(chat) <> 0 Then
                'Analize chat...
                'Call Statistics.ParseChat(chat)
                
                If .dios = False Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Whisper" message.
'


Private Sub HandleWhisper(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        Dim targetCharIndex As Integer
        Dim targetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = buffer.ReadInteger()
        chat = buffer.ReadASCIIString()
        
        targetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        targetPriv = UserList(targetUserIndex).flags.Privilegios
        
            If targetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                    If LenB(chat) <> 0 Then
                        Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, vbBlue)
                        Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, vbBlue)
                        Call FlushBuffer(targetUserIndex)
                        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                            Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("a " & UserList(targetUserIndex).name & "> " & chat, .Char.CharIndex, vbYellow))
                        End If
                    End If
            End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Walk" message.
'


Private Sub HandleWalk(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    

                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                'If not under a spell effect, show char
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
        
        If .flags.Muerto = 1 Then
            Call Empollando(UserIndex)
        Else
            .flags.EstaEmpo = 0
            .EmpoCont = 0
        End If
    End With
End Sub

''
'RequestPositionUpdate" message.
'


Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    'Call WritePosUpdate(UserIndex)
    Call WriteParalizeOK(UserIndex)
End Sub

''
'Attack" message.
'


Private Sub HandleAttack(ByVal UserIndex As Integer)
'

'10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

''
'PickUp" message.
'


Private Sub HandlePickUp(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
        '    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call GetObj(UserIndex)
    End With
End Sub

''
'CombatModeToggle" message.
'


Private Sub HanldeCombatModeToggle(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        If .admin = True Or .dios = True Then
            Call WriteSafeModeOn(UserIndex)
        End If
    End With
End Sub

''
'SafeToggle" message.
'


Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteSafeModeOff(UserIndex)
        Else
            'Call WriteSafeModeOn(UserIndex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
'ResuscitationSafeToggle" message.
'


Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
'
'Author: Rapsodius
'Creation Date: 10/10/07
'
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        'Call WriteRangingMap(UserIndex)
    End With
End Sub

''
'RequestGuildLeaderInfo" message.
'


Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'RequestAtributes" message.
'


Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)
End Sub

''
'RequestFame" message.
'


Private Sub HandleRequestFame(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)
End Sub

''
'RequestSkills" message.
'


Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
End Sub

''
'RequestMiniStats" message.
'


Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)
End Sub

''
'CommerceEnd" message.
'


Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
'UserCommerceEnd" message.
'


Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        

        
        'Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
'BankEnd" message.
'


Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
'UserCommerceOk" message.
'


Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted

End Sub

''
'UserCommerceReject" message.
'


Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        



        

    End With
End Sub

''
'Drop" message.
'


Private Sub HandleDrop(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim amount As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Then Exit Sub
        If Slot = FLAGORO Then
            If amount > 10000 Then Exit Sub
            Call WriteUpdateGold(UserIndex)
        Else
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                Call DropObj(UserIndex, Slot, amount, .Pos.map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
'CastSpell" message.
'


Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
    End With
End Sub

''
'LeftClick" message.
'


Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, X, Y)
    End With
End Sub

''
'DoubleClick" message.
'


Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.map, X, Y)
    End With
End Sub

''
'Work" message.
'


Private Sub HandleWork(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
                If .flags.Navegando = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub

''
'UseSpellMacro" message.
'


Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

''
'UseItem" message.
'


Private Sub HandleUseItem(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(UserIndex, Slot)
    End With
End Sub

''
'CraftBlacksmith" message.
'


Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        'Call HerreroConstruirItem(UserIndex, Item)
    End With
End Sub

''
'CraftCarpenter" message.
'


Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        'Call CarpinteroConstruirItem(UserIndex, Item)
    End With
End Sub

''
'WorkLeftClick" message.
'


Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
                        Or Not InMapBounds(.Pos.map, X, Y) Then
            Exit Sub
        End If
        
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                    ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
                        DummyInt = 1
                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        'Call WriteConsoleMsg(UserIndex, "Sos un flgger chitero(?.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        'Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Call UsuarioAtacaUsuario(UserIndex, tU)
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                With .Invent
                    DummyInt = .MunicionEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    'Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                    
                    If .Object(DummyInt).amount > 0 Then
                        'QuitarUserInvItem unequipps the ammo, so we equip it again
                        .MunicionEqpSlot = DummyInt
                        .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                        .Object(DummyInt).Equipped = 1
                    Else
                        .MunicionEqpSlot = 0
                        .MunicionEqpObjIndex = 0
                    End If
                    'Call UpdateUserInv(False, UserIndex, DummyInt)
                End With
                '-----------------------------------
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                End If
                
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                
                    'Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Pesca
            
            Case eSkill.Robar
            
            Case eSkill.Talar
            
            Case eSkill.Mineria
            
            Case eSkill.Domar
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
            Case eSkill.Herreria

        End Select
    End With
End Sub

''
'CreateNewGuild" message.
'


Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String
        
        desc = buffer.ReadASCIIString()
        GuildName = buffer.ReadASCIIString()
        site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)

            'Update tag
        '     Call RefreshCharStatus(UserIndex)
        'Else
        '    Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        'End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SpellInfo" message.
'


Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
'EquipItem" message.
'


Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemslot As Byte
        
        itemslot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate item slot
        If itemslot > MAX_INVENTORY_SLOTS Or itemslot < 1 Then Exit Sub
        
        If .Invent.Object(itemslot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemslot)
    End With
End Sub

''
'ChangeHeading" message.
'


Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'

'06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
'ModifySkills" message.
'


Private Sub HandleModifySkills(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then

                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then

            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
End Sub

''
'Train" message.
'


Private Sub HandleTrain(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim petIndex As Byte
        
        petIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If petIndex > 0 And petIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(petIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
'CommerceBuy" message.
'


Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
    End With
End Sub

''
'BankExtractItem" message.
'


Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
    End With
End Sub

''
'CommerceSell" message.
'


Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
    End With
End Sub

''
'BankDeposit" message.
'


Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
    End With
End Sub

''
'ForumPost" message.
'


Private Sub HandleForumPost(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim file As String
        Dim title As String
        Dim msg As String
        Dim postFile As String
        
        Dim handle As Integer
        Dim i As Long
        Dim Count As Integer
        
        title = buffer.ReadASCIIString()
        msg = buffer.ReadASCIIString()
        
        'If .flags.TargetObj > 0 Then
            'file = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
            'If FileExist(file, vbNormal) Then
                'Count = val(GetVar(file, "INFO", "CantMSG"))
                
                'If there are too many messages, delete the forum
                'If Count > MAX_MENSAJES_FORO Then
                    'For i = 1 To Count
                        'Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
                    'Next i
                    'Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
         '           Count = 0
          '      End If
           ' Else
                'Starting the forum....
            '    Count = 0
            'End If
            
            handle = FreeFile()
            'postFile = Left$(file, Len(file) - 4) & CStr(Count + 1) & ".for"
            
            'Create file
            'Open postFile For Output As handle
            'Print #handle, title
            'Print #handle, msg
            'Close #handle
            
            'Update post count
            'Call WriteVar(file, "INFO", "CantMSG", Count + 1)
        'End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())
    End With
End Sub

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim codex() As String
        
        desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'UserCommerceOffer" message.
'


Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadLong()
        
        
    End With
End Sub

''
'GuildAcceptPeace" message.
'


Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildRejectAlliance" message.
'


Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildRejectPeace" message.
'


Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildAcceptAlliance" message.
'


Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildOfferPeace" message.
'


Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildOfferAlliance" message.
'


Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildAllianceDetails" message.
'


Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildPeaceDetails" message.
'


Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildRequestJoinerInfo" message.
'


Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildAlliancePropList" message.
'


Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'GuildPeacePropList" message.
'


Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

End Sub

''
'GuildDeclareWar" message.
'


Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildNewWebsite" message.
'


Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        buffer.ReadASCIIString
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildAcceptNewMember" message.
'


Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildRejectNewMember" message.
'


Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
'

'01/08/07
'Last Modification by: (liquid)
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildKickMember" message.
'


Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim guildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildUpdateNews" message.
'


Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        buffer.ReadASCIIString

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildMemberInfo" message.
'


Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        buffer.ReadASCIIString

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildOpenElections" message.
'


Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim error As String
        

    End With
End Sub

''
'GuildRequestMembership" message.
'


Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
           Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildRequestDetails" message.
'


Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        buffer.ReadASCIIString

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Online" message.
'


Private Sub HandleOnline(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    Dim i As Long
    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To LastUser
            If LenB(UserList(i).name) <> 0 Then
                    Count = Count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'Quit" message.
'


Private Sub HandleQuit(ByVal UserIndex As Integer)
'

'04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'
    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub

''
'GuildLeave" message.
'


Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    Dim guildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        'guildIndex = m_EcharMiembroDeClan(UserIndex, .name)
       '
        'If guildIndex > 0 Then
        '    Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
        '    Call SendData(SendTarget.ToGuildMembers, guildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        'Else
        '    Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        'End If
    End With
End Sub

''
'RequestAccountState" message.
'


Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
        Call UserList(UserIndex).incomingData.ReadByte
End Sub

Private Sub HandlePetStand(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
'PetFollow" message.
'


Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
'TrainList" message.
'


Private Sub HandleTrainList(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

''
'Rest" message.
'


Private Sub HandleRest(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
'Meditate" message.
'


Private Sub HandleMeditate(ByVal UserIndex As Integer)
'

'04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             'Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            .Char.loops = INFINITE_LOOPS
            .Char.FX = FXIDs.FXMEDITARXXGRANDE
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
'Resucitate" message.
'


Private Sub HandleResucitate(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'Heal" message.
'


Private Sub HandleHeal(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHP = .Stats.MaxHP
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'RequestStats" message.
'


Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'Help" message.
'


Private Sub HandleHelp(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

''
'CommerceStart" message.
'


Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

    End With
End Sub

''
'BankStart" message.
'


Private Sub HandleBankStart(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        .envios_recibido = .envios_recibido + 1
        If .envios_recibido > 20 Then
            Call WriteConsoleMsg(UserIndex, "Hemos detectado un posible speedhack en tu pc, desactivalo o serás echado. Tenés " & (.envios_recibido - 20) & " de 20 advertencias antes de ser baneado.", FONTTYPE_VENENO)
            Call WriteChatOverHead(UserIndex, "¡APAGÁ EL SH! Tenés " & (.envios_recibido - 20) & " de 20 advertencias antes de ser echado.", UserList(UserIndex).Char.CharIndex, vbYellow)
            If .envios_recibido > 40 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
            End If
        End If
        'Dead people can't commerce
    End With
End Sub

''
'Enlist" message.
'


Private Sub HandleEnlist(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        

    End With
End Sub

''
'Information" message.
'


Private Sub HandleInformation(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
'Reward" message.
'


Private Sub HandleReward(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        

    End With
End Sub

''
'RequestMOTD" message.
'


Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendMOTD(UserIndex)
End Sub

''
'UpTime" message.
'


Private Sub HandleUpTime(ByVal UserIndex As Integer)
'

'01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = time & " día, " & UpTimeStr
    Else
        UpTimeStr = time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
'PartyLeave" message.
'


Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'PartyCreate" message.
'


Private Sub HandlePartyCreate(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'PartyJoin" message.
'


Private Sub HandlePartyJoin(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'Inquiry" message.
'


Private Sub HandleInquiry(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'ConsultaPopular.SendInfoEncuesta (UserIndex)
End Sub

''
'GuildMessage" message.
'


Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            'Call Statistics.ParseChat(chat)
            
            If .guildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .guildIndex, PrepareMessageGuildChat(.name & "> " & chat))
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                'Call SendData(SendTarget.ToClanArea, userindex, UserList(userindex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(userindex).Char.CharIndex))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'PartyMessage" message.
'


Private Sub HandlePartyMessage(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
End Sub

''
'CentinelReport" message.
'


Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        .incomingData.ReadInteger
    End With
End Sub

''
'GuildOnline" message.
'


Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
    End With
End Sub

''
'PartyOnline" message.
'


Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    

End Sub

''
'CouncilMessage" message.
'


Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            'Call Statistics.ParseChat(chat)
            
            If .bando = eKip.epk Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .bando = eKip.eCUI Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RoleMasterRequest" message.
'


Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GMRequest" message.
'


Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.name)
        Else
            Call Ayuda.Quitar(.name)
            Call Ayuda.Push(.name)
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
'BugReport" message.
'


Private Sub HandleBugReport(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim N As Integer
        
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = buffer.ReadASCIIString()
        
        N = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChangeDescription" message.
'


Private Sub HandleChangeAdminStat(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Description As String
        
        Description = buffer.ReadASCIIString()
        
        If Description = adminpasswd Then
            If UserList(UserIndex).admin = False Then
                UserList(UserIndex).admin = True
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El usuario " & UserList(UserIndex).name & " se identificó como admin!", FontTypeNames.FONTTYPE_TALK))
            Else
                Call WriteConsoleMsg(UserIndex, "Dejás de ser admin de esta partida!", FONTTYPE_TALK)
                UserList(UserIndex).admin = False
            End If
        Else
            If UserList(UserIndex).admin = True Then Call WriteConsoleMsg(UserIndex, "Dejás de ser admin de esta partida!", FONTTYPE_TALK)
            UserList(UserIndex).admin = False
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildVote" message.
'


Private Sub HandleGuildVote(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote As String
        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Punishments" message.
'


Private Sub HandlePunishments(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim name As String
        Dim Count As Integer
        
        name = buffer.ReadASCIIString()
        
        If LenB(name) <> 0 Then
            If (InStrB(name, "\") <> 0) Then
                name = Replace(name, "\", "")
            End If
            If (InStrB(name, "/") <> 0) Then
                name = Replace(name, "/", "")
            End If
            If (InStrB(name, ":") <> 0) Then
                name = Replace(name, ":", "")
            End If
            If (InStrB(name, "|") <> 0) Then
                name = Replace(name, "|", "")
            End If
            
                    Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChangePassword" message.
'


Private Sub HandleChangePassword(ByVal UserIndex As Integer)
'

'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'
#If SeguridadAlkon Then
    If UserList(UserIndex).incomingData.length < 65 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        
        'Remove packet ID
        Call buffer.ReadByte
        

        oldPass = buffer.ReadASCIIString()
        newPass = buffer.ReadASCIIString()

        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo", FontTypeNames.FONTTYPE_INFO)
        Else
            .passwd = newPass
            Call WriteConsoleMsg(UserIndex, "La clave de acceso ha cambiado.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
'Gamble" message.
'


Private Sub HandleGamble(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Integer
        
        amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else

            

            

            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

''
'InquiryVote" message.
'


Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        'Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

''
'BankExtractGold" message.
'


Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        
        amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If amount > 0 And amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - amount
             .Stats.GLD = .Stats.GLD + amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

''
'LeaveFaction" message.
'


Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
    End With
End Sub

''
'BankDepositGold" message.
'


Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        
        amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If amount > 0 And amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + amount
            .Stats.GLD = .Stats.GLD - amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
'Denounce" message.
'


Private Sub HandleDenounce(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String
        
        Text = buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            'Call Statistics.ParseChat(Text)
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GuildFundate" message.
'


Private Sub HandleGuildFundate(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim error As String
        
        clanType = .incomingData.ReadByte()
        
    End With
End Sub

''
'PartyKick" message.
'


Private Sub HandlePartyKick(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        map = .incomingData.ReadInteger()
        If .dios = False Then
        If .admin = False Then Exit Sub
        End If
        If map <= NumMaps Then
            servermap = map

            frmMain.mapax.ListIndex = map - 1
            Call cambiarmapa
        End If
    End With
End Sub

''
'PartySetLeader" message.
'


Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
End Sub

''
'PartyAcceptMember" message.
'


Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)
'

'04/13/2008 (NicoNZ)
'Ya no se puede saber si esta ON o no un personaje
'mediante este comando
End Sub

''
'GuildMemberList" message.
'


Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = buffer.ReadASCIIString()
        

        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GMMessage" message.
'


Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'

'01/08/07
'Last Modification by: (liquid)
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ShowName" message.
'


Private Sub HandleShowName(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .admin = True Or .dios = True Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
UserList(UserIndex).incomingData.ReadByte
End Sub


Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .incomingData.ReadByte
    End With
End Sub

''
'GoNearby" message.
'


Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'

'01/10/07
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        

            If .admin = True Then
                If tIndex <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.map, X, Y, True)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For
                        Next X
                        
                        If found Then Exit For
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        'End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Comment" message.
'


Private Sub HandleComment(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ServerTime" message.
'


Private Sub HandleServerTime(ByVal UserIndex As Integer)
'

'01/08/07
'Last Modification by: (liquid)
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'Where" message.
'


Private Sub HandleWhere(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If .admin = True Or .dios = True Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)

            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'CreaturesInMap" message.
'


Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'

'30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        .incomingData.ReadInteger
    End With
End Sub

''
'WarpMeToTarget" message.
'


Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .dios = False Then Exit Sub
        
        Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, False)
    End With
End Sub
Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim map As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If .dios = True Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
                Else
                tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, X, Y) Then
                    Call WarpUserChar(tUser, map, X, Y, False)
                    'Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    'Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Silence" message.
'


Private Sub HandleSilence(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SOSShowList" message.
'


Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

''
'SOSRemove" message.
'


Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString()

        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'GoToChar" message.
'


Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        Call .incomingData.CopyBuffer(buffer)
        If (tUser > 0) And (.dios = True) Then
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.Y + 1, True)
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
        End If
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Invisible" message.
'


Private Sub HandleInvisible(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        If .dios = False Then Exit Sub
        Call DoAdminInvisible(UserIndex)
    End With
End Sub

''
'GMPanel" message.
'


Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

    End With
End Sub


'GMPanel" message.
'


Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'

'01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'
    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Dim total As Integer
        total = maxusers + MAXNPCS
        ReDim names(1 To total) As String
        Count = 1
        
        For i = 1 To maxusers
            If (LenB(UserList(i).name) <> 0) And UserList(i).flags.AdminInvisible = 0 Then
                    names(Count) = UserList(i).name & IIf(UserList(i).flags.Muerto = 1, " [MUERTO]", "") & "@" & i & "@0@" & UserList(i).Stats.UsuariosMatados & "@" & UserList(i).Stats.muertes & "@" & UserList(i).Stats.puntos & "@" & CInt(UserList(i).admin) & "@" & CInt(UserList(i).bando) & "@" & UserList(i).modName & "@-"
                    Count = Count + 1
            End If
        Next i
        For i = 1 To MAXNPCS
            If Npclist(i).Numero <> 0 Then
                    names(Count) = Npclist(i).name & "@" & i & "@1@" & Npclist(i).bando
                    Count = Count + 1
            End If
        Next i
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
    End With
End Sub

''
'Working" message.
'


Private Sub HandleWorking(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        

    End With
End Sub

''
'Hiding" message.
'


Private Sub HandleHiding(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'Jail" message.
'


Private Sub HandleJail(ByVal UserIndex As Integer)
'

'05/17/06
'
'
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'KillNPC" message.
'


Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'

'04/22/08 (NicoNZ)
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'WarnUser" message.
'


Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/26/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'EditChar" message.
'


Private Sub HandleEditChar(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/28/06
'
'
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim commandString As String
        Dim N As Byte
        
        UserName = buffer.ReadASCIIString()
        
        'If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        'Else
        '    tUser = NameIndex(UserName)
        'End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        UserList(tUser).bando = CInt(Arg2) + 1
        For LoopC = 1 To NUMCLASES
            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
        Next LoopC
        If LoopC > NUMCLASES Then
            'Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
            WriteElejirPJ (tUser)
            Exit Sub
        Else
            UserList(tUser).clase = LoopC
        End If
        valido = True
        UserList(tUser).showName = True
        Call LoadUserStats(tUser)
        Call DarCuerpoYCabeza(tUser)
        RefreshCharStatus tUser
        UpdateUserInv True, tUser, 0
        Call UpdateUserHechizos(True, tUser, 0)
        UserList(tUser).OrigChar = UserList(tUser).Char
        UserList(tUser).ultimomatado = 0
        Call UserDie(tUser)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharInfo" message.
'


Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
'
'Author: Fredy Horacio Treboux (liquid)
'01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim targetName As String
        Dim targetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharStats" message.
'


Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharGold" message.
'


Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharInventory" message.
'


Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        Call .incomingData.CopyBuffer(buffer)
        If .admin = False Then Exit Sub
        tUser = NameIndex(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharBank" message.
'


Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharSkills" message.
'


Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Namex As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        Namex = buffer.ReadASCIIString()
        Call .incomingData.CopyBuffer(buffer)
        If .admin = True Or .dios = True Then
            Select Case UCase(Namex)
                Case "INVI"
                    valeinvi = False
                    frmMain.invii.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Invisibilidad esta DESACTIVADA"))
                Case "ESTU"
                    valeestu = False
                    frmMain.estuu.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Estupidez esta DESACTIVADO"))
                Case "BOTS"
                    frmMain.Check2.value = vbUnchecked
                    frmMain.Frame2.Visible = frmMain.Check2.value
                    botsact = frmMain.Check2.value
                    If botsact = False Then
                    pretorianosVivos = 0
                Dim i As Integer
                        For i = 1 To 100
                        If Npclist(i).flags.NPCActive = True Then Call QuitarNPC(i)
                        Next i
                        End If
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Se desactivaron los BOTS"))
                Case "RESU"
                    valeresu = False
                    frmMain.resuu.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Resucitar esta DESACTIVADO"))
                Case "DEATHMATCH"
                    deathm = False
                    frmMain.deathms.value = vbUnchecked
                    frmMain.ffire.Enabled = True
                    atacaequipo = False
                    frmMain.ffire.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("SE DESACTIVÓ EL FUEGO ALIADO!!"))
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("SE DESACTIVÓ LA MODALIDAD DEATHMATCH!"))
                Case "FATUOS"
                    fatuos = False
                    frmMain.fatu.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Las invocaciones están DESACTIVADAS"))
                Case "FUEGOALIADO"
                If deathm = False Then
                    atacaequipo = False
                    frmMain.ffire.value = vbUnchecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("SE DESACTIVÓ EL FUEGO ALIADO!!"))
                End If
            End Select
        End If
        
        
        'If we got here then packet is complete, copy data back to original queue
        
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ReviveChar" message.
'


Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Namex As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        Namex = buffer.ReadASCIIString()
        Call .incomingData.CopyBuffer(buffer)
        
        If .admin = True Or .dios = True Then
            Select Case UCase(Namex)
                Case "INVI"
                    valeinvi = True
                    frmMain.invii.value = vbChecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Invisibilidad esta ACTIVADA"))
                Case "ESTU"
                    valeestu = True
                    frmMain.estuu.value = vbChecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Estupidez esta ACTIVADA"))
                Case "BOTS"
                    frmMain.Check2.value = vbChecked
                    frmMain.Frame2.Visible = frmMain.Check2.value
                    botsact = frmMain.Check2.value
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Se activaron los BOTS"))
                Case "RESU"
                    valeresu = True
                    frmMain.resuu.value = vbChecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Resucitar esta ACTIVADO"))
                Case "DEATHMATCH"
                    deathm = True
                    frmMain.deathms.value = vbChecked
                    frmMain.ffire.value = vbChecked
                    frmMain.ffire.Enabled = False
                    atacaequipo = True
                    Dim i As Integer
                    For i = 1 To maxusers
                        With UserList(i)
                            If .ConnID <> -1 Then
                                If .ConnIDValida And .flags.UserLogged Then
                                                Call UserDieInterno(i)
                                                Call ResetFrags(i)
                                End If
                            End If
                        End With
                    Next i
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("SE ACTIVÓ LA MODALIDAD DEATHMATCH!"))
                    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(65, NO_3D_SOUND, NO_3D_SOUND))
                Case "FATUOS"
                    fatuos = True
                    frmMain.fatu.value = vbChecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Las invocaciones están ACTIVADAS"))
                Case "FUEGOALIADO"
                    atacaequipo = True
                    frmMain.ffire.value = vbChecked
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("¡¡CUIDADO, SE ACTIVÓ EL FUEGO ALIADO!!"))
            End Select
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'OnlineGM" message.
'


Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'
'Author: Fredy Horacio Treboux (liquid)
'12/28/06
'
'
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .dios = False Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).admin = True Then _
                    list = list & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
'OnlineMap" message.
'


Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'Forgive" message.
'


Private Sub HandleForgive(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Kick" message.
'


Private Sub HandleKick(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If .admin = True Or .dios = True Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                If UserList(tUser).dios = True Then
                    Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call CloseSocket(tUser)
                End If
            ElseIf tUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No te podés hechar vos mismo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario no encontrado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'Execute" message.
'


Private Sub HandleExecute(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'BanChar" message.
'


Private Sub HandleBanChar(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        
            Call BanCharacter(UserIndex, UserName, reason)

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'UnbanChar" message.
'


Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim i As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If .admin = True Or .dios = True Then
            For i = 1 To BanIps.Count
                BanIps.Remove 1
            Next i
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'NPCFollow" message.
'


Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        Call .incomingData.CopyBuffer(buffer)
            If tUser > 0 And .dios = True Then
                    Call WriteConsoleMsg(tUser, .name & " te há trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y + 1, True)
            End If
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SpawnListRequest" message.
'


Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte


    End With
End Sub

''
'SpawnCreature" message.
'


Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        

    End With
End Sub

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
    End With
End Sub

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte

    End With
End Sub

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If .dios = True Then
            If LenB(message) <> 0 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'NickToIP" message.
'


Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If .dios = True Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then

                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                    lista = lista & UserList(LoopC).name & ", "
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'IPToNick" message.
'


Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'GuildOnlineMembers" message.
'


Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'TeleportCreate" message.
'


Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If .dios = False Then Exit Sub
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.map, .Pos.X, .Pos.Y - 1)
            .TileExit.map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
    End With
End Sub

''
'TeleportDestroy" message.
'


Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    With UserList(UserIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .dios = False Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.amount, mapa, X, Y)
                
                If MapData(.TileExit.map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
'RainToggle" message.
'


Private Sub HandleRainToggle(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .dios = False Then Exit Sub
        Lloviendo = Not Lloviendo
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
'SetCharDescription" message.
'


Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = buffer.ReadASCIIString()
        

        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ForceMIDIToMap" message.
'


Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
    End With
End Sub

''
'ForceWAVEToMap" message.
'


Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
    End With
End Sub

''
'RoyalArmyMessage" message.
'


Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .bando = eKip.epk Then
            Call SendData(SendTarget.ToCiudadanos, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChaosLegionMessage" message.
'


Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .bando = eKip.eCUI Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'CitizenMessage" message.
'


Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .bando = eKip.epk Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'CriminalMessage" message.
'


Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .bando = eKip.eCUI Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'TalkAsNPC" message.
'


Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/29/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'DestroyAllItemsInArea" message.
'


Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'AcceptRoyalCouncilMember" message.
'


Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChaosCouncilMember" message.
'


Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ItemsInTheFloor" message.
'


Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'MakeDumb" message.
'


Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If .dios = True Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'MakeDumbNoMore" message.
'


Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If .dios = True Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'DumpIPTables" message.
'


Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .dios = False Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
'CouncilKick" message.
'


Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SetTrigger" message.
'


Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
    End With
End Sub

''
'AskTrigger" message.
'


Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'04/13/07
'
'
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'BannedIPList" message.
'


Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        If .dios = False Then
        If .admin = False Then Exit Sub
        End If
        Dim lista As String
        Dim LoopC As Long
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'BannedIPReload" message.
'


Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

    End With
End Sub

''
'GuildBan" message.
'


Private Sub HandleGuildBan(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'BanIP" message.
'


Private Sub HandleBanIP(ByVal UserIndex As Integer)
'

'05/12/08
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedip As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        buffer.ReadBoolean
        tUser = NameIndex(buffer.ReadASCIIString())
            If tUser <= 0 Then
                'Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedip = UserList(tUser).ip
            End If
        reason = buffer.ReadASCIIString()
        If LenB(bannedip) > 0 Then
            If .admin = True Or .dios = True Then
                Call CloseSocket(tUser)
                Call BanIpAgrega(bannedip)
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'UnbanIP" message.
'


Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'

'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedip As String
        
        bannedip = .incomingData.ReadByte() & "."
        bannedip = bannedip & .incomingData.ReadByte() & "."
        bannedip = bannedip & .incomingData.ReadByte() & "."
        bannedip = bannedip & .incomingData.ReadByte()

    End With
End Sub

''
'CreateItem" message.
'


Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        tObj = .incomingData.ReadInteger()
        
        If .dios = False Then Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Objeto As Obj

        Objeto.amount = 1
        Objeto.ObjIndex = tObj
        Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
'DestroyItems" message.
'


Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.name, "/DEST")
        
        If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
'ChaosLegionKick" message.
'


Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RoyalArmyKick" message.
'


Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ForceMIDIAll" message.
'


Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If .admin = False Then Exit Sub
        
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
'ForceWAVEAll" message.
'


Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .dios = False Then
        If .admin = False Then Exit Sub
        End If
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
'RemovePunishment" message.
'


Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'TileBlockedToggle" message.
'


Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub

        Call LogGM(.name, "/BLOQ")
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
'KillNPCNoRespawn" message.
'


Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)
    End With
End Sub

''
'KillAllNearbyNPCs" message.
'


Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.name, "/MASSKILL")
    End With
End Sub

''
'LastIP" message.
'


Private Sub HandleLastIP(ByVal UserIndex As Integer)
'
'Author: Nicolas Matias Gonzalez (NIGO)
'12/30/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        UserName = buffer.ReadASCIIString()
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChatColor" message.
'


Public Sub HandleChatColor(ByVal UserIndex As Integer)
'
'
'12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If .admin Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
'Ignored" message.
'


Public Sub HandleIgnored(ByVal UserIndex As Integer)
'
'
'12/23/06
'Ignore the user
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'CheckSlot" message.
'


Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'Check one Users Slot in Particular from Inventory
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        Slot = buffer.ReadByte() 'Que Slot?
        tIndex = NameIndex(UserName)  'Que user index?
        
        Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & UserName)
           
        If tIndex > 0 Then
            If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ResetAutoUpdate" message.
'


Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)
'
'
'12/23/06
'Reset the AutoUpdate
'
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        If UCase$(.name) <> "MARAXUS" Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'Restart" message.
'


Public Sub HandleRestart(ByVal UserIndex As Integer)
'
'
'12/23/06
'Restart the game
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        'If not(.admin=true or .dios=true) Then Exit Sub
        'If UCase$(.name) <> "MARAXUS" Then Exit Sub
        
        'time and Time BUG!
        'Call LogGM(.name, .name & " reinicio el mundo")
        
        'Call ReiniciarServidor(True)
    End With
End Sub

''
'ReloadObjects" message.
'


Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'
'
'12/23/06
'Reload the objects
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'ReloadSpells" message.
'


Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'
'
'12/23/06
'Reload the spells
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub


Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'
'
'12/23/06
'Reload the Server`s NPC
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'RequestTCPStats" message
'userIndex The index of the user sending the message

Public Sub HandleRequestTCPStats(ByVal UserIndex As Integer)
'
'
'12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Send the TCP`s stadistics
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .admin = True Then Exit Sub
                
        Dim list As String
        Dim Count As Long
        Dim i As Long
        

    
        Call WriteConsoleMsg(UserIndex, "Los datos están en BYTES.", FontTypeNames.FONTTYPE_INFO)
        
        'Send the stats
        With TCPESStats
            Call WriteConsoleMsg(UserIndex, "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
        End With
        
        'Search for users that are working
        For i = 1 To LastUser
            With UserList(i)
                If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
                    If .outgoingData.length > 0 Then
                        list = list & .name & " (" & CStr(.outgoingData.length) & "), "
                        Count = Count + 1
                    End If
                End If
            End With
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Posibles pjs trabados: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'KickAllChars" message
'
'userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'
'
'12/23/06
'Kick all the chars that are online
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'Night" message
'
'userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
'
'
'12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        'If not(.admin=true or .dios=true) Then Exit Sub
        'If UCase$(.name) <> "MARAXUS" Then Exit Sub
        If .dios = False Then
        If .admin = False Then Exit Sub
        End If
        Call WEBCLASS.enviarpjs
        Call restartround
        
    End With
End Sub

''
'ShowServerForm" message
'
'userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'
'
'12/23/06
'Show the server form
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

    End With
End Sub

''
'CleanSOS" message
'
'userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'
'
'12/23/06
'Clean the SOS
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'SaveChars" message
'
'userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'
'
'12/23/06
'Save the characters
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'ChangeMapInfoBackup" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'
'
'12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        .incomingData.ReadBoolean
    End With
End Sub

''
'ChangeMapInfoPK" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'
'
'12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
    End With
End Sub

''
'ChangeMapInfoRestricted" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChangeMapInfoNoMagic" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
    End With
End Sub

''
'ChangeMapInfoNoInvi" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'InviSinEfecto -> Options: "1", "0"
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
    End With
End Sub
            
''
'ChangeMapInfoNoResu" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
    End With
End Sub

''
'ChangeMapInfoLand" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChangeMapInfoZone" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'
'Author: Pablo (ToxicWaste)
'26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SaveMap" message
'
'userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'
'
'12/24/06
'Saves the map
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        

    End With
End Sub

''
'ShowGuildMessages" message
'
'userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
'
'
'12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'DoBackUp" message
'
'userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'
'
'12/24/06
'Show guilds messages
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'ToggleCentinelActivated" message
'
'userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'
'
'12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

    End With
End Sub

''
'AlterName" message
'
'userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'

'12/26/06
'Change user name
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim guildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'AlterName" message
'
'userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
'

'12/26/06
'Change user password
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'AlterPassword" message
'
'userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
'

'12/26/06
'Change user password
'
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'HandleCreateNPC" message
'
'userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'

'12/24/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub


''
'CreateNPCWithRespawn" message
'
'userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'

'12/24/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If Not (.admin = True Or .dios = True) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo con respawn " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

''
'ImperialArmour" message
'
'userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
'

'12/24/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()

    End With
End Sub

''
'ChaosArmour" message
'
'userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
'

'12/24/06
'
'
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()

    End With
End Sub

''
'NavigateToggle" message
'
'userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'

'01/12/07
'
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .dios = False Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
    End With
End Sub

''
'ServerOpenToUsersToggle" message
'
'userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'

'12/24/06
'
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'TurnOffServer" message
'
'userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
'

'12/24/06
'Turns off the server
'
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte

    End With
End Sub

''
'TurnCriminal" message
'
'userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
'

'12/26/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        

                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ResetFactions" message
'
'userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'

'12/26/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RemoveCharFromGuild" message
'
'userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
'

'12/26/06
'
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim guildIndex As Integer
        
        UserName = buffer.ReadASCIIString()

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'RequestCharMail" message
'
'userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
'

'12/26/06
'Request user mail
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SystemMessage" message
'
'userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'
'
'12/29/06
'Send a message to all the users
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'SetMOTD" message
'
'userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
'
'
'03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
'ChangeMOTD" message
'
'userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'

'12/29/06
'Change the MOTD
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    End With
End Sub

''
'Ping" message
'
'userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'
'
'12/24/06
'Show guilds messages
'
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)
    End With
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
'CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'

'05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteElejirPJ(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "NPCSwing" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCSwing)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "NPCKillUser" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UserSwing" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserSwing)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateNeeded" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateNeeded(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateNeeded" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UpdateNeeded)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "SafeModeOn" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "SafeModeOff" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOn(ByVal UserIndex As Integer)
'
'Author: Rapsodius
'10/10/07
'Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOn)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ResuscitationSafeOff" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOff(ByVal UserIndex As Integer)
'
'Author: Rapsodius
'10/10/07
'Writes the "ResuscitationSafeOff" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    'Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOff)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "NobilityLost" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NobilityLost)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
'map The new map to load.
'version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal map As Integer, ByVal version As Integer)
'

'05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(map)
        Call .WriteInteger(version)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'

'05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
'target Part of the body where the user was hitted.
'damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)
'

'05/17/06
'Writes the "NPCHitUser" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
'UserIndex User to which the message is intended.
'damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)
'

'05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal attackerIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, color))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, name, criminal, privileges))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

On Error GoTo Errhandler
    Dim tmp As String
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(40)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.def)
        Call .WriteSingle(100)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(1)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.def)
        Call .WriteLong(obData.Valor)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAttributes(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

End Sub


Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

End Sub



Public Sub WriteRangingMap(ByVal UserIndex As Integer)
On Error GoTo Errhandler
Exit Sub
    Dim i As Long
    Dim validIndexes(21) As Integer
    Dim Count As Integer
    Dim BBMANDAaa As String
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        For i = 1 To maxusers
            'If UserList(i).ConnID <> -1 Then
                If UserList(i).ConnIDValida = True And UserList(i).flags.UserLogged = True Then
                    
                        Count = Count + 1
                        validIndexes(Count) = i

                End If
            'End If
        Next i
        Dim k As Integer
            BBMANDAaa = BBMANDAaa & "SCORES@"
            For k = 1 To Count
                If UserList(validIndexes(k)).flags.AdminInvisible = 0 Then
                    BBMANDAaa = BBMANDAaa & "ç" & (CStr(UserList(validIndexes(k)).nick))
                    BBMANDAaa = BBMANDAaa & "ç" & (CInt(UserList(validIndexes(k)).bando))
                    BBMANDAaa = BBMANDAaa & "ç" & (CInt(validIndexes(k)))
                    BBMANDAaa = BBMANDAaa & "ç@"
                End If
            Next k
    End With
Exit Sub

Errhandler:
Debug.Print "KB"
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteRestOK(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteBlind(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDumb(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)

On Error GoTo Errhandler
    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(Obj.amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.def)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteFame(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        
        Call .WriteLong(UserList(UserIndex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.NobleRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.Promedio)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        Dim inter As New clsSecu
        Call .WriteLong(inter.INT_USEITEMU)
        Call .WriteLong(inter.INT_USEITEMDCK)
        Call .WriteLong(inter.INT_CAST_ATTACK)
        Call .WriteLong(inter.INT_CAST_SPELL)
        Call .WriteLong(inter.INT_ARROWS)
        Call .WriteLong(inter.INT_ATTACK)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal message As String)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

On Error GoTo Errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            tmp = tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            tmp = tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        ' Prepare guild member's list
        tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            tmp = tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            tmp = tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal url As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(url)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call UserList(UserIndex).outgoingData.WriteBoolean(CBool(UserList(UserIndex).flags.Paralizado))
    Call WritePosUpdate(UserIndex)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteTradeOK(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBankOK(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Long)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(ObjData(ObjIndex).name)
        Call .WriteLong(amount)
        Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
        Call .WriteInteger(ObjData(ObjIndex).MinHIT)
        Call .WriteInteger(ObjData(ObjIndex).def)
        Call .WriteLong(123)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            tmp = tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            tmp = tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(tmp) <> 0 Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)

On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            tmp = tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WritePong(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub FlushBuffer(ByVal UserIndex As Integer)

    Dim sndData As String
    
    With UserList(UserIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
    End With
End Sub


Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageGuildChat(ByVal chat As String) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessagePauseToggle() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageRainToggle() As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(name)
        Call .WriteByte(criminal)
        Call .WriteByte(privileges)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)

        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, isCriminal As Boolean, Tag As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteBoolean(isCriminal)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function
