VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenduz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BUFFER_LEN = 256

Public InetState As Boolean

Private Type Cole
enviado As Boolean
Datos As String
url As String
End Type

Private Cola() As Cole
Private Ultimo As Integer
Private total As Integer
Private ENCola As Integer



Private Type enviadox
    name As String
    users As String
    mapa As String
End Type

Private enviado As enviadox

Private ronda_envios As Integer


Sub TryRequest()
Dim delimiter As String
delimiter = "~"
Dim tempStr As String
tempStr = menduz & delimiter & "/*/"

If frmMain.Inet1.StillExecuting = False And ENCola > 0 Then
    If Len(Cola(1).Datos) > 0 Then
        If Cola(1).enviado = False Then
            frmMain.Inet1.Execute Cola(1).Datos, "GET"
            Debug.Print "ENVIADO>" & Cola(1).Datos
        Else
            frmMain.Inet1.Execute Cola(1).url & securemenduz(tempStr & criptox(Cola(1).Datos, menduz)), "GET"
            Debug.Print "ENVIADO>" & Cola(1).url & securemenduz(tempStr & criptox(Cola(1).Datos, menduz))
            Debug.Print "ENVIADO-DEC>" & Cola(1).url & tempStr & Cola(1).Datos
        End If
    End If
    ENCola = ENCola - 1
    moverlugar
    ReDim Preserve Cola(ENCola) As Cole
End If

End Sub

Private Sub moverlugar()
Dim i As Integer
For i = 0 To ENCola
    Cola(i) = Cola(i + 1)
Next i
End Sub

Sub ADDRequest(request As String)
    ENCola = ENCola + 1
    ReDim Preserve Cola(ENCola) As Cole
    Cola(ENCola).Datos = request
    Cola(ENCola).enviado = False
End Sub

Sub ADDENVIARANK(request As String, url As String)
    ENCola = ENCola + 1
    ReDim Preserve Cola(ENCola) As Cole
    Debug.Print "+COLA-RANK>" & request
    Cola(ENCola).Datos = request
    Cola(ENCola).enviado = True
    Cola(ENCola).url = url
End Sub

Public Sub enviarpjs()
If menduz = "mzbbfdtt" Then Exit Sub
ronda_envios = ronda_envios + 1
Dim i As Integer
For i = 1 To maxusers
    enviar1pj CInt(i)
Next i
End Sub

Sub enviar1pj(i As Integer)
If menduz = "mzbbfdtt" Then Exit Sub
Dim delimiter As String
delimiter = "~"
        With UserList(i)
            'Conexion activa?
            If .ConnID <> -1 Then
                '¿User valido?
                If .ConnIDValida And .flags.UserLogged And LenB(.passwd) > 0 And .passwd <> "NOTIENEPASSWD" Then
                    'Temp1 = Temp1 & i & delimiter & UserList(i).nick & delimiter & sBase64Encode(UserList(i).passwd) & delimiter & UserList(i).Stats.puntosenv & delimiter & UserList(i).Stats.UsuariosMatadosenv & delimiter & UserList(i).Stats.muertesenv & "@"
                    ADDENVIARANK i & delimiter & UserList(i).nick & delimiter & sBase64Encode(UserList(i).passwd) & delimiter & UserList(i).Stats.puntosenv & delimiter & UserList(i).Stats.UsuariosMatadosenv & delimiter & UserList(i).Stats.muertesenv & "@", WEBSERVER & "u.php?a=upd&datos="
                    .Stats.UsuariosMatadosenv = 0
                    .Stats.muertesenv = 0
                    .Stats.puntosenv = 0
                    '.envios_ultimo = GetTickCount()
                    '.envios_qeronda = ronda_envios
                End If
            End If
        End With
End Sub


Public Sub PingToWeb()
If serverrunning = True Then
    TryRequest
    Dim tempStr As String
    Dim delimiter As String
    Dim result As String
    delimiter = "~"
    
    result = Replace(Replace(svname, delimiter, ""), "=", "")
    If enviado.name <> result Then
        tempStr = "SERVERNAME=" & result
        enviado.name = result
    End If
    result = Replace(Replace(frmMain.mapax.list(servermap - 1), delimiter, ""), "=", "")
    If enviado.mapa <> result Then
        tempStr = tempStr & delimiter & "SERVERMAP=" & result
        enviado.mapa = result
    End If
    If enviado.users <> CStr(NumUsers) Then
        tempStr = tempStr & delimiter & "SERVERUSERS=" & NumUsers
        enviado.users = NumUsers
    End If
    
    Debug.Print "JO>" & tempStr
    result = ""
    If LenB(tempStr) > 0 Then
        result = "&datos=" & securemenduz(tempStr)
    End If
    ADDRequest WEBSERVER & "u.php?a=ping&version=" & ULTIMAVERSION & result
    'result = GetUrlSource(WEBSERVER & "u.php?a=ping&datos=" & securemenduz(TempStr))
    'PharseResultWeb result
End If
End Sub

Public Sub CrearServerWeb()
If serverrunning = True Then
    Dim tempStr As String
    Dim delimiter As String
    Dim result As String
    delimiter = "~"
    tempStr = Replace(svname, delimiter, "") & delimiter & Replace(frmMain.mapax.list(servermap - 1), delimiter, "") & delimiter & NumUsers & delimiter & Puerto & delimiter & hIP & delimiter & ULTIMAVERSION
    'result = GetUrlSource(WEBSERVER & "u.php?a=crear&datos=" & securemenduz(TempStr))
    'PharseResultWeb result
    ADDRequest WEBSERVER & "u.php?a=crear&datos=" & securemenduz(tempStr)
End If
End Sub

Public Sub BorrarServerWeb()
If serverrunning = True Then
    Call GetUrlSource(WEBSERVER & "u.php?a=borra")
End If
End Sub

Private Function GetUrlSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    iResult = InternetCloseHandle(hInternet)
    GetUrlSource = sData
End Function


Function PharseResultWeb(result As String)
Dim arrayres() As String
Dim i, k As Integer
arrayres = Split(result, "@")
Debug.Print "RECIVEVIAWEB>" & result
If result Like "*UPD@*" Then
    Dim arraypjs() As String
    Dim arrayitems() As String
    Dim iii As Integer
    Dim ias As Integer
    Dim itemslot As Integer
    Dim pjs() As String
    Dim tiene As Boolean
    menduz = arrayres(1)
    Debug.Print "KEYSEC>" & menduz
    pjs = Split(arrayres(2), "|")
        For i = 1 To UBound(pjs)
            arraypjs = Split(pjs(i), "ç")
            If Len(arraypjs(1)) > 0 Then
            
                k = arraypjs(1)
                
                If UserList(k).ConnID <> -1 Then
                    If UserList(k).ConnIDValida And UserList(k).flags.UserLogged Then
                        
                        'AVISO QUE SE ENVIO
                       ' UserList(k).envios_recibido = GetTickCount()
                       ' UserList(k).enviados_recibe_enronda = ronda_envios
                        '/AVISO
                        
                        If arraypjs(2) Like "*1*" Then
                            Call WriteConsoleMsg(k, "Tus estadisticas se han actualizado en el ranking de la web!", FontTypeNames.FONTTYPE_INFO)
                            

                            
                            'CLAN
                            If Len(arraypjs(3)) > 1 Then
                                UserList(k).modName = CStr(arraypjs(3))
                            Else
                                UserList(k).modName = vbNullString
                            End If
                            '/CLAN
                            'ITEMS
                                arrayitems = Split(arraypjs(4), "-")
                                For ias = LBound(arrayitems) To UBound(arrayitems)
                                    If LenB(arrayitems(ias)) > 0 Then
                                        If IsNumeric(arrayitems(ias)) = True Then
                                            UserList(k).Invent.Object(itemslot + 13).amount = 1
                                            UserList(k).Invent.Object(itemslot + 13).ObjIndex = CInt(arrayitems(ias))
                                            itemslot = itemslot + 1
                                            tiene = True
                                        End If
                                    End If
                                Next ias
                            If tiene = True Then
                                UpdateUserInv True, k, 0
                            End If                           '/ITEMS
                            'DIOS
                            UserList(k).dios = CBool(arraypjs(5))
                            
                            '/DIOS
                            'BANIP
                            If Len(arraypjs(6)) > 1 Then
                                If arraypjs(6) = "S" Then
                                    Dim bannedip As String
                                    bannedip = UserList(k).ip
                                    If LenB(bannedip) > 0 Then
                                            Call CloseSocket(k)
                                            Call BanIpAgrega(bannedip)
                                    End If
                                End If
                            End If
                            '/BANIP

                            RefreshCharStatus k
                            tiene = False
                            itemslot = 0
                        ElseIf arraypjs(2) Like "*2*" Then
                            UserList(k).dios = 0
                            UserList(k).modName = vbNullString
                            RefreshCharStatus k
                            Call WriteConsoleMsg(k, "Tu persojane no está registrado en el ranking, para figurar en el ranking registrate en http://ao.noicoder.com/", FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(k).dios = 0
                            UserList(k).modName = vbNullString
                            RefreshCharStatus k
                            Call WriteConsoleMsg(k, "Tu contraseña no es la misma que la de el ranking, el ranking no ha podido ser actualizado =(, si te equivocaste de contraseña escribí /CONTRASEÑA", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
            End If
        Next i
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor>Ranking enviado...", FontTypeNames.FONTTYPE_SERVER))
ElseIf result Like "*PING@*" Then
    If result Like "*PING@OK*" Then
        menduz = arrayres(2)
        Debug.Print "KEYSEC>" & menduz
    Else
        Call CrearServerWeb
    End If
ElseIf result Like "*MSJ@*" Then
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("WEB>" & arrayres(1), FontTypeNames.FONTTYPE_SERVER))
ElseIf result Like "*MSG@*" Then
MsgBox arrayres(1)
ElseIf result Like "*!@CERRARSERVER@!*" Then
    End
ElseIf result Like "*CREAR@*" Then
    If result Like "*CREAR@OK*" Then
        menduz = arrayres(2)
        Debug.Print "KEYSEC>" & menduz
    Else
        Call BorrarServerWeb
        Call CrearServerWeb
    End If
End If
End Function
