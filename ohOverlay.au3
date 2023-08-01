#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ohOverlay.ico
#AutoIt3Wrapper_Res_File_Add=Res\ohSmall.ico, RT_ICON, 100, 2057
#AutoIt3Wrapper_Outfile_x64=..\_.rc\ohOverlay.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=OhioHealth Overlay Utility
#AutoIt3Wrapper_Res_Fileversion=23.327.1644.45
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_Fileversion_Use_Template=%YY.%MO%DD.%HH%MI.%SE
#AutoIt3Wrapper_Res_ProductName=ohOverlay
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=n
#AutoIt3Wrapper_Res_HiDpi=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]


[ohOverlay]
-Implement Wait For Host
    -Implement dialog for when host is offline.
    -%s is currently offline or is not responding.
    -[Continue] [Wait] [Abort]
-If item is pinned, and Item->Unpin, and WaitHost: Unping will remove cancel the awaited task.
-If item is not pinned, then Wait->Abort will remove it from pin. Do not offer Unpin.

-Add logging routines for debugging.
-Implement watching for Host DNS/Ping.

[_Common]
-Implement wrapper for crash handling.


#ce ----------------------------------------------------------------------------

Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
Opt("GUIOnEventMode",1)

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <ButtonConstants.au3>
#include <WinAPIConv.au3>
#include <WinAPI.au3>
#include <Misc.au3>
#include <GuiMenu.au3>
#include <WinAPISys.au3>
#include <GuiToolTip.au3>
#include <SendMessage.au3>

#include "Includes\Common\InfinityCommon.au3"
#include "Includes\Common\ResourcesEx.au3"
#include "Includes\Common\_Common.au3"
#include "Includes\embedding\psexec64_exe.au3"

;~ If Not (@Compiled) Then
;~     DllCall('Shcore.dll', 'int', 'SetProcessDPIAwareness', 'int', 2)
;~     DllCall("User32.dll","bool","SetProcessDPIAware")
;~     DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
;~ EndIf

$g_bSingleInstance=True
Global $gsPromptContinue=@LF&"Would you like to continue anyway?"
Global $gsAlias="ohOverlay"
Global $g_iHtit = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
Global $g_iFrame = _WinAPI_GetSystemMetrics($SM_CXDLGFRAME)
Global Const $MA_NOACTIVATE = 3
Global Const $MA_NOACTIVATEANDEAT = 4
Global $gCtxMain, $gidBtn, $gidCtxDoIt, $gidCtxExit
Global $gidDummyMenu
Local $iSizeIco=16
Local $iWidth=$iSizeIco*2
$iHeight=$iSizeIco+3+3
$iRight=$iSizeIco+3+3
$iWidth=$iSizeIco*2
$iTop=22+6
Local $iLeft=0
Global $ghCtxMenu
Global $ghGUI
Global $sDesc=""
Global $sDllUser32 = DllOpen("user32.dll")
Global $aAuth[0][3]
Global $gidAuthSep, $gidAuthAdd, $gidClip, $gidClipSend, $gidClipSendMacro, $gidClipSendRaw, $gidMainSepA, $gidMainSepB, $gidCtxDismiss, $gidCtxExit, $gCtxMain, $gidAuth, $gidCtxClipOpenSN
Global $gidClipWatchMenu
Global $gaRes, $gaResLast, $gidClipMenu, $gidClipMenuSep, $sClipAct
Global $gidCtxClipPing, $gidCtxClipResolve, $gidCtxClipRemote, $gidCtxClipMgmt, $gidCtxClipCmd, $gidCtxClipActions, $gidCtxClipFixFw, $gidCtxClipExplore, $gidCtxClipReg, $gidClipMenuPin, $gidClipMenuUnpin
Global $sPSTools="\\<redacted>\Utils\PSTools\"
Global $gbAway, $gbAwayLast
Global $sIcon=@AutoItExe
If Not @Compiled Then $sIcon = @ScriptDir&"\Res\ohSmall.ico"
Global $aDisplays, $aMousePos[4], $aMon[4]; For monitor detection
Local $iLeftLast,$iTopLast
Global $aClipAct[]=[-1]
Global $gsTooltip
Global $ghToolTip
Global $aWatch[0][8]; iMode|sIPv4|sHostname|iTx|iRx|iRtt|iRttMax|iRttMin
Global $aPins[0][2]
Global $aMenu[0]
Global $gsConfig=$g_sDataDir&"\ohOverlay.ini"
;$aPins[0][0]="dt220833"
;
;Local $aColls=_smsCollections('dt220833')
;_ArrayDisplay($aColls)
;lxit

Func _watchDisplay()
    $aPos=WinGetPos($ghGUI)
    If $aPos[0]<>$iLeft Or $aPos[1]<>$iTop Then
        WinMove($ghGUI,"",$iLeft,$iTop)
        If @Compiled Then _Resource_SetToCtrlID($gidBtn,1,$RT_ICON,Default,True)
        _WinAPI_SetDPIAwareness($ghGUI)
        ;GUISetState(@SW_SHOWNOACTIVATE, $ghGUI)
    EndIf
    ;GUISetState(@SW_HIDE, $ghGUI)
    GUISetState(@SW_SHOWNOACTIVATE, $ghGUI)
EndFunc

Func _watchProc()
    ;0: Resolve
EndFunc

Func initGui()
    $iLeft=@DesktopWidth-$iRight
    $ghGUI = GUICreate('WM_NCHITTEST', $iWidth, $iHeight, $iLeft, $iTop, $WS_POPUP, $WS_EX_TOPMOST+$WS_EX_TOOLWINDOW+$WS_EX_NOACTIVATE)
    If @Compiled Then
        $gidBtn=GUICtrlCreateIcon(@AutoItExe,-1,3,3,$iSizeIco,$iSizeIco)
        _Resource_SetToCtrlID($gidBtn,100,$RT_ICON,Default,True)
    Else
        $gidBtn=GUICtrlCreateIcon($sIcon,0,3,3,$iSizeIco,$iSizeIco)
    EndIf
    GUISetBkColor(0xC0C0C0)
    GUIRegisterMsg($WM_NCHITTEST, 'WM_NCHITTEST')
    GUIRegisterMsg($WM_SYSCOMMAND, "On_WM_SYSCOMMAND")
    GUIRegisterMsg($WM_MOUSEACTIVATE, 'WM_EVENTS')
    $gidDummyMenu = GUICtrlCreateDummy()
    $gCtxMain = GUICtrlCreateContextMenu($gidDummyMenu)
    _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gCtxMain),$MNS_NOCHECK+$MNS_AUTODISMISS)
    $hRgn = _WinAPI_CreateRoundRectRgn(0, 0, $iWidth, $iSizeIco+3+3+1, $iSizeIco+1, $iSizeIco+3+3+1)
    $hRgnCut=_WinAPI_CreateRectRgn($iSizeIco+3+3,0,$iWidth,$iSizeIco+3+3+1)
    _WinAPI_CombineRgn ( $hRgn, $hRgn, $hRgnCut, $RGN_DIFF)
    _WinAPI_DeleteObject($hRgnCut)
    _WinAPI_SetWindowRgn($ghGUI, $hRgn)
    GUICtrlSetOnEvent($gidBtn,"_btnEvent")
    GUIRegisterMsg($WM_DISPLAYCHANGE, "onDisplayChange")
    AdlibRegister("_watchDisplay",250)
    AdlibRegister('posTrack',64)
    GUISetState(@SW_SHOWNOACTIVATE, $ghGUI)
    ;ToolTip("Please Wait...",$iInitialPos[0], $iInitialPos[1]+22,"",0,4)
    ;GUIDelete($ghGUI)
EndFunc

Func UninitGui()
    GUICtrlSetOnEvent($gidBtn,"")
    GUIRegisterMsg($WM_DISPLAYCHANGE, "")
    AdlibUnregister("_watchDisplay")
    AdlibUnregister("posTrack")
    GUIDelete($ghGUI)
EndFunc

Func waitForIt()
    Local $bAbort=False
    Do
        _ToolTip("Click Left: Send, Right: Abort")
        Sleep(10)
        If _IsPressed("02", $sDllUser32) Then
            _ToolTip("")
            Sleep(250)
            Return False
        EndIf
    Until _IsPressed("01", $sDllUser32)
    _ToolTip("")
    Sleep(250)
    Return True
EndFunc

Func _ctxExit()
    Exit 0
EndFunc

Func _ctxAuthUPE()
    Local $iIdx=_ctxGetAuthParIdx()
    If Not waitForIt() Then Return
    Send($aAuth[$iIdx][0],1)
    Send("{tab}",0)
    Send(_authDecrypt($aAuth[$iIdx][1]),1)
    Send("{enter}",0)
EndFunc

Func _ctxAuthPE()
    Local $iIdx=_ctxGetAuthParIdx()
    If Not waitForIt() Then Return
    Send(_authDecrypt($aAuth[$iIdx][1]),1)
    Send("{enter}",0)
EndFunc

Func _ctxAuthUP()
    Local $iIdx=_ctxGetAuthParIdx()
    If Not waitForIt() Then Return
    Send($aAuth[$iIdx][0],1)
    Send("{tab}",0)
    Send(_authDecrypt($aAuth[$iIdx][1]),1)
EndFunc

Func _ctxAuthP()
    Local $iIdx=_ctxGetAuthParIdx()
    If Not waitForIt() Then Return
    Send(_authDecrypt($aAuth[$iIdx][1]),1)
EndFunc

Func _authDecrypt($sToken)
    Return _CryptUnprotectData(_Base64Decode($sToken),$g_sCryptDesc)
EndFunc
;~ Func _ctxAuthAdd()
;~     _promptCred()
;~ EndFunc

Func _ctxAuthRemove()
    ;Return MsgBox(48,"ohOverlay","Not yet Implemented")
    Local $iIdx=_ctxGetAuthParIdx()
    Local $bRemove=False
    $iRet=MsgBox(49,"ohOverlay","Warning: Are you sure you want to remove this credential?")
    If $iRet==1 Then
        Local $aAuthNew[0][3], $iAuthX=UBound($aAuth,1), $iAuthY=UBound($aAuth,2)
        For $i=0 To $iAuthX-1
            If $i==$iIdx Then
                $bRemove=True
                ContinueLoop
            EndIf
            $iMax=UBound($aAuthNew,1)
            ReDim $aAuthNew[$iMax+1][$iAuthY]
            For $j=0 To $iAuthY-1
                $aAuthNew[$iMax][$j]=$aAuth[$i][$j]
            Next
        Next
        $aAuth=$aAuthNew
        If $bRemove Then _saveAuth(True)
    EndIf
EndFunc

Func _ctxClipMacro()
    Local $sClip=ClipGet()
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sClip=$aPins[$iIdx][0]
    $sClip=StringReplace($sClip,@CRLF,"{enter}")
    $sClip=StringReplace($sClip,@TAB,"{tab}")
    waitForIt()
    Send($sClip,0)
EndFunc

Func _ctxClipRaw()
    Local $sClip=ClipGet()
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sClip=$aPins[$iIdx][0]
    $sClip=StringReplace($sClip,@CRLF,@CR)
    waitForIt()
    Send($sClip,1)
EndFunc

Func _ctxGetAuthParIdx()
    For $i=0 To UBound($aAuth,1)-1
        If Not IsArray($aAuth[$i][2]) Then ContinueLoop
        $aTemp=$aAuth[$i][2]
        For $j=UBound($aTemp,1)-1 To 0 Step -1
            If @GUI_CtrlId==$aTemp[$j] Then Return $i
        Next
    Next
EndFunc

Func _ctxGetPinParIdx()
    For $i=0 To UBound($aPins,1)-1
        $aTemp=$aPins[$i][1]
        For $j=UBound($aTemp,1)-1 To 0 Step -1
            If @GUI_CtrlId==$aTemp[$j] Then
                ConsoleWrite("_ctxGetPinParIdx:"&$i&@CRLF)
                Return SetError(0,0,$i)
            EndIf
        Next
    Next
    SetError(1,0,0)
EndFunc

Func _ctxReload()
    _ClearMenuEvt()
    _DeleteCxt()
    _checkAuth()
    _InitMenu()
EndFunc

Func _ctxReAuth()
    Local $iIdx=_ctxGetAuthParIdx()
    _AuthAdd(True,$iIdx)
EndFunc

Func _ctxClipPing()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    $aRes=_HostCheck($sItem)
    If @error Then Return
    ConsoleWrite(@Error&','&@extended&@CRLF)
    If $aRes[3] Then
        Return MsgBox(64,"ohOverlay - Ping",$sItem&": replied in "&$aRes[4]&"ms")
    EndIf
EndFunc

Func _ctxClipPin()
    Local $bNew=1
    For $i=0 To UBound($aPins,1)-1
        If $sClipAct==$aPins[$i][0] Then
            $bNew=0
            ExitLoop
        EndIf
    Next
    If Not $bNew Then
        MsgBox(48,"ohOverlay",'"'&$sClipAct&'" is already pinned.')
        Return
    EndIf
    $iMax=UBound($aPins,1)
    ReDim $aPins[$iMax+1][2]
    $aPins[$iMax][0]=$sClipAct
    _savePins()
EndFunc


Func _ctxClipUnpin()
    Local $iIdx=_ctxGetPinParIdx()
    Local $aNew[0][2],$bMod=False
    Local $sName=$aPins[$iIdx][0]
    Local $iRet=MsgBox(33,"ohOverlay",'Are you sure you want to unpin "'&$sName&'"?')
    If $iRet<>1 Then Return
    Local $aTemp=$aPins[$iIdx][1]
    For $i=0 To UBound($aPins,1)-1
        If $i==$iIdx Then
            $bMod=True
            ContinueLoop
        EndIf
        $iMax=UBound($aNew,1)
        ReDim $aNew[$iMax+1][2]
        $aNew[$iMax][0]=$aPins[$i][0]
    Next
    If $bMod Then
        _DeleteCxt()
        $aPins=$aNew
        _savePins()
    EndIf
EndFunc


Func _ctxClipResolve()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    $aRes=_HostCheck($sItem)
    If @error Then Return
    _ArrayColInsert($aRes,0)
    $aFields=StringSplit("IPv4|Hostname|TX|RX|RTT",'|',2)
    For $i=0 To UBound($aRes,1)-1
        $aRes[$i][0]=$aFields[$i]
    Next
    _ArrayDisplay($aRes,$gsAlias&" - Resolve","",64)
EndFunc

Func _TipProc()
    _ToolTip($gsTooltip)
EndFunc

Func _ClearTip()
    AdlibUnRegister("_TipProc")
    $gsTooltip=-1
    _ToolTip("")
EndFunc

Func _SetTip($sTip)
    $gsTooltip=$sTip
    _ToolTip($gsTooltip)
    AdlibRegister("_TipProc",8)
EndFunc

Func _resolveWrap($sHost,$bPrompt=False,$bApi=False)
    Local $iRet,$sErrCommon="There was an error resolving the host."&@LF,$sTitle=$gsAlias,$sMsg
    Local $iRetErr=0,$iRetExt=False
    If $bApi Then $bPrompt=False
    If Not $bApi Then _SetTip("Resolving "&$sHost&"...")
    $aRes=_Resolve($sHost,250)
;~     MsgBox(64,@Error,@extended)
;~     _ArrayDisplay($aRes)
    If Not $bApi Then _ClearTip()
    If Not IsArray($aRes) Then
        If Not $bApi Then $sMsg=$sErrCommon&"Failed to get Hostname or IP from host specified."
        $iRetErr=1
    Else
        If $aRes[0]==False And $aRes[1]==False Then
            $iRetErr=4
            If Not $bApi Then $sMsg=$sErrCommon&"Failed to get Hostname and IP."
        ElseIf $aRes[0]==False And $aRes[1]<>False Then
            $iRetErr=3
            If Not $bApi Then $sMsg=$sErrCommon&"IP lookup Failed."
        ElseIf $aRes[1]==False Then
            $iRetErr=2
            If Not $bApi Then $sMsg=$sErrCommon&"Hostname lookup Failed."
        EndIf
    EndIf
    If Not $bApi And $iRetErr Then
        If $iRetErr And $bPrompt Then $sMsg&=$gsPromptContinue
        $iRet=MsgBox($bPrompt ? 49 : 48,$sTitle,$sMsg)
        If $iRet==2 Then $iRetExt=True
    EndIf
    ;ConsoleWrite($iRetErr&','&$iRetExt&@CRLF)
    Return SetError($iRetErr,$iRetExt,$aRes)
EndFunc

Func _HostCheck($sHost,$bPrompt=False,$bApi=False)
    Local $iRet,$iRetErr,$iRetExt,$sMsg,$sErrCommon,$sTitle="",$aRet[5],$aPing
    $aRes=_resolveWrap($sHost,$bPrompt,$bApi)
	if @Error And (Not $bPrompt And Not @Extended) Then
        Return SetError(1,$bPrompt,False)
    Else
        If UBound($aRes,1)==2 Then
            $aRet[0]=$aRes[0]
            $aRet[1]=$aRes[1]
        ElseIf UBound($aRes,1)==1 Then
            If Not @Compiled Then _ArrayDisplay($aRes)
            $aRet[0]=$aRes[0]
            $aRet[1]=''
        Else
            If Not @Compiled Then _ArrayDisplay($aRes)
            $aRet[0]=''
            $aRet[1]=''
        EndIf
        $aPing=_Ping($aRes[0],1000)
    EndIf
    if @Error Then
        $sMsg="The host does not respond to ping."
        If $bPrompt Then $sMsg&=$gsPromptContinue
        $iRet=MsgBox($bPrompt ? 49 : 48,$sTitle,$sMsg)
        If $iRet==2 And $bPrompt Then Return SetError(2,$bPrompt,False)
    EndIf
    ;_ArrayDisplay($aPing)
    $aRet[2]=$aPing[0]
    $aRet[3]=$aPing[1]
    $aRet[4]=$aPing[2]
    ;_ArrayDisplay($aRet)
    Return SetError(0,$bPrompt,$aRet)
EndFunc

Func _ctxClipRemote()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True), $iPid
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    Local $sCmRcPath="C:\Program Files (x86)\OhioHealth\SCCM Remote Control\CmRcViewer.exe"
    If Not FileExists($sCmRcPath) Then
        MsgBox(16,"ohOverlay","Error: SCCM Remote Control not found.")
        Return
    EndIf
    $iPid=Run($sCmRcPath&' '&$sItem)
    ProcessWait($iPid)
    ;WinWaitActive("[CLASS:RemoteToolsFrame]")
    ;Sleep(250)
    _CmRcViewMenuHook($iPid)
    ConsoleWrite('C'&@CRLF)
EndFunc

Func _ctxClipMgmt()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True)
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    ShellExecute("compmgmt.msc","/computer="&$sItem,@SystemDir)
EndFunc

Func _ctxClipCmd()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True)
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    ;Return MsgBox(48,"ohOverlay - FixFw","Not yet Implemented")
    ;ShellExecute($sPSTools&'psexec64.exe','\\'&$sClipAct&' cmd.exe')
    _PSEXEC64('\\'&$sItem&' cmd.exe')
EndFunc

Func _ctxClipActions()
    Return MsgBox(48,"ohOverlay - Actions","Not yet Implemented")
EndFunc

Func _ctxClipFixFw()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True)
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    ;ShellExecute($sPSTools&'psexec64.exe','\\'&$sClipAct&' cmd.exe /c "echo Enable WMI:&&netsh advfirewall firewall set rule group="windows management instrumentation (wmi)" new enable=yes && echo Allow DCOM Incoming:&&netsh advfirewall firewall add rule dir=in name="DCOM" program=%systemroot%\system32\svchost.exe service=rpcss action=allow protocol=TCP localport=135 && echo Allow WMI Incoming:&&netsh advfirewall firewall add rule dir=in name ="WMI" program=%systemroot%\system32\svchost.exe service=winmgmt action = allow protocol=TCP localport=any && echo Allow UnsecApp Incoming:&&netsh advfirewall firewall add rule dir=in name ="UnsecApp" program=%systemroot%\system32\wbem\usecapp.exe action=allow && echo Allow WMI Outgoing:&&netsh advfirewall firewall add rule dir=out name ="WMI_OUT" program=%systemroot%\system32\svchost.exe service=winmgmt action=allow protocol=TCP localport=any && echo Press any key to exit...&&pause>>nul"')
    _PSEXEC64('\\'&$sItem&' cmd.exe /c "echo Enable WMI:&&netsh advfirewall firewall set rule group="windows management instrumentation (wmi)" new enable=yes && echo Allow DCOM Incoming:&&netsh advfirewall firewall add rule dir=in name="DCOM" program=%systemroot%\system32\svchost.exe service=rpcss action=allow protocol=TCP localport=135 && echo Allow WMI Incoming:&&netsh advfirewall firewall add rule dir=in name ="WMI" program=%systemroot%\system32\svchost.exe service=winmgmt action = allow protocol=TCP localport=any && echo Allow UnsecApp Incoming:&&netsh advfirewall firewall add rule dir=in name ="UnsecApp" program=%systemroot%\system32\wbem\usecapp.exe action=allow && echo Allow WMI Outgoing:&&netsh advfirewall firewall add rule dir=out name ="WMI_OUT" program=%systemroot%\system32\svchost.exe service=winmgmt action=allow protocol=TCP localport=any && echo Press any key to exit...&&pause>>nul"')
EndFunc

Func _ctxClipExplore()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True)
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    ConsoleWrite("\\"&$sItem&"\C$"&@CRLF)
    ShellExecute("\\"&$sItem&"\C$")
EndFunc

Func _ctxClipReg()
    Return MsgBox(48,"ohOverlay - Reg","Not yet Implemented")
EndFunc

Func _ctxClipOpenSN()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    $sNum=StringLower($sItem)
    ;Lookup ticket in SN.
    ; get sys_id
    $sMod=_getMod($sNum)
    ConsoleWrite($sItem&','&$sMod&@CRLF)
    If @error Then Return
    ShellExecute("https://<redacted>.service-now.com/"&$sMod&".do?sys_id="&$sNum)
EndFunc

Func _ctxClipExpPdf()
    $sNum=StringLower($sClipAct)
    ;Lookup ticket in SN.
    ; get sys_id
    $sMod=_getMod($sNum)
    If @error Then Return
    ShellExecute("https://<redacted>.service-now.com/"&$sMod&".do?sys_id="&$sNum&"&PDF")
EndFunc

Func _ctxClipReboot()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    Local $vRet=_HostCheck($sItem,True)
    ConsoleWrite(@Error&','&@Extended&@CRLF)
    If @error And Not @extended Then Return
    ShellExecute("shutdown.exe","/m \\"&$sItem&" -r -t 00 -f")
EndFunc

Func _ctxClipPut()
    Local $sItem=$sClipAct
    Local $iIdx=_ctxGetPinParIdx()
    If Not @error Then $sItem=$aPins[$iIdx][0]
    ClipPut($sItem)
EndFunc

Func _getMod($sNum)
    Switch StringRegExpReplace($sNum,"^(ritm|sctask|req|inc)\d{7,10}$","$1")
        Case "ritm"
            Return SetError(0,0,"sc_req_item")
        Case "sctask"
            Return SetError(0,0,"sc_task")
        Case "req"
            Return SetError(0,0,"sc_request")
        Case "inc"
            Return SetError(0,0,"incident")
        Case Else
            Return SetError(1,0,False)
    EndSwitch
EndFunc

Func _InitMenu()
    ; Gen Auth
    $gidAuth = GUICtrlCreateMenu("Auth", $gCtxMain)
    _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gidAuth),$MNS_NOCHECK+$MNS_AUTODISMISS)
    For $i=0 To UBound($aAuth,1)-1
        Local $aTemp[7]
        If $aAuth[$i][1]==-1 Then
            $aTemp[0]=GUICtrlCreateMenu($aAuth[$i][0]&" (Expired)",$gidAuth)
            _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($aTemp[0]),$MNS_NOCHECK+$MNS_AUTODISMISS)
            $aTemp[1]=GUICtrlCreateMenuItem("ReAuth", $aTemp[0])
            $aTemp[2]=GUICtrlCreateMenuItem("Remove", $aTemp[0])
        Else
            $aTemp[0]=GUICtrlCreateMenu($aAuth[$i][0], $gidAuth)
            _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($aTemp[0]),$MNS_NOCHECK+$MNS_AUTODISMISS)
            $aTemp[1]=GUICtrlCreateMenuItem("User{tab}Pass{enter}", $aTemp[0])
            $aTemp[2]=GUICtrlCreateMenuItem("Pass{enter}", $aTemp[0])
            $aTemp[3]=GUICtrlCreateMenuItem("User{tab}Pass", $aTemp[0])
            $aTemp[4]=GUICtrlCreateMenuItem("Pass", $aTemp[0])
            $aTemp[5]=GUICtrlCreateMenuItem("", $aTemp[0])
            $aTemp[6]=GUICtrlCreateMenuItem("Remove", $aTemp[0])
        EndIf
        $aAuth[$i][2]=$aTemp
    Next
    If UBound($aAuth,1)>0 Then
        $gidAuthSep=GUICtrlCreateMenuItem("", $gidAuth)
    EndIf
    $gidAuthAdd = GUICtrlCreateMenuItem("Add", $gidAuth)
    $gidClip = GUICtrlCreateMenu("Clip", $gCtxMain)
    _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gidClip),$MNS_NOCHECK+$MNS_AUTODISMISS)
    Local $aTemp=StringRegExp(ClipGet(),"([a-zA-Z0-9_\-\$]+)",1)
    If Not @error Then
        $sClip=$aTemp[0]
        ;[iTypes]
        ;Desktop
        ;Laptop
        ;Mobile Device
        ;Apple Laptop
        ;Printer
        ;Server
        ;User
        ;Lenovo Serial Number
        ;OhioHealth Asset Number (Open SN Hardware page)
        ;Phone Number
        ;Email
        ;ServiceNow Incident
        ;ServiceNow Task
        ;ServiceNow Request
        ;ServiceNow Request Item
        ;ServiceNow Interaction
        ;
        ;
        Local $bSN=StringRegExp(StringLower($sClip),"^(inc|sctask|ritm|req)\d{7,10}$")
        ConsoleWrite('"'&StringLower($sClip)&'"'&@CRLF)
        Local $bCI=StringRegExp(StringLower($sClip),"^(?:dt|lt|al|md)\d{6}|pr\d{4,}$")
        Local $bIsIP=StringRegExp(StringLower($sClip),"^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$")
        Local $bIsHostname=StringRegExp(StringLower($sClip),"^(([a-zA-Z0-9_]|[a-zA-Z0-9_][a-zA-Z0-9_\-]*[a-zA-Z0-9_])\.)*([A-Za-z0-9_]|[A-Za-z0-9_][A-Za-z0-9_\-]*[A-Za-z0-9_])$")
        ConsoleWrite($bSN&','&$bCI&','&$bIsIP&','&$bIsHostname&@CRLF)
        If $bSN Or ($bCI Or $bIsIP Or $bIsHostname) Then
            $sClipAct=$sClip
            $gidClipMenu = GUICtrlCreateMenu($sClip, $gidClip)
            _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gidClipMenu),$MNS_NOCHECK+$MNS_AUTODISMISS)
            $gidClipMenuPin=GUICtrlCreateMenuItem("Pin", $gidClipMenu)
            GUICtrlCreateMenuItem("", $gidClipMenu)
        EndIf
        If $bSN Then
            Dim $aClipAct[2]
            $aClipAct[0]=0
            $aClipAct[1]=GUICtrlCreateMenuItem("Open", $gidClipMenu)
            ;$aClipAct[2]=GUICtrlCreateMenuItem("pdfExport", $gidClipMenu)
        Else
            If $bCI Or $bIsIP Or $bIsHostname Then
                ConsoleWrite("!"&$bSN&','&$bCI&','&$bIsIP&','&$bIsHostname&@CRLF)
                Dim $aClipAct[11]
                $aClipAct[0]=1
                $aClipAct[1]=GUICtrlCreateMenuItem("Ping", $gidClipMenu)
                $aClipAct[2]=GUICtrlCreateMenuItem("Check Host", $gidClipMenu)
                If StringRegExp(StringLower($sClip),"(?:dt|lt)\d{6}") Then
                    $aClipAct[3]=GUICtrlCreateMenuItem("Explore", $gidClipMenu)
                    $aClipAct[4]=-1;GUICtrlCreateMenuItem("Registry", $gidClipMenu)
                    $aClipAct[5]=GUICtrlCreateMenuItem("SCCM Remote", $gidClipMenu)
                    $aClipAct[6]=GUICtrlCreateMenuItem("Computer Management", $gidClipMenu)
                    $aClipAct[7]=GUICtrlCreateMenuItem("Command Prompt", $gidClipMenu)
                    $aClipAct[8]=-1;GUICtrlCreateMenuItem("cmActions", $gidClipMenu)
                    $aClipAct[9]=GUICtrlCreateMenuItem("Remote Firewall Fix", $gidClipMenu)
                    $aClipAct[10]=GUICtrlCreateMenuItem("Reboot", $gidClipMenu)
                EndIf
    ;~             $aClipAct[12]=GUICtrlCreateMenuItem("", $gidClipMenu)
    ;~             $gidClipWatchMenu=GUICtrlCreateMenu("Watch", $gidClipMenu)
    ;~             $aClipAct[10]=GUICtrlCreateMenuItem("Ping", $gidClipWatchMenu)
    ;~             $aClipAct[11]=GUICtrlCreateMenuItem("Resolve", $gidClipWatchMenu)
        ;~     ElseIf StringRegExp(StringLower($sClip),"(?:[a-z]\d)?[a-z]{3,4}\d{3,4}") Then; Opid
        ;~         $sClipAct=$sClip
        ;~         $gidClipMenu = GUICtrlCreateMenuItem($sClip, $gidClip)
        ;~         ;_GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gidClip),$MNS_NOCHECK+$MNS_AUTODISMISS)
            EndIf
        EndIf
        If $bSN Or $bCI Or $bIsIP Or $bIsHostname Then $gidClipMenuSep=GUICtrlCreateMenuItem("", $gidClip)
    EndIf
    $gidClipSend = GUICtrlCreateMenu("Send", $gidClip)
    _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($gidClipSend),$MNS_NOCHECK+$MNS_AUTODISMISS)
    $gidClipSendMacro = GUICtrlCreateMenuItem("/w Macros", $gidClipSend)
    $gidClipSendRaw = GUICtrlCreateMenuItem("Raw", $gidClipSend)
    ;$gidClipMacroHelp=GUICtrlCreateMenuItem("Macros?", $gidClipSend)
    $gidMainSepA=GUICtrlCreateMenuItem("", $gCtxMain)
    ; Pins
    _ArrayNaturalSort($aPins)
    If UBound($aPins,1) Then
        For $z=0 To UBound($aPins,1)-1
            If $aPins[$z][0]=="" Then ContinueLoop
            Local $aTemp=_GenCtx($aPins[$z][0],$gCtxMain)
            $aPins[$z][1]=$aTemp
        Next
        $gidMainSepB=GUICtrlCreateMenuItem("", $gCtxMain)
    EndIf
    ; Footer
    $gidCtxDismiss = GUICtrlCreateMenuItem("Dismiss", $gCtxMain)
    $gidCtxExit = GUICtrlCreateMenuItem("Exit", $gCtxMain)
    _SetMenuEvt()
EndFunc

Func _GenCtx($sItem,$idMenu)
    Local $aRet[1]
    Local $sLow=StringLower($sItem)
    Local $bSN=StringRegExp($sLow,"^(inc|sctask|ritm|req)\d{7,10}$")
    Local $bCI=StringRegExp($sLow,"^(?:dt|lt|al|md)\d{6}|pr\d{4,}$")
    Local $bIsIP=StringRegExp($sLow,"^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$")
    Local $bIsHostname=StringRegExp($sLow,"^(([a-zA-Z0-9_]|[a-zA-Z0-9_][a-zA-Z0-9_\-]*[a-zA-Z0-9_])\.)*([A-Za-z0-9_]|[A-Za-z0-9_][A-Za-z0-9_\-]*[A-Za-z0-9_])$")
    $aRet[0]=GUICtrlCreateMenu($sLow,$idMenu)
    _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($aRet[0]),$MNS_NOCHECK+$MNS_AUTODISMISS)
    Local $iLast
    ;GUICtrlCreateMenuItem("To Clip", $gidClipMenu)
    ;GUICtrlCreateMenu("Send", $gidClipMenu)
    If $bSN Then
        ReDim $aRet[8]
        $aRet[1]=0
    Else
        If $bCI Or $bIsIP Or $bIsHostname Then
            ReDim $aRet[17]
            $aRet[1]=1
        EndIf
    EndIf
    If $bSN Or ($bCI Or $bIsIP Or $bIsHostname) Then
        $aRet[2] = GUICtrlCreateMenuItem("To Clip", $aRet[0])
        $aRet[3] = GUICtrlCreateMenu("Send", $aRet[0])
        _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($aRet[3]),$MNS_NOCHECK+$MNS_AUTODISMISS)
        $aRet[4] = GUICtrlCreateMenuItem("/w Macros", $aRet[3])
        $aRet[5] = GUICtrlCreateMenuItem("Raw", $aRet[3])
        $aRet[6] = GUICtrlCreateMenuItem("", $aRet[0])
    EndIf
    If $bSN Then
        $aRet[7]=GUICtrlCreateMenuItem("Open", $aRet[0])
        ;$aClipAct[2]=GUICtrlCreateMenuItem("pdfExport", $gidClipMenu)
    Else
        If $bCI Or $bIsIP Or $bIsHostname Then
            $aRet[7]=GUICtrlCreateMenuItem("Ping", $aRet[0])
            $aRet[8]=GUICtrlCreateMenuItem("Resolve", $aRet[0])
            If StringRegExp($sLow,"(?:dt|lt)\d{6}") Then
                $aRet[9]=GUICtrlCreateMenuItem("Explore", $aRet[0])
                $aRet[10]=-1;GUICtrlCreateMenuItem("Registry", $aRet[0])
                $aRet[11]=GUICtrlCreateMenuItem("SCCM Remote", $aRet[0])
                $aRet[12]=GUICtrlCreateMenuItem("Computer Management", $aRet[0])
                $aRet[13]=GUICtrlCreateMenuItem("Command Prompt", $aRet[0])
                $aRet[14]=-1;GUICtrlCreateMenuItem("cmActions", $gidClipMenu)
                $aRet[15]=GUICtrlCreateMenuItem("Remote Firewall Fix", $aRet[0])
                $aRet[16]=GUICtrlCreateMenuItem("Reboot", $aRet[0])
            EndIf
        EndIf
    EndIf
    $iMax=UBound($aRet,1)
    ReDim $aRet[$iMax+1]
    GUICtrlCreateMenuItem("", $aRet[0])
    $aRet[$iMax]=GUICtrlCreateMenuItem("UnPin", $aRet[0])
    Return $aRet
EndFunc
Func _SetMenuEvt()
    For $i=0 To UBound($aAuth,1)-1
        If Not IsArray($aAuth[$i][2]) Then ContinueLoop
        $aTemp=$aAuth[$i][2]
        If $aAuth[$i][1]==-1 Then
            GUICtrlSetOnEvent($aTemp[1],"_ctxReAuth")
            GUICtrlSetOnEvent($aTemp[2],"_ctxAuthRemove")
        Else
            GUICtrlSetOnEvent($aTemp[1],"_ctxAuthUPE")
            GUICtrlSetOnEvent($aTemp[2],"_ctxAuthPE")
            GUICtrlSetOnEvent($aTemp[3],"_ctxAuthUP")
            GUICtrlSetOnEvent($aTemp[4],"_ctxAuthP")
            GUICtrlSetOnEvent($aTemp[6],"_ctxAuthRemove")
        EndIf
    Next
    GUICtrlSetOnEvent($gidAuthAdd,"_ctxAuthAdd")
    Switch $aClipAct[0]
        Case 0
            GUICtrlSetOnEvent($aClipAct[1],"_ctxClipOpenSN")
            ;GUICtrlSetOnEvent($aClipAct[2],"_ctxClipExpPdf")
        Case 1
            GUICtrlSetOnEvent($aClipAct[1],"_ctxClipPing")
            GUICtrlSetOnEvent($aClipAct[2],"_ctxClipResolve")
            GUICtrlSetOnEvent($aClipAct[3],"_ctxClipExplore")
            ;GUICtrlSetOnEvent($aClipAct[4],"_ctxClipReg")
            GUICtrlSetOnEvent($aClipAct[5],"_ctxClipRemote")
            GUICtrlSetOnEvent($aClipAct[6],"_ctxClipMgmt")
            GUICtrlSetOnEvent($aClipAct[7],"_ctxClipCmd")
            ;GUICtrlSetOnEvent($aClipAct[8],"_ctxClipActions")
            GUICtrlSetOnEvent($aClipAct[9],"_ctxClipFixFw")
            GUICtrlSetOnEvent($aClipAct[10],"_ctxClipReboot")
    EndSwitch
    For $z=0 To UBound($aPins,1)-1
        $aTemp=$aPins[$z][1]
;~         $aRet[2] = GUICtrlCreateMenu("To Clip", $aRet[0])
;~         $aRet[3] = GUICtrlCreateMenu("Send", $aRet[0])
;~         _GUICtrlMenu_SetMenuStyle(GUICtrlGetHandle($aRet[2]),$MNS_NOCHECK+$MNS_AUTODISMISS)
;~         $aRet[4] = GUICtrlCreateMenuItem("/w Macros", $aRet[2])
;~         $aRet[5] = GUICtrlCreateMenuItem("Raw", $aRet[2])
        GUICtrlSetOnEvent($aTemp[2],"_ctxClipPut")
        GUICtrlSetOnEvent($aTemp[4],"_ctxClipMacro")
        GUICtrlSetOnEvent($aTemp[5],"_ctxClipRaw")
        Switch $aTemp[1]
            Case 0
                GUICtrlSetOnEvent($aTemp[7],"_ctxClipOpenSN")
                ;GUICtrlSetOnEvent($aClipAct[2],"_ctxClipExpPdf")
            Case 1
                GUICtrlSetOnEvent($aTemp[7],"_ctxClipPing")
                GUICtrlSetOnEvent($aTemp[8],"_ctxClipResolve")
                GUICtrlSetOnEvent($aTemp[9],"_ctxClipExplore")
                ;GUICtrlSetOnEvent($aClipAct[10],"_ctxClipReg")
                GUICtrlSetOnEvent($aTemp[11],"_ctxClipRemote")
                GUICtrlSetOnEvent($aTemp[12],"_ctxClipMgmt")
                GUICtrlSetOnEvent($aTemp[13],"_ctxClipCmd")
                ;GUICtrlSetOnEvent($aClipAct[14],"_ctxClipActions")
                GUICtrlSetOnEvent($aTemp[15],"_ctxClipFixFw")
                GUICtrlSetOnEvent($aTemp[16],"_ctxClipReboot")
        EndSwitch
        $iMax=UBound($aTemp,1)-1
        GUICtrlSetOnEvent($aTemp[$iMax],"_ctxClipUnpin")

;~         For $y=1 To UBound($aTemp,1)-1
;~         Next
    Next
    GUICtrlSetOnEvent($gidClipMenuPin,"_ctxClipPin")
    GUICtrlSetOnEvent($gidClipSendMacro,"_ctxClipMacro")
    GUICtrlSetOnEvent($gidClipSendRaw,"_ctxClipRaw")
    GUICtrlSetOnEvent($gidCtxDismiss,"_ctxReload")
    GUICtrlSetOnEvent($gidCtxExit,"_ctxExit")
EndFunc

Func _ClearMenuEvt()
    For $i=0 To UBound($aAuth,1)-1
        If Not IsArray($aAuth[$i][2]) Then ContinueLoop
        $aTemp=$aAuth[$i][2]
        $iTempX=UBound($aTemp,1)
        If $aAuth[$i][1]==-1 Then
            GUICtrlSetOnEvent($aTemp[2],"")
            GUICtrlSetOnEvent($aTemp[1],"")
            GUICtrlSetOnEvent($aTemp[0],"")
        Else
            For $j=0 To $iTempX-3
                GUICtrlSetOnEvent($aTemp[$j],"")
            Next
            GUICtrlSetOnEvent($aTemp[$iTempX-1],"")
        EndIf
    Next
    GUICtrlSetOnEvent($gidAuthAdd,"")
    For $i=1 To UBound($aClipAct,1)-1
        GUICtrlSetOnEvent($aClipAct[$i],"")
    Next
    For $z=0 To UBound($aPins,1)-1
        $aTemp=$aPins[$z][1]
        For $y=1 To UBound($aTemp,1)-1
            GUICtrlSetOnEvent($aTemp[$y],"")
        Next
    Next
    GUICtrlSetOnEvent($gidClipMenuPin,"")
    GUICtrlSetOnEvent($gidClipMenuUnpin,"")
    GUICtrlSetOnEvent($gidCtxClipActions,"")
    GUICtrlSetOnEvent($gidClipSendMacro,"")
    GUICtrlSetOnEvent($gidClipSendRaw,"")
    GUICtrlSetOnEvent($gidCtxDismiss,"")
    GUICtrlSetOnEvent($gidCtxExit,"")
EndFunc

Func _DeleteCxt()
    _ClearMenuEvt()
    GUICtrlDelete($gidCtxExit)
    GUICtrlDelete($gidCtxDismiss)
    GUICtrlDelete($gidMainSepA)
    ; Pins
    For $z=0 To UBound($aPins,1)-1
        Local $aTemp=$aPins[$z][1]
        For $y=UBound($aTemp,1)-1 To 0 Step -1
            GUICtrlDelete($aTemp[$y])
        Next
    Next
    GUICtrlDelete($gidMainSepB)
    GUICtrlDelete($gidClipSendRaw)
    GUICtrlDelete($gidClipSendMacro)
    GUICtrlDelete($gidClipMenuSep)
    For $i=1 To UBound($aClipAct,1)-1
        GUICtrlDelete($aClipAct[$i])
    Next
    GUICtrlDelete($gidClipMenu)
    GUICtrlDelete($gidClipSend)
    GUICtrlDelete($gidClip)
    GUICtrlDelete($gidAuthAdd)
    GUICtrlDelete($gidAuthSep)
    ;$gidCtxClipOpenSN
    For $i=0 To UBound($aAuth,1)-1
        $aTemp=$aAuth[$i][2]
        If $aTemp==-1 Then ContinueLoop
        If $aAuth[$i][1]==-1 Then
            GUICtrlDelete($aTemp[0])
        Else
            For $j=UBound($aTemp,1)-1 To 0 Step -1
                GUICtrlDelete($aTemp[$j])
            Next
        EndIf
    Next
    GUICtrlDelete($gidAuth)
EndFunc

;PsaltyDS
Func _FocusCtrlID($hWnd, $sTxt = "")
    Local $hFocus = ControlGetHandle($hWnd, $sTxt, ControlGetFocus($hWnd, $sTxt))
    If IsHWnd($hFocus) Then
        Return _WinAPI_GetDlgCtrlID($hFocus)
    Else
        Return SetError(1, 0, 0)
    EndIf
EndFunc   ;==>_FocusCtrlID

Func _btnEvent()
    _DeleteCxt()
    ;_checkAuth()
    _InitMenu()
    ShowMenu($ghGUI, $gCtxMain, $gidBtn)
EndFunc

Func WM_EVENTS($hWnd, $MsgID, $WParam, $LParam)
    Switch $hWnd
        Case $ghGUI
            Switch $MsgID
                Case $WM_MOUSEACTIVATE
                    ; Check mouse position
                    Local $aMouse_Pos = GUIGetCursorInfo($ghGUI)
                    If $aMouse_Pos[4] <> 0 Then
                        Local $word = _WinAPI_MakeLong($aMouse_Pos[4], $BN_CLICKED)
                        ;ConsoleWrite($aMouse_Pos[4]&@CRLF)
                        _SendMessage($ghGUI, $WM_COMMAND, $word, GUICtrlGetHandle($aMouse_Pos[4]))
                    EndIf
                    Return $MA_NOACTIVATEANDEAT
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc

;Const $SC_MOVE = 0xF010
Func On_WM_SYSCOMMAND($hWnd, $Msg, $wParam, $lParam)
    If $hWnd==$ghGUI Then
        If BitAND($wParam, 0xFFF0) == 0xF010 Then Return 0
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func WM_NCHITTEST($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $wParam, $lParam
    Return $HTCAPTION
EndFunc   ;==>WM_NCHITTEST

Func WM_MOVING($hWnd, $iMsg, $wParam, $lParam)
    If $hWnd = $ghGUI Then
        Local $bLimited = False
        ;Current rect is passed in lParam
        Local $tRect = DllStructCreate($tagRect, $lParam) ; (IS NOT X, Y, W, H) => (IS X, Y, X+W, Y+H)
        Local $iLeft = DllStructGetData($tRect, 1)
        Local $iTop = DllStructGetData($tRect, 2)
        Local $iRight = DllStructGetData($tRect, 3)
        Local $iBottom = DllStructGetData($tRect, 4)
        Local $iWidth, $iHeight

        ;Check left and right of window against imposed limits
        ;ConsoleWrite($iRight&@CRLF)
;~         If $iLeft < $iMinPos[0] Then
;~             $iWidth = $iRight - $iLeft ;Calculate Original Width
;~             $iLeft = $iMinPos[0]
;~             $iRight = $iLeft + $iWidth ;Have to keep our same Width
;~             $bLimited = True
;~         ElseIf $iRight > $iInitialPos[0]+32 Then
;~             $iWidth = $iRight - $iLeft ;Calculate Original Width
;~             $iRight = $iInitialPos[0]+32
;~             $iLeft = $iRight - $iWidth ;Have to keep our same Width
;~             $bLimited = True
;~         EndIf

        ;Check top and bottom of window against imposed limits
;~         If $iTop < $iMinPos[1] Then
;~             $iHeight = $iBottom - $iTop ;Calculate Original Height
;~             $iTop = $iMinPos[1]
;~             $iBottom = $iTop + $iHeight ;Have to keep our same Height
;~             $bLimited = True
;~         ElseIf $iBottom > @DesktopHeight-48 Then
;~             $iHeight = $iBottom - $iTop ;Calculate Original Height
;~             $iBottom =  @DesktopHeight-48
;~             $iTop = $iBottom - $iHeight ;Have to keep our same Height
;~             $bLimited = True
;~         EndIf
;~         If $bLimited Then
;~             DllStructSetData($tRect, 1, $iLeft)
;~             DllStructSetData($tRect, 2, $iTop)
;~             DllStructSetData($tRect, 3, $iRight)
;~             DllStructSetData($tRect, 4, $iBottom)
;~             _WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam);Pass new Rect on to default window procedure
;~         EndIf
        ;$iCurX=MouseGetPos(0)
        ;If $iCurX<>$iMinPos[0] Then
        ;    $iCurY=MouseGetPos(1)
        ;    MouseMove($iMinPos[0],$iCurY,1)
        ;EndIf
        Return 1 ; True
    EndIf
    Return $GUI_RUNDEFMSG ;Default Handler
EndFunc   ;==>WM_MOVING

; Show a menu in a given GUI window which belongs to a given GUI ctrl
Func ShowMenu($hWnd, $nContextID, $nContextControlID, $iMouse=0)
    Local $hMenu = GUICtrlGetHandle($nContextID)
    Local $iCtrlPos = ControlGetPos($hWnd, "", $nContextControlID)
    Local $X = $iCtrlPos[0]
    Local $Y = $iCtrlPos[1] + $iCtrlPos[3]
    ClientToScreen($hWnd, $X, $Y)
    If $iMouse Then
        $X = MouseGetPos(0)
        $Y = MouseGetPos(1)
    EndIf
    DllCall("user32.dll", "int", "TrackPopupMenuEx", "hwnd", $hMenu, "int", 0, "int", $X, "int", $Y, "hwnd", $hWnd, "ptr", 0)
EndFunc

; Convert the client (GUI) coordinates to screen (desktop) coordinates
Func ClientToScreen($hWnd, ByRef $x, ByRef $y)
    Local $stPoint = DllStructCreate("int;int")
    DllStructSetData($stPoint, 1, $x)
    DllStructSetData($stPoint, 2, $y)
    DllCall("user32.dll", "int", "ClientToScreen", "hwnd", $hWnd, "ptr", DllStructGetPtr($stPoint))
    $x = DllStructGetData($stPoint, 1)
    $y = DllStructGetData($stPoint, 2)
    ; release Struct not really needed as it is a local
    $stPoint = 0
EndFunc

Func _ctxAuthAdd()
    _AuthAdd()
EndFunc

Func _AuthAdd($bReAuth=False,$iAuth=-1)
    Opt("GUIOnEventMode",0)
;~ 	Local $bWarnRemember=True
;~ 	Local $bRemember=False
	Local $iWidth=256+128
	Local $hWnd = GUICreate(($bReAuth ? "Refresh" : "Add")&" Credentials", $iWidth, 104+38)
	GUISetFont(10, 400, 0, "Consolas")
	$gidOPID=GUICtrlCreateInput("", 8, 8, $iWidth-16, 23)
    If $bReAuth Then
        GUICtrlSetData($gidOPID,$aAuth[$iAuth][0])
        GUICtrlSetState($gidOPID,$GUI_DISABLE)
        $bEnValidate[1]=True
    Endif
	_GUICtrlEdit_SetCueBanner($gidOPID, "OPID", True)
	$gidPass = GUICtrlCreateInput("", 8, 40, $iWidth-16, 23, $ES_PASSWORD)
	_GUICtrlEdit_SetCueBanner($gidPass, "Password", True)
;~ 	$idRemember = GUICtrlCreateCheckbox("Remember (Encrypted)", 8, 72, 185, 17)
	$idStatus = _GUICtrlStatusBar_Create($hWnd)
	_GUICtrlStatusBar_SetText($idStatus,"Ready")
    Local $iBtnM=8
    Local $iBtnW=72
    Local $iLeft=($iWidth/2)-(($iBtnW+$iBtnM+$iBtnW)/2)
	$idBtnCancel = GUICtrlCreateButton("Cancel",$iLeft, 72, 72, 32)
	$gidBtnValidate = GUICtrlCreateButton($bReAuth ? "Refresh" : "Add",$iLeft+$iBtnM+$iBtnW,72, 72, 32)
	GuiCtrlSetState(-1,$GUI_DISABLE)
    ;_GuiRoundCorners($h_win, $i_x1, $i_y1, $i_x2, $i_y2, $i_x3, $i_y3)

	GUISetState(@SW_SHOW)
	GUIRegisterMsg($WM_COMMAND, "authWM_COMMAND")
	#EndRegion ### END Koda GUI section ###
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $idBtnCancel, $GUI_EVENT_CLOSE
				GUIRegisterMsg($WM_COMMAND,"")
                GUIDelete($hWnd)
                Opt("GUIOnEventMode",1)
                Return
;~ 			Case $idRemember
;~ 				if GuiCtrlRead($idRemember)==$GUI_CHECKED Then
;~ 					$bRemember=True
;~ 					If $bWarnRemember Then
;~ 						$sMsg="Credentials will be stored in your user profile directory and will only be accessible from this account on this machine."
;~ 						MsgBox(48,"ohAuth - Remember Credentials",$sMsg,0,$hWnd)
;~ 						$bWarnRemember=False
;~ 					EndIf
;~ 				Else
;~ 					$bRemember=False
;~ 				EndIf
			Case $gidBtnValidate
				GUIRegisterMsg($WM_COMMAND,"")
				If Not $bReAuth Then GUICtrlSetState($gidOPID,$GUI_DISABLE)
				GUICtrlSetState($gidPass,$GUI_DISABLE)
;~ 				GUICtrlSetState($idRemember,$GUI_DISABLE)
				GUICtrlSetState($gidBtnValidate,$GUI_DISABLE)
				GUICtrlSetState($idBtnCancel,$GUI_DISABLE)
				_GUICtrlStatusBar_SetText($idStatus,"Validating...")
				$sUser=GUICtrlRead($gidOPID)
                $bNew=True
                For $i=0 To UBound($aAuth,1)-1
                    If StringCompare($aAuth[$i][0],$sUser)==0 Then
                        $bNew=False
                        ExitLoop
                    EndIf
                Next
                If $bNew Or $bReAuth Then
                    $sPass=GUICtrlRead($gidPass)
                    $iRet=_AD_Open("DS\"&$sUser, $sPass,"","",1)
                    If $iRet Then
                        GuiCtrlSetData($gidPass,$g_aAuthSalts[@HOUR])
                        _GUICtrlStatusBar_SetText($idStatus,"Validating...Success")
                        _AD_Close()
                        sleep(250)
                        _GUICtrlStatusBar_SetText($idStatus,"Encrypting...")
                        $sPass=_Base64Encode(_CryptProtectData($sPass))
                        _GUICtrlStatusBar_SetText($idStatus,"Encrypting...Done")
                        sleep(250)
                        _GUICtrlStatusBar_SetText($idStatus,"Saving Token...")
                        If $bReAuth Then
                        EndIf
                        $iMax=UBound($aAuth,1)
                        ReDim $aAuth[$iMax+1][3]
                        $aAuth[$iMax][0]=$sUser
                        $aAuth[$iMax][1]=$sPass
                        $aAuth[$iMax][2]=-1
                        ;IniWrite($sAuthIni,"ohAuth","opid/PAM",$aohAuth[0])
                        ;IniWrite($sAuthIni,"ohAuth","token",$aohAuth[1])
                        _saveAuth()
                        sleep(250)
                        _GUICtrlStatusBar_SetText($idStatus,"Saving Token...Done")
                        sleep(250)
                        _ctxReload()
                        GUIRegisterMsg($WM_COMMAND,"")
                        GUIDelete($hWnd)
                        Opt("GUIOnEventMode",1)
                        Return
                    Else
                        If @Error==8 Then
                            _GUICtrlStatusBar_SetText($idStatus,"Validating...Authentication Failed")
                        Else
                            _GUICtrlStatusBar_SetText($idStatus,"Validating...Internal Failure")
                        EndIf
                        GUIRegisterMsg($WM_COMMAND,"authWM_COMMAND")
                        If Not $bReAuth Then GUICtrlSetState($gidOPID,$GUI_ENABLE)
                        GUICtrlSetState($gidPass,$GUI_ENABLE)
    ;~ 					GUICtrlSetState($idRemember,$GUI_ENABLE)
                        GUICtrlSetState($gidBtnValidate,$GUI_ENABLE)
                        GUICtrlSetState($idBtnCancel,$GUI_ENABLE)
                        ContinueLoop
                    EndIf
                Else
                    MsgBox(48,"ohOverlay","Warning: This account already exists, please remove the exiting account first.")
                    GUIRegisterMsg($WM_COMMAND,"authWM_COMMAND")
                    If Not $bReAuth Then GUICtrlSetState($gidOPID,$GUI_ENABLE)
                    GUICtrlSetState($gidPass,$GUI_ENABLE)
;~ 					GUICtrlSetState($idRemember,$GUI_ENABLE)
                    GUICtrlSetState($gidBtnValidate,$GUI_ENABLE)
                    GUICtrlSetState($idBtnCancel,$GUI_ENABLE)
                    ContinueLoop
                EndIf
		EndSwitch
	WEnd
EndFunc

Global $bEnValidate[]=[False,False,False,False]
Global $gidOPID, $gidPass, $gidBtnValidate

Func authWM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    Local $iCode,$inID,$bMod=False
    $iId =  BitAND($wParam, 0xFFFF)
    $iCode = BitShift($wParam, 16)
    Switch $iId
        Case $gidPass
            Switch $iCode
                Case $EN_CHANGE
                    $bMod=True
                    $bEnValidate[0]=GUICtrlRead($iId)<>""
            EndSwitch
        Case $gidOPID
            Switch $iCode
                Case $EN_CHANGE
                    $bMod=True
                    $bEnValidate[1]=GUICtrlRead($iId)<>""
            EndSwitch
    EndSwitch
    If $bMod Then
        If $bEnValidate[0] And $bEnValidate[0]==$bEnValidate[1] Then
            $bEnValidate[3]=$GUI_ENABLE
        Else
            $bEnValidate[3]=$GUI_DISABLE

        EndIf

        if $bEnValidate[3]<>$bEnValidate[2] Then
            $bEnValidate[2]=$bEnValidate[3]
            GuiCtrlSetState($gidBtnValidate,$bEnValidate[3])
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func _saveAuth($bPurge=False)
    Local $aAuthDat,$aDiff[0],$bHas
    If $bPurge Then
        $aAuthDat=IniReadSection($gsConfig,"ohAuth")
        For $i=0 To UBound($aAuthDat,1)-1
            $bHas=False
            For $j=0 To UBound($aAuth,1)-1
                If $aAuth[$j][0]==$aAuthDat[$i][0] Then
                    $bHas=True
                    ExitLoop
                EndIf
            Next
            If Not $bHas Then IniDelete($gsConfig,"ohAuth",$aAuthDat[$i][0])
        Next
    EndIf
    For $i=0 To UBound($aAuth,1)-1
        IniWrite($gsConfig,"ohAuth",$aAuth[$i][0],$aAuth[$i][1])
    Next
EndFunc

Func _savePins()
    Local $sPins=""
    Local $iMax=UBound($aPins,1)
    For $i=0 To $iMax-1
        $sPins&=$aPins[$i][0]
        If $i<$iMax-1 Then $sPins&='|'
    Next
    IniWrite($gsConfig,"Config","Pins",$sPins)
EndFunc

Func _loadPins()
    Local $sIni=IniRead($gsConfig,"Config","Pins",""),$iMax
    If StringInStr($sIni,'|') Then
        Local $aTemp=StringSplit($sIni,'|')
        If @error Then Return SetError(1,0,0)
        Local $iMax
        For $i=1 To $aTemp[0]
            If StringStripCR(StringStripWS($aTemp[$i],15))=="" Then ContinueLoop
            $iMax=UBound($aPins,1)
            ReDim $aPins[$iMax+1][2]
            $aPins[$iMax][0]=$aTemp[$i]
        Next
    Else
        If StringStripCR(StringStripWS($sIni,15))=="" Then Return SetError(2,0,0)
        $iMax=UBound($aPins,1)
        ReDim $aPins[$iMax+1][2]
        $aPins[$iMax][0]=$sIni
    EndIf
    Return SetError(0,0,1)
EndFunc

Func _loadAuth()
    $aIni=IniReadSection($gsConfig,"ohAuth")
    For $i=1 To UBound($aIni,1)-1
        $iMax=UBound($aAuth,1)
        ReDim $aAuth[$iMax+1][3]
        $aAuth[$iMax][0]=$aIni[$i][0]
        $aAuth[$iMax][1]=$aIni[$i][1]
        $aAuth[$iMax][2]=-1
        If $aAuth[$iMax][1]==-1 Then ContinueLoop
        If Not _authValidate($aIni[$i][0],$aIni[$i][1]) Then
            $aAuth[$iMax][1]=-1
        EndIf
    Next
EndFunc

Func _checkAuth()
    Local $bUpdate=False
    For $i=0 To UBound($aAuth,1)-1
        If $aAuth[$i][1]==-1 Then ContinueLoop
        $iRet=_authValidate($aAuth[$i][0],$aAuth[$i][1])
        ConsoleWrite("checkAuth("&$aAuth[$i][0]&"):"&$iRet&@CRLF)
        If Not $iRet Then
            $aAuth[$i][1]=-1
            $bUpdate=True
        EndIf
    Next
    If $bUpdate Then _saveAuth()
EndFunc

Func _authValidate($sUser,$sToken)
    $sDesc=""
	$iRet=_AD_Open("DS\"&$sUser, _CryptUnprotectData(_Base64Decode($sToken),$sDesc),"","",1)
    If $iRet Then Return True
    Return False
EndFunc

Func _GuiRoundCorners($h_win, $i_x3, $i_y3)
    Local $XS_pos, $XS_reta, $XS_retb, $XS_ret2
    $XS_pos = WinGetPos($h_win)
    $XS_reta = _WinAPI_CreateRoundRectRgn ( 0, 0, $XS_pos[2], $XS_pos[3]-30, $i_x3, $i_y3 )
    $XS_retb = _WinAPI_CreateRectRgn (0, $XS_pos[2]-$XS_pos[1]/2, $XS_pos[2], $XS_pos[3] )
    $XS_retc = _WinAPI_CombineRgn ( $XS_reta, $XS_reta, $XS_retb, $RGN_OR  )
    _WinAPI_DeleteObject($XS_retb)
    _WinAPI_SetWindowRgn($h_win, $XS_reta)
EndFunc   ;==>_GuiRoundCorners

Func onDisplayChange($hWnd, $nMsgID, $wParam, $lParam)
    ConsoleWrite('Resolution changed to "' & @DesktopWidth & 'x' & @DesktopHeight & '".'&@CRLF)
    getMonInfo()
    posTrack()
    Return $GUI_RUNDEFMSG
EndFunc

Func getMonInfo()
    Local $aPos
    $aDisplays = _WinAPI_EnumDisplayMonitors()
    If Not IsArray($aDisplays) Then Return $GUI_RUNDEFMSG
    ReDim $aDisplays[$aDisplays[0][0] + 1][17]
    For $i = 1 To $aDisplays[0][0]
        $aPos = _WinAPI_GetPosFromRect($aDisplays[$i][1])
        For $j = 0 To 3
            $aDisplays[$i][$j + 1] = $aPos[$j]
        Next
        $aInfo = _WinAPI_GetMonitorInfo($aDisplays[$i][0])
        If IsArray($aInfo) Then
            For $j = 1 To 4
                $aDisplays[$i][$j + 4] = DllStructGetData($aInfo[0], $j)
            Next
            For $j = 1 To 4
                $aDisplays[$i][$j + 8] = DllStructGetData($aInfo[1], $j)
            Next
            $aDisplays[$i][13]=$aInfo[2]
            $aDisplays[$i][14]=$aInfo[3]
        EndIf
        $aDPI = _WinAPI_GetDpiForMonitor($aDisplays[$i][0], $MDT_ANGULAR_DPI);$MDT_DEFAULT)
        If IsArray($aDPI) Then
            $aDisplays[$i][15]=$aDPI[0]
            $aDisplays[$i][16]=$aDPI[1]
        EndIf
    Next
EndFunc

; Track mouse and update GUI position.
Func posTrack()
    Local $iTimer=TimerInit()
    Local $iMon=-1, $tPos = _WinAPI_GetMousePos()
    $aMousePos[0]=DllStructGetData($tPos,1)
    $aMousePos[2]=DllStructGetData($tPos,2)
    If $aMousePos[0]<>$aMousePos[1] Or $aMousePos[2]<>$aMousePos[3] Then
        $aMousePos[1]=$aMousePos[0]
        $aMousePos[3]=$aMousePos[2]
        Local $hMon=_WinAPI_MonitorFromPoint($tPos)
        For $i=1 To UBound($aDisplays,1)-1
            If $aDisplays[$i][0]==$hMon Then
                $iMon=$i
                ExitLoop
            EndIf
        Next
        If $iMon=-1 Then Return SetError(1,0,False)
        $iLeft=$aDisplays[$iMon][1]+$aDisplays[$iMon][3]-$iRight
        $iTop=$aDisplays[$iMon][2]+22+6
        If $iLeft<>$iLeftLast Or $iTop<>$iTopLast Then
            $iLeftLast=$iLeft
            $iTopLast=$iTop
            _watchDisplay()
        EndIf
    EndIf
EndFunc


Func _ToolTip($sMsg,$iPosX=Default,$iPosY=Default)
    Local $hDC,$hFont,$hOldFont,$aPos
    If $sMsg="" Then
        If IsHWnd($ghToolTip) Then _GUIToolTip_Destroy($ghToolTip)
        Return
    EndIf
    If Not IsHWnd($ghToolTip) Then
        $ghToolTip=_GUIToolTip_Create(0)
        $hDC=_WinAPI_GetDC(0)
        $hFont=_WinAPI_CreateFont(14, 0, 0, 0, 800, False, False, False, $DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, 'Consolas')
        $hOldFont=_WinAPI_SelectObject($hDC, $hFont)
        _SendMessage($ghToolTip, $WM_SETFONT, $hFont, True)
        _GUIToolTip_AddTool($ghToolTip, 0, $sMsg, 0, 0, 0, 0, 0, BitOR($TTF_TRACK, $TTF_ABSOLUTE))
    EndIf
    If $iPosX=Default Or $iPosY=Default Then
        $aPos=MouseGetPos()
        If $iPosX=Default Then $iPosX=$aPos[0]+16
        If $iPosY=Default Then $iPosY=$aPos[1]+16
    EndIf
    _GUIToolTip_UpdateTipText($ghToolTip,0,0,$sMsg)
    _GUIToolTip_TrackPosition($ghToolTip,$iPosX,$iPosY)
    _GUIToolTip_TrackActivate($ghToolTip,True,0,0)
EndFunc

getMonInfo()
_loadAuth()
_loadPins()
initGui()

While Sleep(1)
WEnd

