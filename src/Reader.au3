#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\assets\epub.ico
#AutoIt3Wrapper_Outfile=..\bin\ePUB Reader.exe
#AutoIt3Wrapper_Res_Description=ePUB Reader
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductName=ePUB Reader
#AutoIt3Wrapper_Res_CompanyName=Andreik
#AutoIt3Wrapper_Res_LegalCopyright=© 2024 Andreik (AutoIt Forum)
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 4 -w 5 -w 6 -w 7
#Au3Stripper_Parameters=/sf /sv /mo /rm /rsln
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <WinAPISysWin.au3>
#include <WinAPIConv.au3>
#include <ButtonConstants.au3>
#include <SQLite.au3>
#include <GDIPlus.au3>
#include <Memory.au3>
#include <IE.au3>
#include <File.au3>
#include <Misc.au3>
#include "Def.au3"
#include "ePub.au3"
#include "Slider.au3"
#include "Resources.au3"

_SQLite_Startup()
_GDIPlus_Startup()

LoadSettings()

$oIE = _IECreateEmbedded()
$oIEEvents = ObjEvent($oIE, "_IEEvent_", "DWebBrowserEvents2")

$hMain = GUICreate($sAppTitle, 1200, 735 + GetTitleBarHeight(), Default, Default, BitOR($WS_SIZEBOX, $WS_MAXIMIZEBOX, $WS_MINIMIZEBOX))
$cReader = GUICtrlCreateObj($oIE, 10, 10, 1180, 665)
For $Index = 0 To 5
    $aCtrlButtons[$Index] = GUICtrlCreateButton('', $Index * 60 + 10, 680, 50, 50, BitOR($BS_BITMAP, $BS_CENTER))
    GUICtrlSetTip($aCtrlButtons[$Index], $aTips[$Index])
    GUICtrlSetCursor($aCtrlButtons[$Index], 0)
Next
BitmapToCtrl($aCtrlButtons[0], EPUB_Icon())
BitmapToCtrl($aCtrlButtons[1], Back_Icon(True))
BitmapToCtrl($aCtrlButtons[2], Home_Icon())
BitmapToCtrl($aCtrlButtons[3], Stop_Icon())
BitmapToCtrl($aCtrlButtons[4], Back_Icon())
BitmapToCtrl($aCtrlButtons[5], Refresh_Icon())

$mSlider = CreateSlider($hMain, 370, 680, 820, 50, $aColors[0], $aColors[1], $aColors[2])
BitmapToCtrl($mSlider['Prev'], Next_Icon(True))
BitmapToCtrl($mSlider['Next'], Next_Icon())
BitmapToCtrl($mSlider['Config'], Settings_Icon())
SetSliderLabel($mSlider, 'Operation: Idle')

_IENavigate($oIE, 'about:blank')
GUISetState(@SW_SHOW, $hMain)

GUIRegisterMsg($WM_SIZING, 'WM_SIZING')
GUIRegisterMsg($WM_SIZE, 'WM_SIZE')

_IEDocWriteHTML($oIE, GetDefaultHTML())

If $CmdLine[0] Then
    $mEPUB = Load_ePUB($CmdLine[1])
    If Not @error Then SetSliderLabel($mSlider, 'Operation: Idle')
EndIf

While True
    $iMsg = GUIGetMsg()
    Switch $iMsg
        Case $aCtrlButtons[0]
            $mEPUB = SelectEPUB()
        Case $aCtrlButtons[1]
            NavigateBackward()
        Case $aCtrlButtons[2]
            LoadFirstChapter()
        Case $aCtrlButtons[3]
            StopLoading()
        Case $aCtrlButtons[4]
            NavigateForward()
        Case $aCtrlButtons[5]
            RefreshChapter()
        Case $mSlider['Config']
            Settings()
        Case $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESTORE
			If IsObj(_IEGetObjById($oIE, 'epub-center')) Then
				Sleep(250)
				_IEAction($oIE, 'refresh')
			Else
				_IENavigate($oIE, _IEPropertyGet($oIE, 'locationurl'))
                If IsMap($mEPUB) Then
                    If $mSlider['ActiveChapter'] = 1 Then InsertCSS(FixCoverCSS())
                EndIf
			EndIf
        Case $GUI_EVENT_CLOSE
            ExitLoop
    EndSwitch
    If IsMap($mEPUB) Then
        Switch SliderEvent($mSlider, $iMsg)
            Case $CHAPTER_CHANGED, $PREV_CHAPTER, $NEXT_CHAPTER
                LoadChapter($mEPUB, $mSlider['ActiveChapter'])
        EndSwitch
    EndIf
WEnd

SetSliderLabel($mSlider, 'Operation: Release resources')
ePUB_Release($mEPUB, $aCollector)
If $fDeleteRes Then CleanUpWorkspace()
If $fDeleteCfg Then DeleteConfig()
_GDIPlus_Shutdown()
_SQLite_Shutdown()

Func SelectEPUB()
    SetSliderLabel($mSlider, 'Operation: Select ePUB file')
    Local $sPath = FileOpenDialog('Select ePUB', @ScriptDir, 'ePUB Book (*.epub)', 3)
    If $sPath Then
        $sPath &= (StringRight($sPath, 5) <> '.epub' ? '.epub' : '')
        Local $mEPUB = Load_ePUB($sPath)
        If Not @error Then SetSliderLabel($mSlider, 'Operation: Idle')
    Else
        SetSliderLabel($mSlider, 'Operation: Idle')
    EndIf
    Return $mEPUB
EndFunc

Func NavigateBackward()
    If IsMap($mEPUB) Then
        If $fAllowBack Then
            SetSliderLabel($mSlider, 'Operation: Navigate backward')
            _IEAction($oIE, 'back')
            UpdateSlider()
        Else
            SetSliderLabel($mSlider, 'Error: Not allowed')
        EndIf
    EndIf
EndFunc

Func LoadFirstChapter()
    If IsMap($mEPUB) Then
        SetActiveChapter($mSlider, 1)
        LoadChapter($mEPUB, $mSlider['ActiveChapter'])
    EndIf
EndFunc

Func StopLoading()
    If IsMap($mEPUB) Then
        SetSliderLabel($mSlider, 'Operation: Stop loading')
        _IEAction($oIE, 'stop')
        SetSliderLabel($mSlider, 'Operation: Idle')
    EndIf
EndFunc

Func RefreshChapter()
    If IsMap($mEPUB) Then
        SetSliderLabel($mSlider, 'Operation: Refresh current chapter')
        _IENavigate($oIE, _IEPropertyGet($oIE, 'locationurl'))
        SetSliderLabel($mSlider, 'Operation: Idle')
    EndIf
EndFunc

Func NavigateForward()
    If IsMap($mEPUB) Then
        If $fAllowForward Then
            SetSliderLabel($mSlider, 'Operation: Navigate forward')
            _IEAction($oIE, 'forward')
            UpdateSlider()
        Else
            SetSliderLabel($mSlider, 'Error: Not allowed')
        EndIf
    EndIf
EndFunc

Func _IEEvent_NavigateComplete2($oIEpDisp, $sIEURL)
    #forceref $oIEpDisp, $sIEURL
	InsertCSS($sCustomCSS)
    If IsMap($mEPUB) Then UpdateSlider()
EndFunc

Func _IEEvent_CommandStateChange($sCommand, $fEnable)
    Switch $sCommand
        Case 0x01
            $fAllowForward = $fEnable
        Case 0x02
			$fAllowBack = $fEnable
    EndSwitch
EndFunc

Func UpdateSlider()
    _IELoadWait($oIE)
    Local $sLocation = _IEPropertyGet($oIE, 'locationname')
    Local $iLen = StringLen($sLocation)
    Local $aChapters = $mEPUB['Chapters']
    For $Index = 1 To $mEPUB['NumOfChapters']
        If StringRight($aChapters[$Index][2], $iLen) = $sLocation Then
            SetActiveChapter($mSlider, $Index)
            ExitLoop
        EndIf
    Next
    SetSliderLabel($mSlider, 'Operation: Idle')
EndFunc

Func Load_ePUB($sPath)
    If IsMap($mEPUB) Then
        SetSliderLabel($mSlider, 'Operation: Release previous ePUB')
        WinSetTitle($hMain, '' , $sAppTitle)
        ePUB_Release($mEPUB, $aCollector, True)
        ResetSlider($mSlider)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Initialize ePUB')
    Local $mEPUB = ePUB_Init($sPath)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_INIT, @error))
        Return
    EndIf
    SetSliderLabel($mSlider, 'Operation: Get files')
    Local $aFile = ePUB_GetFiles($mEPUB)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_GETFILES, @error))
        Return
    EndIf
    SetSliderLabel($mSlider, 'Operation: Extract content')
    ePUB_Extract($mEPUB, $aFile)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_EXTRACT, @error))
        Return
    EndIf
    SetSliderLabel($mSlider, 'Operation: Validate ePUB')
    If Not ePUB_Validate($mEPUB) Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_VALIDATE, @error))
        Return
    EndIf
    SetSliderLabel($mSlider, 'Operation: Get OPF')
    Local $aOPF = ePUB_GetOPF($mEPUB)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_GETOPF, @error))
        Return
    EndIf
    SetSliderLabel($mSlider, 'Operation: Read OPF')
    Local $aChapters = ePUB_ReadOPF($mEPUB, $aOPF)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_READOPF, @error))
        Return
    Else
        $mEPUB['NumOfChapters'] = @extended
    EndIf
    Local $aMetadata = SQLite_Query($mEPUB['DB'], 'SELECT value FROM metadata WHERE meta = "title";')
    If @extended = 1 Then WinSetTitle($hMain, '' , $sAppTitle & ' - ' & StringStripWS($aMetadata[1][0], 7))
    If $mEPUB['NumOfChapters'] Then
        $mEPUB['Chapters'] = $aChapters
        SetSliderLimits($mSlider, 1, $mEPUB['NumOfChapters'])
        $mSlider['Max'] = $mEPUB['NumOfChapters']
        SetActiveChapter($mSlider, 1)
        LoadChapter($mEPUB, $mSlider['ActiveChapter'])
        For $Index = 1 To $mSlider['Max']
            GUICtrlSetTip($mSlider[$Index], $aChapters[$Index][3])
        Next
    EndIf
    Return $mEPUB
EndFunc

Func LoadChapter($mEPUB, $iChapter)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    SetSliderLabel($mSlider, 'Operation: Load chapter')
    If $iChapter < 1 Or $iChapter > $mEPUB['NumOfChapters'] Then
        SetSliderLabel($mSlider, 'Error: Invalid chapter index')
        Return SetError(2, 0, False)
    EndIf
    Local $aChapters = $mEPUB['Chapters']
    Local $sChapter = $aChapters[$iChapter][2]
    Local $aConfig = SQLite_Query($mEPUB['DB'], 'SELECT value FROM config WHERE cfg="path" LIMIT 1;')
    If @extended = 1 Then
        Local $sConfig = ($aConfig[1][0] ? '\' & $aConfig[1][0] : '')
        If StringRight($sChapter, 4) = '.xml' Then
            Local $sSource = $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & $sConfig & '\' & $sChapter
            Local $sDestination = $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & $sConfig & '\' & StringTrimRight($sChapter, 4) & '__fix.xhtml'
            FileCopy($sSource, $sDestination)
            _IENavigate($oIE, 'file:///' & StringReplace($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & $sConfig & '\' & StringTrimRight($sChapter, 4) & '__fix.xhtml', '\', '/'))
        Else
            _IENavigate($oIE, 'file:///' & StringReplace($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & $sConfig & '\' & $sChapter, '\', '/'))
        EndIf
        _IELoadWait($oIE)
        If $iChapter = 1 Then InsertCSS(FixCoverCSS())
    EndIf
    SetSliderLabel($mSlider, 'Operation: Idle')
EndFunc

Func WM_SIZING($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
    If $hWnd = $hMain Then
        Local $tRECT = DllStructCreate('long Left;long Top;long Right;long Bottom;', $lParam)
        If $tRECT.Right - $tRECT.Left < 800 Then $tRECT.Right = $tRECT.Left + 800
        If $tRECT.Bottom - $tRECT.Top < 300 Then $tRECT.Bottom = $tRECT.Top + 300
        Return True
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func WM_SIZE($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam
	If $hWnd = $hMain Then
		ResizeControls(BitAND($lParam, 0xFFFF), BitShift($lParam, 16))
		Return True
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc

Func BitmapToCtrl($cCtrl, $bData)
	Local $hHBITMAP = _GDIPlus_BitmapCreateFromMemory($bData, True)
	_WinAPI_DeleteObject(GUICtrlSendMsg($cCtrl, 0x00F7, 0, $hHBITMAP))
	_WinAPI_DeleteObject($hHBITMAP)
EndFunc

Func ResizeControls($iWidth, $iHeight)
	GUICtrlSetPos($cReader, 10, 10, $iWidth - 20 , $iHeight - 70)
    For $Index = 0 To 5
        GUICtrlSetPos($aCtrlButtons[$Index], $Index * 60 + 10, $iHeight - 55, 50, 50)
    Next
    GUICtrlSetPos($mSlider['Frame'], 370, $iHeight - 55, $iWidth - 500, $mSlider['ElementHeigth'])
    GUICtrlSetPos($mSlider['Label'], 370, $iHeight - 55 + $mSlider['ElementHeigth'], $iWidth - $mSlider['ElementHeigth'] - 510, $mSlider['ElementHeigth'])
    GUICtrlSetPos($mSlider['Prev'], $iWidth - 120, $iHeight - 55, $mSlider['ElementHeigth'] * 2, $mSlider['ElementHeigth'] * 2)
    GUICtrlSetPos($mSlider['Next'], $iWidth - 60, $iHeight - 55, $mSlider['ElementHeigth'] * 2, $mSlider['ElementHeigth'] * 2)
    GUICtrlSetPos($mSlider['Config'], $iWidth - $mSlider['ElementHeigth'] - 130, $iHeight - 55 + $mSlider['ElementHeigth'], $mSlider['ElementHeigth'], $mSlider['ElementHeigth'])
    $mSlider['W'] = $iWidth - 500
    $mSlider['Y'] = $iHeight - 55
    If $mSlider['Min'] <> Null And $mSlider['Max'] <> Null Then
        Local $iLocalMin = $mSlider['Min']
        Local $iLocalMax = $mSlider['Max']
        $iLocalMax -= ($iLocalMin <> 1 ? $iLocalMin - 1 : 0)
        $iLocalMin = 1
        Local $iSegLen = Int($mSlider['W'] / $iLocalMax)
        Local $iLeft = Mod($mSlider['W'], $iLocalMax)
        Local $iPos, $iSegWidth, $iDelta = $mSlider['Min'] - $iLocalMin, $iCurrentChapter
        For $Index = $iLocalMin To $iLocalMax
            $iPos = $mSlider['X'] + (($Index - 1) * $iSegLen) + ($iLocalMax - $Index < $iLeft ? $iLeft - ($iLocalMax - $Index) - 1 : 0)
            $iSegWidth = $iSegLen + ($iLocalMax - $Index < $iLeft ? 1 : 0)
            $iCurrentChapter = $Index + $iDelta
            GUICtrlSetPos($mSlider[$iCurrentChapter], $iPos, $mSlider['Y'], $iSegWidth, $mSlider['ElementHeigth'])
        Next
    EndIf
EndFunc

Func InsertCSS($sCSS)
	Local $oHead = _IETagNameGetCollection($oIE, 'head', 0)
	If Not @error Then _IEDocInsertHTML($oHead, '<style>' & $sCSS & '</style>', 'beforeend')
EndFunc

Func GetCustomCSS()
    Return 'html, body{margin:1rem; padding:0;}'
EndFunc

Func FixCoverCSS()
    Return 'svg{position:fixed; top:0; bottom:0; left:0; right:0}'
EndFunc

Func GetDefaultCSS()
    Local $sDefaultCSS = 'html, body {margin: 0; padding: 0; width: 100%; height: 100%; background-color: #212529; cursor:default;} '
    $sDefaultCSS &= '#wrap {position: relative; width: 100%;height: 100%;display: table;} '
    $sDefaultCSS &= '#mid {*position: absolute; *top: 50%; *width: 100%; *text-align: center;} '
    $sDefaultCSS &= '#epub-center {*position: relative; *top: -50%; font-size: 30px;font-family: "Segoe UI"; color: #dee2e6;}'
    Return $sDefaultCSS
EndFunc

Func GetDefaultHTML()
    Local $sDefaultHTML = '<!DOCTYPE html><html lang="en"><head><style>' & GetDefaultCSS() & '</style></head>'
    $sDefaultHTML &= '<body onselectstart="return false;"><div id="wrap"><div id="mid">'
    $sDefaultHTML &= '<div id="epub-center">Load your ePUB</div></div></div></body></html>'
    Return $sDefaultHTML
EndFunc

Func GetDefaultColor($iType)
    Switch $iType
        Case 0
            Return 0x242423
        Case 1
            Return 0x333333
        Case 2
            Return 0x9e2a2b
    EndSwitch
EndFunc

Func LoadSettings($fApply = False)
    Local $sCfgDB = $__sWorkingDir & '\' & $__sInternalName & '\config.sqlite'
    If FileExists($sCfgDB) Then
        Local $mSettings[]
        Local $hDB = _SQLite_Open($sCfgDB)
        Local $aSettings = SQLite_Query($hDB, 'SELECT property, value FROM settings;')
        For $Index = 1 To @extended
            $mSettings[$aSettings[$Index][0]] = $aSettings[$Index][1]
        Next
        $sCustomCSS = (MapExists($mSettings, 'CSS') ? $mSettings['CSS'] : GetCustomCSS())
        $aColors[0] = (MapExists($mSettings, 'PrimaryColor') ? $mSettings['PrimaryColor'] : GetDefaultColor(0))
        $aColors[1] = (MapExists($mSettings, 'AlternateColor') ? $mSettings['AlternateColor'] : GetDefaultColor(1))
        $aColors[2] = (MapExists($mSettings, 'SelectionColor') ? $mSettings['SelectionColor'] : GetDefaultColor(2))
        $fDeleteRes = (MapExists($mSettings, 'DeleteRes') ? ($mSettings['DeleteRes'] = 'True' ? True : False) : False)
        _SQLite_Close($hDB)
    Else
        $sCustomCSS = GetCustomCSS()
        $aColors[0] = GetDefaultColor(0)
        $aColors[1] = GetDefaultColor(1)
        $aColors[2] = GetDefaultColor(2)
        $fDeleteRes = False
    EndIf
    If $fApply Then ApplySettings()
EndFunc

Func SaveSettings($sCSS, $iPrimaryColor, $iAlternateColor, $iSelectionColor, $fDeleteWS)
    Local $sPrimaryColor = String('0x' & Hex($iPrimaryColor, 6))
    Local $sAlternateColor = String('0x' & Hex($iAlternateColor, 6))
    Local $sSelectionColor = String('0x' & Hex($iSelectionColor, 6))
    Local $sDelete = ($fDeleteWS ? 'True' : 'False')
    Local $sPrefix = 'INSERT INTO settings(property, value) VALUES'
    Local $sConflict = ' ON CONFLICT(property) DO UPDATE SET value = '
    Local $bCreateTable = (FileExists($__sWorkingDir & '\' & $__sInternalName & '\config.sqlite') ? False : True)
    Local $hDB = _SQLite_Open($__sWorkingDir & '\' & $__sInternalName & '\config.sqlite')
    If $bCreateTable Then _SQLite_Exec($hDB, 'CREATE TABLE settings(property VARCHAR(32) PRIMARY KEY, value TEXT);')
    _SQLite_Exec($hDB, $sPrefix & '("CSS", ' & _SQLite_FastEscape($sCSS) & ')' & $sConflict & _SQLite_FastEscape($sCSS) & ';')
    _SQLite_Exec($hDB, $sPrefix & '("PrimaryColor", ' & _SQLite_FastEscape($sPrimaryColor) & ')' & $sConflict & _SQLite_FastEscape($sPrimaryColor) & ';')
    _SQLite_Exec($hDB, $sPrefix & '("AlternateColor", ' & _SQLite_FastEscape($sAlternateColor) & ')' & $sConflict & _SQLite_FastEscape($sAlternateColor) & ';')
    _SQLite_Exec($hDB, $sPrefix & '("SelectionColor", ' & _SQLite_FastEscape($sSelectionColor) & ')' & $sConflict & _SQLite_FastEscape($sSelectionColor) & ';')
    _SQLite_Exec($hDB, $sPrefix & '("DeleteRes", ' & _SQLite_FastEscape($sDelete) & ')' & $sConflict & _SQLite_FastEscape($sDelete) & ';')
    _SQLite_Close($hDB)
EndFunc

Func ApplySettings()
    If IsMap($mSlider) Then
        $mSlider['Color'] = $aColors[0]
        $mSlider['AltColor'] = $aColors[1]
        $mSlider['ActiveColor'] = $aColors[2]
        Local $iLocalMin = $mSlider['Min']
        Local $iLocalMax = $mSlider['Max']
        $iLocalMax -= ($iLocalMin <> 1 ? $iLocalMin - 1 : 0)
        $iLocalMin = 1
        Local $iCurrentChapter
        For $Index = $iLocalMin To $iLocalMax
            $iCurrentChapter = $Index + $mSlider['Min'] - $iLocalMin
            GUICtrlSetBkColor($mSlider[$iCurrentChapter], Mod($Index, 2) ? $mSlider['Color'] : $mSlider['AltColor'])
        Next
        SetActiveChapter($mSlider, $mSlider['ActiveChapter'])
    EndIf
    If IsObj($oIE) Then InsertCSS($sCustomCSS)
EndFunc

Func Settings()
    Local $iTemp, $sCSS, $fDeleteWS, $iPrimaryColor = $aColors[0], $iAlternateColor = $aColors[1], $iSelectionColor = $aColors[2]
    Local $hGUI = GUICreate('Settings', 405, 360 + GetTitleBarHeight(), Default, Default, $DS_SETFOREGROUND, Default, $hMain)
    GUICtrlCreateTab(10, 10, 380, 300)
    GUICtrlCreateTabItem('CSS')
    Local $cCSSEdit = GUICtrlCreateEdit($sCustomCSS, 20, 40, 360, 260, BitOR($WS_VSCROLL, 0x0004))
    GUICtrlCreateTabItem('Colors')
    Local $cPrimaryColorLabel = GUICtrlCreateLabel('Slider primary color', 30, 40, 140, 25, 0x200)
    Local $cAlternateColorLabel = GUICtrlCreateLabel('Slider alternate color', 30, 65, 140, 25, 0x200)
    Local $cSelectionColorLabel = GUICtrlCreateLabel('Slider selection color', 30, 90, 140, 25, 0x200)
    Local $cPrimaryColor = GUICtrlCreateLabel('', 170, 42, 25, 20)
    Local $cAlternateColor = GUICtrlCreateLabel('', 170, 67, 25, 20)
    Local $cSelectionColor = GUICtrlCreateLabel('', 170, 92, 25, 20)
    GUICtrlCreateTabItem('Workspace')
    Local $cSizeLabel = GUICtrlCreateLabel('Current workspace size', 30, 40, 150, 25, 0x200)
    Local $cSize = GUICtrlCreateLabel(GetWorkspaceSize(), 190, 40, 100, 25, 0x200)
    Local $cDeleteRes = GUICtrlCreateCheckbox('Clean up workspace', 30, 65, 140, 25)
    Local $cDeleteCfg = GUICtrlCreateCheckbox('Delete application settings on exit', 30, 90, 230, 25)
    GUICtrlCreateTabItem('Info')
    Local $cVersion = GUICtrlCreateLabel('Application version: ' & $sVersion, 30, 40, 300, 25, 0x200)
    Local $cAuthor = GUICtrlCreateLabel('Author: Andreik (AutoIt Forum)', 30, 65, 300, 25, 0x200)
    Local $cCredits = GUICtrlCreateEdit('', 30, 90, 340, 200, BitOR($WS_VSCROLL, 0x0800))
    GUICtrlCreateTabItem('')
    Local $cSave = GUICtrlCreateButton('Save', 10, 320, 120, 30)
    Local $cReset = GUICtrlCreateButton('Reset', 140, 320, 120, 30)
    Local $cCancel = GUICtrlCreateButton('Cancel', 270, 320, 120, 30)
    Local $sCredits = 'Epub icons created by shohanur.rahman13 - Flaticon (https://www.flaticon.com/free-icons/epub)' & @CRLF & @CRLF
    $sCredits &= 'Start icons created by hqrloveq - Flaticon (https://www.flaticon.com/free-icons/start)' & @CRLF & @CRLF
    $sCredits &= 'Home button icons created by hqrloveq - Flaticon (https://www.flaticon.com/free-icons/home-button)' & @CRLF & @CRLF
    $sCredits &= 'Ui icons created by Dewi Sari - Flaticon (https://www.flaticon.com/free-icons/ui)' & @CRLF & @CRLF
    $sCredits &= 'Refresh icons created by Freepik - Flaticon (https://www.flaticon.com/free-icons/refresh)' & @CRLF & @CRLF
    $sCredits &= 'Next icons created by KP Arts - Flaticon (https://www.flaticon.com/free-icons/next)' & @CRLF & @CRLF
    $sCredits &= 'Settings icons created by Freepik - Flaticon (https://www.flaticon.com/free-icons/settings)' & @CRLF
    GUICtrlSetData($cCredits, $sCredits)
    GUICtrlSetFont($cSizeLabel, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cSize, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cDeleteRes, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cDeleteCfg, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cCSSEdit, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cCredits, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cVersion, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cAuthor, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cPrimaryColorLabel, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cAlternateColorLabel, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cSelectionColorLabel, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cSave, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cReset, 11, 500, 0, 'Segoe UI')
    GUICtrlSetFont($cCancel, 11, 500, 0, 'Segoe UI')
    GUICtrlSetBkColor($cPrimaryColor, $aColors[0])
    GUICtrlSetBkColor($cAlternateColor, $aColors[1])
    GUICtrlSetBkColor($cSelectionColor, $aColors[2])
    GUICtrlSetBkColor($cCredits, 0xFFFFFF)
    GUICtrlSetCursor($cPrimaryColor, 0)
    GUICtrlSetCursor($cAlternateColor, 0)
    GUICtrlSetCursor($cSelectionColor, 0)
    GUICtrlSetState($cDeleteRes, ($fDeleteRes ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlSetState($cDeleteCfg, ($fDeleteCfg ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUISetState(@SW_SHOW, $hGUI)
    While True
        Switch GUIGetMsg()
            Case $cPrimaryColor
                $iTemp = _ChooseColor(2, $iPrimaryColor, 2, $hGUI)
                If Not @error Then
                    $iPrimaryColor = $iTemp
                    GUICtrlSetBkColor($cPrimaryColor, $iPrimaryColor)
                EndIf
            Case $cAlternateColor
                $iTemp = _ChooseColor(2, $iAlternateColor, 2, $hGUI)
                If Not @error Then
                    $iAlternateColor = $iTemp
                    GUICtrlSetBkColor($cAlternateColor, $iAlternateColor)
                EndIf
            Case $cSelectionColor
                $iTemp = _ChooseColor(2, $iSelectionColor, 2, $hGUI)
                If Not @error Then
                    $iSelectionColor = $iTemp
                    GUICtrlSetBkColor($cSelectionColor, $iSelectionColor)
                EndIf
            Case $cSave
                $sCSS = GUICtrlRead($cCSSEdit)
                $fDeleteWS = (GUICtrlRead($cDeleteRes) = $GUI_CHECKED ? True : False)
                $fDeleteCfg = (GUICtrlRead($cDeleteCfg) = $GUI_CHECKED ? True : False)
                SaveSettings($sCSS, $iPrimaryColor, $iAlternateColor, $iSelectionColor, $fDeleteWS)
                LoadSettings(True)
                ExitLoop
            Case $cReset
                $iPrimaryColor = GetDefaultColor(0)
                $iAlternateColor = GetDefaultColor(1)
                $iSelectionColor = GetDefaultColor(2)
                $sCSS = GetCustomCSS()
                GUICtrlSetBkColor($cPrimaryColor, $iPrimaryColor)
                GUICtrlSetBkColor($cAlternateColor, $iAlternateColor)
                GUICtrlSetBkColor($cSelectionColor, $iSelectionColor)
                GUICtrlSetData($cCSSEdit, $sCSS)
                GUICtrlSetState($cDeleteRes, $GUI_UNCHECKED)
                GUICtrlSetState($cDeleteCfg, $GUI_UNCHECKED)
            Case $cCancel
                ExitLoop
        EndSwitch
    WEnd
    WinActivate($hMain)
    GUIDelete($hGUI)
EndFunc

Func GetWorkspaceSize()
    Local $aFile = _FileListToArrayRec($__sWorkingDir & '\' & $__sInternalName, Default, 1, 1, 0, 2)
    If IsArray($aFile) Then
        Local $iBytes = 0
        For $Index = 1 To $aFile[0]
            $iBytes += FileGetSize($aFile[$Index])
        Next
        If $iBytes < 1024 Then
            Return $iBytes & ' Bytes'
        Else
            $iBytes = Round($iBytes / 1024, 2)
            If $iBytes >= 1024 Then
                Return Round($iBytes / 1024, 2) & ' MB'
            Else
                Return $iBytes & ' KB'
            EndIf
        EndIf
    EndIf
    Return '0 Bytes'
EndFunc

Func CleanUpWorkspace()
    Local $aFile = _FileListToArray($__sWorkingDir & '\' & $__sInternalName, Default, Default, True)
    If IsArray($aFile) Then
        For $Index = 1 To $aFile[0]
            If $aFile[$Index] <> $__sWorkingDir & '\' & $__sInternalName & '\config.sqlite' Then
                If StringInStr(FileGetAttrib($aFile[$Index]), 'D') > 0 Then
                    DirRemove($aFile[$Index], 1)
                Else
                    FileDelete($aFile[$Index])
                EndIf
            EndIf
        Next
    EndIf
EndFunc

Func DeleteConfig()
    Local $sCfgDB = $__sWorkingDir & '\' & $__sInternalName & '\config.sqlite'
    If FileExists($sCfgDB) Then FileDelete($sCfgDB)
EndFunc

Func GetTitleBarHeight()
    Local $hGUI = GUICreate('', 100, 100, 0, 0, BitOR($WS_SIZEBOX, $WS_MAXIMIZEBOX, $WS_MINIMIZEBOX))
    Local $tWindow = _WinAPI_GetWindowRect($hGUI)
    Local $tClient = _WinAPI_GetClientRect($hGUI)
    Local $tPoint = DllStructCreate('long X;long Y;')
    $tPoint.X = $tClient.Left
    $tPoint.Y = $tClient.Top
    _WinAPI_ClientToScreen($hGUI, $tPoint)
    GUIDelete($hGUI)
    Return $tPoint.Y - $tWindow.Top
EndFunc

Func GetErrorMessage($iErrorType, $iError, $iExtended = 0)
    Switch $iErrorType
        Case $EPUB_INIT
            Switch $iError
                Case 1
                    Return 'Invalid file path'
                Case 2
                    Return 'Invalid file format'
                Case 3
                    Return 'Failed to create Shell Application object'
                Case 4
                    Switch $iExtended
                        Case 1
                            Return 'Failed to open internal database'
                        Case 2
                            Return 'Failed to create config table'
                        Case 3
                            Return 'Failed to create manifest table'
                        Case 4
                            Return 'Failed to create spine table'
                        Case 5
                            Return 'Failed to create guides table'
                        Case 6
                            Return 'Failed to create metadata table'
                        Case 7
                            Return 'Failed to create IDX_Manifest index'
                        Case 8
                            Return 'Failed to create IDX_Guides index'
                    EndSwitch
                Case 5
                    Return 'Failed to create unique temporary directory'
                Case 6
                    Return 'Failed to move ePUB in temporary directory'
            EndSwitch
        Case $EPUB_GETFILES
            Switch $iError
                Case 1
                    Return 'Invalid ePUB structure'
                Case 2
                    Return 'Failed to locate ePUB file'
                Case 3
                    Return 'Empty ePUB file'
            EndSwitch
        Case $EPUB_EXTRACT
            Switch $iError
                Case 1
                    Return 'Invalid ePUB structure'
                Case 2
                    Return 'Invalid file structure'
                Case 3
                    Return 'Failed to locate ePUB file'
                Case 4
                    Return 'Empty ePUB file'
            EndSwitch
        Case $EPUB_VALIDATE
            Switch $iError
                Case 1
                    Return 'Invalid ePUB structure'
                Case 2
                    Return 'Failed to locate mimetype file'
                Case 3
                    Return 'Invalid mimetype'
            EndSwitch
        Case $EPUB_GETOPF
            Switch $iError
                Case 1
                    Return 'Invalid ePUB structure'
                Case 2
                    Return 'Failed to locate container file'
                Case 3
                    Return 'Failed to detect OPF location'
                Case 4
                    Return 'Failed to save OPF location'
            EndSwitch
        Case $EPUB_READOPF
            Switch $iError
                Case 1
                    Return 'Invalid ePUB structure'
                Case 2
                    Return 'Invalid OPF structure'
                Case 3
                    Return 'Invalid OPF location'
                Case 4
                    Return 'Failed to read ePUB version'
                Case 5
                    Return 'Invalid ePUB version'
                Case 6
                    Return 'Failed to parse metadata'
                Case 7
                    Return 'Failed to parse references'
                Case 8
                    Return 'Failed to parse manifest'
            EndSwitch
    EndSwitch
EndFunc
