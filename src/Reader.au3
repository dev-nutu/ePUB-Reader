#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\assets\epub.ico
#AutoIt3Wrapper_Outfile=..\bin\ePUB Reader.exe
#AutoIt3Wrapper_Res_Description=ePUB Reader
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
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

$mSlider = CreateSlider($hMain, 370, 680, 820, 50, $aColors)
BitmapToCtrl($mSlider['Prev'], Next_Icon(True))
BitmapToCtrl($mSlider['Next'], Next_Icon())
BitmapToCtrl($mSlider['Config'], Settings_Icon())
BitmapToCtrl($mSlider['Jump'], Jump_Icon())
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
        Case $mSlider['Jump']
            JumpToChapter()
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
    Local $aChapters = $mEPUB['Chapters']
    Local $aMatch
    For $Index = 1 To $mEPUB['NumOfChapters']
        $aMatch = StringRegExp($sLocation, '(?:.*)(' & EscapeRegEx($aChapters[$Index][2]) & ')(?:.*)', 3)
        If IsArray($aMatch) Then
            SetActiveChapter($mSlider, $Index)
            ExitLoop
        EndIf
    Next
    SetSliderLabel($mSlider, 'Operation: Idle')
EndFunc

Func EscapeRegEx($sText)
    Return StringRegExpReplace($sText, '[\.\^\$\*\+\?\(\)\[\{\\\|]', '\\$0')
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
        Return SetError(1, 0, False)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Get files')
    Local $aFile = ePUB_GetFiles($mEPUB)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_GETFILES, @error))
        Return SetError(2, 0, False)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Extract content')
    ePUB_Extract($mEPUB, $aFile)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_EXTRACT, @error))
        Return SetError(3, 0, False)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Validate ePUB')
    If Not ePUB_Validate($mEPUB) Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_VALIDATE, @error))
        Return SetError(4, 0, False)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Get OPF')
    Local $aOPF = ePUB_GetOPF($mEPUB)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_GETOPF, @error))
        Return SetError(5, 0, False)
    EndIf
    SetSliderLabel($mSlider, 'Operation: Read OPF')
    Local $aChapters = ePUB_ReadOPF($mEPUB, $aOPF)
    If @error Then
        SetSliderLabel($mSlider, 'Error: ' & GetErrorMessage($EPUB_READOPF, @error))
        Return SetError(6, 0, False)
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
        If $tRECT.Right - $tRECT.Left < 900 Then $tRECT.Right = $tRECT.Left + 900
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
    GUICtrlSetPos($mSlider['Label'], 370, $iHeight - 55 + $mSlider['ElementHeigth'], $iWidth - $mSlider['ElementHeigth'] * 2 - 510, $mSlider['ElementHeigth'])
    GUICtrlSetPos($mSlider['Prev'], $iWidth - 120, $iHeight - 55, $mSlider['ElementHeigth'] * 2, $mSlider['ElementHeigth'] * 2)
    GUICtrlSetPos($mSlider['Next'], $iWidth - 60, $iHeight - 55, $mSlider['ElementHeigth'] * 2, $mSlider['ElementHeigth'] * 2)
    GUICtrlSetPos($mSlider['Config'], $iWidth - $mSlider['ElementHeigth'] - 130, $iHeight - 55 + $mSlider['ElementHeigth'], $mSlider['ElementHeigth'], $mSlider['ElementHeigth'])
    GUICtrlSetPos($mSlider['Jump'], $iWidth - $mSlider['ElementHeigth'] * 2 - 130, $iHeight - 55 + $mSlider['ElementHeigth'], $mSlider['ElementHeigth'], $mSlider['ElementHeigth'])
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

Func JumpToChapter()
    If Not IsMap($mEPUB) Then
        SetSliderLabel($mSlider, 'Error: ePUB is not loaded')
        Return SetError(1, 0, False)
    EndIf
    Local $hGUI = GUICreate('Jump to chapter', 240, 100 + GetTitleBarHeight(), Default, Default, $DS_SETFOREGROUND, Default, $hMain)
    Local $cCurrentChapter = GUICtrlCreateInput($mSlider['ActiveChapter'], 45, 20, 60, 25, 0x2001)
    Local $cDiv = GUICtrlCreateLabel('/', 105, 20, 30, 25, 0x201)
    Local $cMaxChapter = GUICtrlCreateInput($mSlider['Max'], 135, 20, 60, 25, 0x2801)
    Local $cJump = GUICtrlCreateButton('Jump', 10, 60, 105, 30)
    Local $cCancel = GUICtrlCreateButton('Cancel', 125, 60, 105, 30)
    GUICtrlSetFont($cCurrentChapter, $aFont[0], $aFont[1], 0, $aFont[2])
    GUICtrlSetFont($cDiv, $aFont[0], $aFont[1], 0, $aFont[2])
    GUICtrlSetFont($cMaxChapter, $aFont[0], $aFont[1], 0, $aFont[2])
    GUICtrlSetFont($cJump, $aFont[0], $aFont[1], 0, $aFont[2])
    GUICtrlSetFont($cCancel, $aFont[0], $aFont[1], 0, $aFont[2])
    GUICtrlSetBkColor($cCurrentChapter, $aColors[3])
    GUICtrlSetBkColor($cMaxChapter, $aColors[3])
    GUICtrlSetColor($cCurrentChapter, $aColors[4])
    GUICtrlSetColor($cMaxChapter, $aColors[4])
    GUICtrlSetColor($cDiv, $aColors[5])
    GUISetState(@SW_SHOW, $hGUI)
    While True
        Switch GUIGetMsg()
            Case $cJump
                Local $iChapter = Int(GUICtrlRead($cCurrentChapter))
                If $iChapter = $mSlider['ActiveChapter'] Then ExitLoop
                If $iChapter < 1 Or $iChapter > $mSlider['Max'] Then
                    SetSliderLabel($mSlider, 'Error: Invalid chapter range')
                Else
                    GUISetState(@SW_HIDE, $hGUI)
                    SetActiveChapter($mSlider, $iChapter)
                    LoadChapter($mEPUB, $mSlider['ActiveChapter'])
                    ExitLoop
                EndIf
            Case $cCancel
                ExitLoop
        EndSwitch
    WEnd
    SetSliderLabel($mSlider, 'Operation: Idle')
    WinActivate($hMain)
    GUIDelete($hGUI)
EndFunc

Func GetDefaultColor($iType)
    Switch $iType
        Case 0  ; Primary slider color
            Return 0x242423
        Case 1  ; Alternate slider color
            Return 0x333333
        Case 2  ; Selection slider color
            Return 0x9e2a2b
        Case 3  ; Inputs background color
            Return 0xFFFFFF
        Case 4  ; Inputs text color
            Return 0x000000
        Case 5  ; Labels text color
            Return 0x000000
        Case 6  ; Slider frame color
            Return 0x808080
    EndSwitch
EndFunc

Func GetDefaultFont($iType)
    Switch $iType
        Case 0  ; Font size
            Return 11
        Case 1  ; Font width
            Return 500
        Case 2  ; Font name
            Return 'Segoe UI'
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
        For $Index = 0 To UBound($aColors) - 1
            $aColors[$Index] = (MapExists($mSettings, $aColorProperty[$Index]) ? $mSettings[$aColorProperty[$Index]] : GetDefaultColor($Index))
        Next
        For $Index = 0 To UBound($aFont) - 1
            $aFont[$Index] = (MapExists($mSettings, $aFontProperty[$Index]) ? $mSettings[$aFontProperty[$Index]] : GetDefaultFont($Index))
        Next
        $fDeleteRes = (MapExists($mSettings, 'DeleteRes') ? ($mSettings['DeleteRes'] = 'True' ? True : False) : False)
        _SQLite_Close($hDB)
    Else
        $sCustomCSS = GetCustomCSS()
        For $Index = 0 To UBound($aColors) - 1
            $aColors[$Index] = GetDefaultColor($Index)
        Next
        For $Index = 0 To UBound($aFont) - 1
            $aFont[$Index] = GetDefaultFont($Index)
        Next
        $fDeleteRes = False
    EndIf
    If $fApply Then ApplySettings()
EndFunc

Func SaveSettings($sCSS, $fDeleteWS, $aColorSettings, $aFontSettings)
    Local $iColorProperties = UBound($aColorSettings)
    If UBound($aColorProperty) <> $iColorProperties Then
        SetSliderLabel($mSlider, 'Error: Invalid color properties')
        Return SetError(1, 0, False)
    EndIf
    Local $iFontProperties = UBound($aFontSettings)
    If UBound($aFontProperty) <> $iFontProperties Then
        SetSliderLabel($mSlider, 'Error: Invalid font properties')
        Return SetError(2, 0, False)
    EndIf
    Local $sColor, $sFont
    Local $sDelete = ($fDeleteWS ? 'True' : 'False')
    Local $sPrefix = 'INSERT INTO settings(property, value) VALUES'
    Local $sConflict = ' ON CONFLICT(property) DO UPDATE SET value = '
    Local $bCreateTable = (FileExists($__sWorkingDir & '\' & $__sInternalName & '\config.sqlite') ? False : True)
    Local $hDB = _SQLite_Open($__sWorkingDir & '\' & $__sInternalName & '\config.sqlite')
    If $bCreateTable Then _SQLite_Exec($hDB, 'CREATE TABLE settings(property VARCHAR(32) PRIMARY KEY, value TEXT);')
    _SQLite_Exec($hDB, $sPrefix & '("CSS", ' & _SQLite_FastEscape($sCSS) & ')' & $sConflict & _SQLite_FastEscape($sCSS) & ';')
    For $Index = 0 To $iColorProperties - 1
        $sColor = String('0x' & Hex($aColorSettings[$Index], 6))
        _SQLite_Exec($hDB, $sPrefix & '(' & _SQLite_FastEscape($aColorProperty[$Index]) & ', ' & _SQLite_FastEscape($sColor) & ')' & $sConflict & _SQLite_FastEscape($sColor) & ';')
    Next
    For $Index = 0 To $iFontProperties - 1
        $sFont = String($aFontSettings[$Index])
        _SQLite_Exec($hDB, $sPrefix & '(' & _SQLite_FastEscape($aFontProperty[$Index]) & ', ' & _SQLite_FastEscape($sFont) & ')' & $sConflict & _SQLite_FastEscape($sFont) & ';')
    Next
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
        GUICtrlSetFont($mSlider['Label'], $aFont[0], $aFont[1], 0, $aFont[2])
        GUICtrlSetColor($mSlider['Label'], $aColors[5])
        GUICtrlSetBkColor($mSlider['Frame'], $aColors[6])
        SetActiveChapter($mSlider, $mSlider['ActiveChapter'])
    EndIf
    If IsObj($oIE) Then InsertCSS($sCustomCSS)
EndFunc

Func Settings()
    ; Local variables
    Local $vTemp, $sCSS, $fDeleteWS, $iMsg, $iSelect, $cFontGroup
    Local $aLabel[14], $aButton[4], $aInput[5], $aColorPicker[7], $aCheckbox[2]
    Local $aColorSettings = $aColors
    Local $aOptSettings[2] = [$fDeleteRes, $fDeleteCfg]
    Local $aFontSettings = $aFont
    Local $sCredits = 'Epub icons created by shohanur.rahman13 - Flaticon (https://www.flaticon.com/free-icons/epub)' & @CRLF & @CRLF
    $sCredits &= 'Start icons created by hqrloveq - Flaticon (https://www.flaticon.com/free-icons/start)' & @CRLF & @CRLF
    $sCredits &= 'Home button icons created by hqrloveq - Flaticon (https://www.flaticon.com/free-icons/home-button)' & @CRLF & @CRLF
    $sCredits &= 'Ui icons created by Dewi Sari - Flaticon (https://www.flaticon.com/free-icons/ui)' & @CRLF & @CRLF
    $sCredits &= 'Refresh icons created by Freepik - Flaticon (https://www.flaticon.com/free-icons/refresh)' & @CRLF & @CRLF
    $sCredits &= 'Next icons created by KP Arts - Flaticon (https://www.flaticon.com/free-icons/next)' & @CRLF & @CRLF
    $sCredits &= 'Settings icons created by Freepik - Flaticon (https://www.flaticon.com/free-icons/settings)' & @CRLF & @CRLF
    $sCredits &= '<Ui icons created by Smashicons - Flaticon (https://www.flaticon.com/free-icons/ui)' & @CRLF
    ; UI
    Local $hGUI = GUICreate('Settings', 405, 360 + GetTitleBarHeight(), Default, Default, $DS_SETFOREGROUND, Default, $hMain)
    GUICtrlCreateTab(10, 10, 380, 300)
    GUICtrlCreateTabItem('CSS')
    $aInput[0] = GUICtrlCreateEdit($sCustomCSS, 20, 40, 360, 260, BitOR($WS_VSCROLL, 0x0004))  ; CSS Edit
    GUICtrlCreateTabItem('Theme')
    $aLabel[0] = GUICtrlCreateLabel('Slider primary color', 30, 40, 160, 25, 0x200)
    $aLabel[1] = GUICtrlCreateLabel('Slider alternate color', 30, 65, 160, 25, 0x200)
    $aLabel[2] = GUICtrlCreateLabel('Slider selection color', 30, 90, 160, 25, 0x200)
    $aLabel[3] = GUICtrlCreateLabel('Inputs background color', 30, 115, 160, 25, 0x200)
    $aLabel[4] = GUICtrlCreateLabel('Inputs text color', 30, 140, 160, 25, 0x200)
    $aLabel[5] = GUICtrlCreateLabel('Labels text color', 30, 165, 160, 25, 0x200)
    $aLabel[6] = GUICtrlCreateLabel('Frame color', 30, 190, 160, 25, 0x200)
    $aColorPicker[0] = GUICtrlCreateLabel('', 190, 42, 25, 20, 0x1000)  ; Primary color
    $aColorPicker[1] = GUICtrlCreateLabel('', 190, 67, 25, 20, 0x1000)  ; Alternate color
    $aColorPicker[2] = GUICtrlCreateLabel('', 190, 92, 25, 20, 0x1000)  ; Selection color
    $aColorPicker[3] = GUICtrlCreateLabel('', 190, 117, 25, 20, 0x1000) ; Inputs Background Color
    $aColorPicker[4] = GUICtrlCreateLabel('', 190, 142, 25, 20, 0x1000) ; Inputs Text Color
    $aColorPicker[5] = GUICtrlCreateLabel('', 190, 167, 25, 20, 0x1000) ; Labels Text Color
    $aColorPicker[6] = GUICtrlCreateLabel('', 190, 192, 25, 20, 0x1000) ; Frame Color
    $cFontGroup = GUICtrlCreateGroup(' Font ', 20, 220, 360, 80)
    $aLabel[7] = GUICtrlCreateLabel('Font size', 30, 235, 90, 25, 0x200)
    $aLabel[8] = GUICtrlCreateLabel('Font width', 210, 235, 90, 25, 0x200)
    $aLabel[9] = GUICtrlCreateLabel('Font name', 30, 265, 90, 25, 0x200)
    $aInput[1] = GUICtrlCreateInput($aFontSettings[0], 120, 235, 60, 25, 0x0801)    ; Font size
    $aInput[2] = GUICtrlCreateInput($aFontSettings[1], 300, 235, 60, 25, 0x0801)    ; Font width
    $aInput[3] = GUICtrlCreateInput($aFontSettings[2], 120, 265, 170, 25, 0x0800)   ; Font name
    $aButton[0] = GUICtrlCreateButton('..', 300, 265, 60, 25)                       ; Select font
    GUICtrlCreateGroup('', -99, -99, 1, 1)
    GUICtrlCreateTabItem('Workspace')
    $aLabel[10] = GUICtrlCreateLabel('Current workspace size', 30, 40, 150, 25, 0x200)
    $aLabel[11] = GUICtrlCreateLabel(GetWorkspaceSize(), 190, 40, 100, 25, 0x200)
    $aCheckbox[0] = GUICtrlCreateCheckbox('Clean up workspace', 30, 65, 160, 25)                    ; Delete resources
    $aCheckbox[1] = GUICtrlCreateCheckbox('Delete application settings on exit', 30, 90, 250, 25)   ; Delete config
    GUICtrlCreateTabItem('Info')
    $aLabel[12] = GUICtrlCreateLabel('Application version: ' & $sVersion, 30, 40, 300, 25, 0x200)
    $aLabel[13] = GUICtrlCreateLabel('Author: Andreik (AutoIt Forum)', 30, 65, 300, 25, 0x200)
    $aInput[4] = GUICtrlCreateEdit('', 30, 90, 340, 200, BitOR($WS_VSCROLL, 0x0800))                ; Credits
    GUICtrlCreateTabItem('')
    $aButton[1] = GUICtrlCreateButton('Save', 10, 320, 120, 30)      ; Save
    $aButton[2] = GUICtrlCreateButton('Reset', 140, 320, 120, 30)    ; Reset
    $aButton[3] = GUICtrlCreateButton('Cancel', 270, 320, 120, 30)   ; Cancel
    ; Set styles
    For $Index = 0 To UBound($aLabel) - 1
        GUICtrlSetFont($aLabel[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
        GUICtrlSetColor($aLabel[$Index], $aColors[5])
    Next
    For $Index = 0 To UBound($aButton) - 1
        GUICtrlSetFont($aButton[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
        If $Index = 0 Then
            GUICtrlSetCursor($aButton[$Index], 0)
            GUICtrlSetTip($aButton[$Index], 'Select font')
        EndIf
    Next
    For $Index = 0 To UBound($aInput) - 1
        GUICtrlSetFont($aInput[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
        GUICtrlSetBkColor($aInput[$Index], $aColors[3])
        GUICtrlSetColor($aInput[$Index], $aColors[4])
        If $Index = 4 Then GUICtrlSetData($aInput[$Index], $sCredits)
    Next
    For $Index = 0 To UBound($aColorPicker) - 1
        GUICtrlSetBkColor($aColorPicker[$Index], $aColors[$Index])
        GUICtrlSetCursor($aColorPicker[$Index], 0)
    Next
    For $Index = 0 To UBound($aCheckbox) - 1
        GUICtrlSetFont($aCheckbox[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
        GUICtrlSetState($aCheckbox[$Index], ($aOptSettings[$Index] ? $GUI_CHECKED : $GUI_UNCHECKED))
    Next
    GUICtrlSetFont($cFontGroup, $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
    GUISetState(@SW_SHOW, $hGUI)

    While True
        $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $aColorPicker[0] To $aColorPicker[6]   ; Color pickers
                $iSelect = $iMsg - $aColorPicker[0]
                $vTemp = _ChooseColor(2, $aColorSettings[$iSelect], 2, $hGUI)
                If Not @error Then
                    $aColorSettings[$iSelect] = $vTemp
                    GUICtrlSetBkColor($aColorPicker[$iSelect], $aColorSettings[$iSelect])
                EndIf
            Case $aButton[0]                            ; Font select
                $vTemp = _ChooseFont($aFontSettings[2], $aFontSettings[0], Default, $aFontSettings[1])
                If Not @error Then
                    $aFontSettings[0] = $vTemp[3]
                    $aFontSettings[1] = $vTemp[4]
                    $aFontSettings[2] = $vTemp[2]
                    GUICtrlSetData($aInput[1], $aFontSettings[0])
                    GUICtrlSetData($aInput[2], $aFontSettings[1])
                    GUICtrlSetData($aInput[3], $aFontSettings[2])
                EndIf
            Case $aButton[1]                            ; Save button
                $sCSS = GUICtrlRead($aInput[0])
                $fDeleteWS = (GUICtrlRead($aCheckbox[0]) = $GUI_CHECKED ? True : False)
                $fDeleteCfg = (GUICtrlRead($aCheckbox[1]) = $GUI_CHECKED ? True : False)
                SaveSettings($sCSS, $fDeleteWS, $aColorSettings, $aFontSettings)
                LoadSettings(True)
                ExitLoop
            Case $aButton[2]                            ; Reset button
                For $Index = 0 To UBound($aColorSettings) - 1
                    $aColorSettings[$Index] = GetDefaultColor($Index)
                Next
                For $Index = 0 To UBound($aFontSettings) - 1
                    $aFontSettings[$Index] = GetDefaultFont($Index)
                Next
                $sCSS = GetCustomCSS()
                For $Index = 0 To UBound($aLabel) - 1
                    GUICtrlSetFont($aLabel[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
                    GUICtrlSetColor($aLabel[$Index], $aColorSettings[5])
                Next
                For $Index = 0 To UBound($aButton) - 1
                    GUICtrlSetFont($aButton[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
                Next
                For $Index = 0 To UBound($aInput) - 1
                    GUICtrlSetFont($aInput[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
                    GUICtrlSetBkColor($aInput[$Index], $aColorSettings[3])
                    GUICtrlSetColor($aInput[$Index], $aColorSettings[4])
                    If $Index = 0 Then GUICtrlSetData($aInput[$Index], $sCSS)
                Next
                For $Index = 0 To UBound($aColorPicker) - 1
                    GUICtrlSetBkColor($aColorPicker[$Index], $aColorSettings[$Index])
                    GUICtrlSetCursor($aColorPicker[$Index], 0)
                Next
                For $Index = 0 To UBound($aCheckbox) - 1
                    GUICtrlSetFont($aCheckbox[$Index], $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
                    GUICtrlSetState($aCheckbox[$Index], $GUI_UNCHECKED)
                Next
                GUICtrlSetFont($cFontGroup, $aFontSettings[0], $aFontSettings[1], 0, $aFontSettings[2])
            Case $aButton[3]                            ; Cancel button
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
