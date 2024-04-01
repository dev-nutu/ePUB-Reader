#include-once

Global Enum $EPUB_INIT, $EPUB_GETFILES, $EPUB_EXTRACT, $EPUB_VALIDATE, $EPUB_GETOPF, $EPUB_READOPF
Global $hMain, $cReader, $oIE, $oIEEvents, $cTitle, $aCtrlButtons[6], $mSlider, $mEPUB, $iMsg
Global $__sInternalName = 'ePUB_AU3'
Global $__sWorkingDir = @TempDir
Global $bDeleteRes, $bDeleteCfg = False
Global $bAllowForward = False, $bAllowBack = False
Global $aTips[6] = ['Load ePUB', 'Back', 'First chapter', 'Stop', 'Next', 'Refresh']
Global $aColorProperty[7] = ['PrimaryColor', 'AlternateColor', 'SelectionColor', 'InputsBkColor', 'InputsTextColor', 'LabelsTextColor', 'FrameColor']
Global $aFontProperty[3] = ['FontSize', 'FontWeight', 'FontName']
Global $sCustomCSS, $aColors[7], $aFont[3], $aCollector[1]
Global Const $sAppTitle = 'ePub Reader'
Global Const $sVersion = '1.2.2.0'
