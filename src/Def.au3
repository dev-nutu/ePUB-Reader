#include-once

Global Enum $EPUB_INIT, $EPUB_GETFILES, $EPUB_EXTRACT, $EPUB_VALIDATE, $EPUB_GETOPF, $EPUB_READOPF
Global $hMain, $cReader, $oIE, $oIEEvents, $cTitle, $aCtrlButtons[6], $mSlider, $mEPUB, $iMsg
Global $__sInternalName = 'ePUB_AU3'
Global $__sWorkingDir = @TempDir
Global $fDeleteRes, $fDeleteCfg = False
Global $fAllowForward = False, $fAllowBack = False
Global $aTips[6] = ['Load ePUB', 'Back', 'First chapter', 'Stop', 'Next', 'Refresh']
Global $sCustomCSS, $aColors[3], $aCollector[1]
Global Const $sAppTitle = 'ePub Reader'
Global Const $sVersion = '1.0'