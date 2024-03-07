#include-once

Global Enum $NO_SLIDER_EVENT, $CHAPTER_CHANGED, $PREV_CHAPTER, $NEXT_CHAPTER

Func CreateSlider($hGUI, $iX, $iY, $iWidth, $iHeight, $iSliderColor = 0x242423, $iSliderAltColor = 0x333333,$iActiveColor = 0x9e2a2b)
    Local $mSlider[]
    $mSlider['GUI'] = $hGUI
    $mSlider['Color'] = $iSliderColor
    $mSlider['AltColor'] = $iSliderAltColor
    $mSlider['ActiveColor'] = $iActiveColor
    $mSlider['Min'] = Null
    $mSlider['Max'] = Null
    $mSlider['X'] = $iX
    $mSlider['Y'] = $iY
    $mSlider['ActiveChapter'] = 0
    $mSlider['ElementHeigth'] = Int($iHeight / 2)
    $mSlider['W'] = $iWidth - ($iHeight * 2) - 20
    $mSlider['Frame'] = GUICtrlCreateLabel('', $iX, $iY, $mSlider['W'], $mSlider['ElementHeigth'], 0x1000)
    $mSlider['Label'] = GUICtrlCreateLabel('', $iX, $iY + $mSlider['ElementHeigth'], $mSlider['W'] - $mSlider['ElementHeigth'] - 10, $mSlider['ElementHeigth'], 0x200)
    $mSlider['Prev'] = GUICtrlCreateButton('', $iX + $mSlider['W'] + 10, $iY, $iHeight, $iHeight, 0x0380)
    $mSlider['Next'] = GUICtrlCreateButton('', $iX + $iWidth - $iHeight, $iY, $iHeight, $iHeight, 0x0380)
    $mSlider['Config'] = GUICtrlCreateButton('', $iX + $mSlider['W'] - $mSlider['ElementHeigth'], $iY + $mSlider['ElementHeigth'], $mSlider['ElementHeigth'], $mSlider['ElementHeigth'], 0x0380)
    GUICtrlSetTip($mSlider['Prev'], 'Previous chapter')
    GUICtrlSetTip($mSlider['Next'], 'Next chapter')
    GUICtrlSetTip($mSlider['Config'], 'Settings')
    GUICtrlSetCursor($mSlider['Prev'], 0)
    GUICtrlSetCursor($mSlider['Next'], 0)
    GUICtrlSetCursor($mSlider['Config'], 0)
    GUICtrlSetFont($mSlider['Label'], 11, 500, 0, 'Segoe UI')
    GUICtrlSetBkColor($mSlider['Frame'], $iSliderColor)
    GUICtrlSetState($mSlider['Frame'], 128)
    Return $mSlider
EndFunc

Func DeleteSlider(ByRef $mSlider)
    If Not IsMap($mSlider) Then Return SetError(1, 0, False)
    GUICtrlDelete($mSlider['Frame'])
    GUICtrlDelete($mSlider['Label'])
    GUICtrlDelete($mSlider['Prev'])
    GUICtrlDelete($mSlider['Next'])
    If $mSlider['Min'] <> Null And $mSlider['Max'] <> Null Then
        For $Index = $mSlider['Min'] To $mSlider['Max']
            If MapExists($mSlider, $Index) Then
                GUICtrlDelete($mSlider[$Index])
            EndIf
        Next
    EndIf
    $mSlider = Null
    Return SetError(0, 0, True)
EndFunc

Func ResetSlider(ByRef $mSlider)
    If Not IsMap($mSlider) Then Return SetError(1, 0, False)
    If $mSlider['Min'] <> Null And $mSlider['Max'] <> Null Then
        For $Index = $mSlider['Min'] To $mSlider['Max']
            If MapExists($mSlider, $Index) Then
                GUICtrlDelete($mSlider[$Index])
            EndIf
        Next
    EndIf
    $mSlider['Min'] = Null
    $mSlider['Max'] = Null
    $mSlider['ActiveChapter'] = 0
    Return SetError(0, 0, True)
EndFunc

Func SliderEvent(ByRef $mSlider, $iMsg)
    If Not IsMap($mSlider) Then Return SetError(1, 0, False)
    If $mSlider['Min'] = Null Or $mSlider['Max'] = Null Then Return SetError(2, 0, False)
    Local $iFirst = $mSlider[$mSlider['Min']]
    Local $iLast = $mSlider[$mSlider['Max']]
    Switch $iMsg
        Case $iFirst To $iLast
            Local $iSelect = $iMsg - $iFirst + 1
            SetActiveChapter($mSlider, $iSelect)
            Return $CHAPTER_CHANGED
        Case $mSlider['Prev']
            SetActiveChapter($mSlider, $mSlider['ActiveChapter'] - 1)
            If Not @error Then Return $PREV_CHAPTER
        Case $mSlider['Next']
            SetActiveChapter($mSlider, $mSlider['ActiveChapter'] + 1)
            If Not @error Then Return $NEXT_CHAPTER
    EndSwitch
    Return SetError(0, 0, $NO_SLIDER_EVENT)
EndFunc

Func SetSliderLimits(ByRef $mSlider, $iMin = 0, $iMax = 0)
    If Not IsMap($mSlider) Then Return SetError(1, 0, False)
    If $iMin > $iMax Then Return SetError(2, 0, False)
    If $mSlider['Min'] <> Null And $mSlider['Max'] <> Null Then
        For $Index = $mSlider['Min'] To $mSlider['Max']
            If MapExists($mSlider, $Index) Then
                GUICtrlDelete($mSlider[$Index])
                MapRemove($mSlider, $Index)
            EndIf
        Next
    EndIf
    Local $iLocalMin = $iMin
    Local $iLocalMax = $iMax
    $iLocalMax -= ($iLocalMin <> 1 ? $iLocalMin - 1 : 0)
    $iLocalMin = 1
    Local $iSegLen = Int($mSlider['W'] / $iLocalMax)
    Local $iLeft = Mod($mSlider['W'], $iLocalMax)
    Local $iPos, $iWidth, $iDelta = $iMin - $iLocalMin, $iCurrentChapter
    For $Index = $iLocalMin To $iLocalMax
        $iPos = $mSlider['X'] + (($Index - 1) * $iSegLen) + ($iLocalMax - $Index < $iLeft ? $iLeft - ($iLocalMax - $Index) - 1 : 0)
        $iWidth = $iSegLen + ($iLocalMax - $Index < $iLeft ? 1 : 0)
        $iCurrentChapter = $Index + $iDelta
        $mSlider[$iCurrentChapter] = GUICtrlCreateLabel('', $iPos, $mSlider['Y'], $iWidth, $mSlider['ElementHeigth'])
        GUICtrlSetBkColor($mSlider[$iCurrentChapter], Mod($Index, 2) ? $mSlider['Color'] : $mSlider['AltColor'])
        GUICtrlSetCursor($mSlider[$iCurrentChapter], 0)
    Next
    $mSlider['Min'] = $iMin
    $mSlider['Max'] = $iMax
    Return SetError(0, 0, True)
EndFunc

Func SetActiveChapter(ByRef $mSlider, $iChapter)
    If Not IsMap($mSlider) Then Return SetError(1, 0, False)
    If $mSlider['Min'] = Null Or $mSlider['Max'] = Null Then Return SetError(2, 0, False)
    If $iChapter < $mSlider['Min'] Or $iChapter > $mSlider['Max'] Then Return SetError(3, 0, False)
    If $mSlider['ActiveChapter'] Then
        Local $iColor = Mod($mSlider['ActiveChapter'], 2) ? $mSlider['Color'] : $mSlider['AltColor']
        GUICtrlSetBkColor($mSlider[$mSlider['ActiveChapter']], $iColor)
    EndIf
    GUICtrlSetBkColor($mSlider[$iChapter], $mSlider['ActiveColor'])
    $mSlider['ActiveChapter'] = $iChapter
    GUICtrlSetState($mSlider[$mSlider['ActiveChapter']], 256)
    SetError(0, 0, True)
EndFunc

Func SetSliderLabel($mSlider, $sText)
    GUICtrlSetData($mSlider['Label'], $sText)
EndFunc