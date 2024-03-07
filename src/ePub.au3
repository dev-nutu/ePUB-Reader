#include-once

Func ePUB_Init($sPath)
    ; Set working directory
    __SetWorkingDir($__sWorkingDir)

    ; Check if the file exists
    If Not FileExists($sPath) Then Return SetError(1, 0, False)

    ; Check if the file is an ePUB
    Local $aPath = StringRegExp($sPath, '(?i)^(?:[a-zA-Z]:.+\\)(.*?)(\.epub)$', 3)
    If @error Then Return SetError(2, @error, False)

    ; Create an ePUB map
    Local $aCollector[1]
    Local $mEPUB[]
    $mEPUB['Chapters'] = Null
    $mEPUB['NumOfChapters'] = 0

    ; Create Shell Application object
    $mEPUB['App'] = ObjCreate('Shell.Application')
    If Not IsObj($mEPUB['App']) Then Return SetError(3, 0, False)

    ; Create a memory database for ePUB file structure
    $mEPUB['DB'] = _SQLite_Open()
    If @error Then Return SetError(4, 1, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE TABLE config (cfg VARCHAR(64) NOT NULL, value VARCHAR(512) NOT NULL);')
    If @error Then Return SetError(4, 2, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE TABLE manifest (href VARCHAR(512) NOT NULL, id VARCHAR(512) NOT NULL);')
    If @error Then Return SetError(4, 3, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE TABLE spine (sequence INTEGER PRIMARY KEY AUTOINCREMENT, idref VARCHAR(512) NOT NULL);')
    If @error Then Return SetError(4, 4, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE TABLE guides (href VARCHAR(512) NOT NULL, title VARCHAR(512) NOT NULL);')
    If @error Then Return SetError(4, 5, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE TABLE metadata (meta VARCHAR(512) NOT NULL, value TEXT NOT NULL);')
    If @error Then Return SetError(4, 6, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE INDEX IDX_Manifest ON manifest (href,id);')
    If @error Then Return SetError(4, 7, False)
    _SQLite_Exec($mEPUB['DB'], 'CREATE INDEX IDX_Guides ON guides (href ASC);')
    If @error Then Return SetError(4, 8, False)

    ; Create unique temporary directory
    $mEPUB['Dir'] = __GetUniqueDir()
    If Not DirCreate($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir']) Then
        ePub_Release($mEPUB, $aCollector)
        Return SetError(5, @error, False)
    EndIf

    ; Move ePUB file to directory and change extension to .zip
    $mEPUB['File'] = $aPath[0] & '.zip'
    If Not FileCopy($sPath, $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File']) Then
        ePub_Release($mEPUB, $aCollector)
        Return SetError(6, @error, False)
    EndIf

    ; Return a map with the details of the file
    Return $mEPUB
EndFunc

Func ePUB_GetFiles($mEPUB)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    If Not FileExists($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File']) Then Return SetError(2, 0, False)
    Local $oApp = $mEPUB['App']
    Local $oItems = $oApp.NameSpace($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File']).Items
    Local $iCount = $oItems.Count
    If $iCount = 0 Then Return SetError(3, 0, False)
    Local $aFile[$iCount + 1][2]
    $aFile[0][0] = $iCount
    Local $oItem
    For $Index = 0 To $iCount - 1
        $oItem = $oItems.Item($Index)
        If IsObj($oItem) Then
            $aFile[$Index + 1][0] = $oItem.Name
            $aFile[$Index + 1][1] = $oItem.IsFolder
        EndIf
    Next
    Return $aFile
EndFunc

Func ePUB_Extract($mEPUB, $aFile, $bShowProgress = False, $bDeleteZip = True)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    If Not IsArray($aFile) Then Return SetError(2, 0, False)
    If Not FileExists($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File']) Then Return SetError(3, 0, False)
    If $aFile[0][0] = 0 Then Return SetError(4, 0, False)
    Local $oApp = $mEPUB['App']
    Local $oItem
    For $Index = 1 To $aFile[0][0]
        $oItem = $oApp.NameSpace($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File']).ParseName($aFile[$Index][0])
        $oApp.NameSpace($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir']).Copyhere($oItem, 16 + ($bShowProgress ? 0 : 4))
    Next
    If $bDeleteZip Then FileDelete($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $mEPUB['File'])
    Return SetError(0, 0, True)
EndFunc

Func ePUB_Validate($mEPUB)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    Local $sMimetypePath = $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\mimetype'
    If Not FileExists($sMimetypePath) Then Return SetError(2, 0, False)
    If FileRead($sMimetypePath) <> 'application/epub+zip' Then  Return SetError(3, 0, False)
    Return SetError(0, 0, True)
EndFunc

Func ePUB_GetOPF($mEPUB)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    Local $sContainerPath = $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\META-INF\container.xml'
    If Not FileExists($sContainerPath) Then Return SetError(2, 0, False)
    Local $sContainerData = FileRead($sContainerPath)
    Local $aContainer = StringRegExp($sContainerData, '(?is)<rootfile(?:.*?)full-path="(?:(.*?)\/){0,1}(.*?)"(?:.*?)media-type="(.*?)"(?:.*?)\/>', 3)
    If @error Then Return SetError(3, @error, False)
    _SQLite_Exec($mEPUB['DB'], 'INSERT INTO config (cfg, value) VALUES ("path", ' & _SQLite_FastEscape($aContainer[0] ? $aContainer[0]  : '') & ')')
    If @error Then Return SetError(4, @error, False)
    Return $aContainer
EndFunc

Func ePUB_ReadOPF($mEPUB, $aOPF)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    If Not IsArray($aOPF) Then Return SetError(2, 0, False)
    Local $iElements = UBound($aOPF)
    If $iElements < 2 Then Return SetError(3, 0, False)
    Local $sOPF = FileRead($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'] & '\' & $aOPF[0] & ($iElements = 3 ? '\' & $aOPF[1] : ''))
    Local $aVersion = StringRegExp($sOPF, '(?is)<package(?:.*?)version="(.*?)"(?:.*?)>', 3)
    If @error Then Return SetError(4, @error, False)
    If $aVersion[0] <> '2.0' And $aVersion[0] <> '3.0' Then Return SetError(5, 0, False)
    Local $aMetadata= StringRegExp($sOPF, '(?is)<dc:(?:.*?)>(.*?)<\/dc:(.*?)>', 3)
    If @error Then Return SetError(6, @error, False)
    For $Index = 0 To UBound($aMetadata) - 1 Step 2
        _SQLite_Exec($mEPUB['DB'], 'INSERT INTO metadata (meta, value) VALUES(' & _SQLite_FastEscape($aMetadata[$Index + 1]) & ',' & _SQLite_FastEscape($aMetadata[$Index]) & ');')
    Next
    Local $aItemRefs = StringRegExp($sOPF, '(?is)<itemref(?:.*?)idref="(.*?)"(?:.*?)\/>', 3)
    If @error Then Return SetError(7, @error, False)
    For $Index = 0 To UBound($aItemRefs) - 1
        _SQLite_Exec($mEPUB['DB'], 'INSERT INTO spine (idref) VALUES(' & _SQLite_FastEscape($aItemRefs[$Index]) & ');')
    Next
    Local $aID, $aHref, $aTitle
    Local $aItems = StringRegExp($sOPF, '(?is)<item (.*?)\/>', 3)
    If @error Then Return SetError(8, @error, False)
    For $Index = 0 To UBound($aItems) - 1
        $aID = StringRegExp($aItems[$Index], '(?is)id="(.*?)"', 3)
        If @error Then ContinueLoop
        $aHref = StringRegExp($aItems[$Index], '(?is)href="(.*?)"', 3)
        If @error Then ContinueLoop
        _SQLite_Exec($mEPUB['DB'], 'INSERT INTO manifest (href, id) VALUES(' & _SQLite_FastEscape($aHref[0]) & ',' & _SQLite_FastEscape($aID[0]) & ');')
    Next
    ; Guides are optional
    Local $aGuide = StringRegExp($sOPF, '<reference (.*?)\/>', 3)
    If Not @error Then
        For $Index = 0 To UBound($aGuide) - 1
            $aHref = StringRegExp($aGuide[$Index], '(?is)href="(.*?)"', 3)
            If @error Then ContinueLoop
            $aTitle = StringRegExp($aGuide[$Index], '(?is)title="(.*?)"', 3)
            If @error Then ContinueLoop
            _SQLite_Exec($mEPUB['DB'], 'INSERT INTO guides (href, title) VALUES(' & _SQLite_FastEscape($aHref[0]) & ',' & _SQLite_FastEscape($aTitle[0]) & ');')
        Next
    EndIf
    Local $aQuery = SQLite_Query($mEPUB['DB'], 'SELECT s.sequence, s.idref, m.href, g.title FROM spine s INNER JOIN manifest m ON s.idref=m.id LEFT JOIN guides g ON m.href=g.href ORDER BY s.sequence;')
    Local $iChapters = @extended
    Return SetError(0, $iChapters, $aQuery)
EndFunc

Func ePUB_Release($mEPUB, ByRef $aCollector, $fSave = False)
    If Not IsMap($mEPUB) Then Return SetError(1, 0, False)
    If MapExists($mEPUB, 'Dir') Then
        If $fSave Then
            $aCollector[0] += 1
            ReDim $aCollector[$aCollector[0] + 1]
            $aCollector[$aCollector[0]] = $__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir']
        Else
            DirRemove($__sWorkingDir & '\' & $__sInternalName & '\' & $mEPUB['Dir'], 1)
            If IsArray($aCollector) Then
                For $Index = 1 To $aCollector[0]
                    DirRemove($aCollector[$Index], 1)
                Next
            EndIf
        EndIf
    EndIf
    If MapExists($mEPUB, 'DB') Then _SQLite_Close($mEPUB['DB'])
    Return SetError(0, 0, True)
EndFunc

Func SQLite_Query($hDB, $sQuery)
    Local $aResult, $iRows, $iColumns
    _SQLite_GetTable2d($hDB, $sQuery, $aResult, $iRows, $iColumns)
    If @error Then
        Return SetError(1, 0, False)
    Else
        Return SetError(0, UBound($aResult, 1) - 1, $aResult)
    EndIf
EndFunc

Func __SetWorkingDir($sPath)
    If Not FileExists($sPath & '\' & $__sInternalName) Then DirCreate($sPath & '\' & $__sInternalName)
EndFunc

Func __GetUniqueDir($iLen = 8)
    Local $sDirName
    Local Static $aChars = StringSplit('abcdefghijklmnopqrstuvwxyz0123456789', '')
    For $Index = 1 To $iLen
        $sDirName &= $aChars[Random(1, $aChars[0], 1)]
    Next
    Return FileExists(@TempDir & '\' & $__sInternalName & '\' & $sDirName) ? __GetUniqueDir() : $sDirName
EndFunc
