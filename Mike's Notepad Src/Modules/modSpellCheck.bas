Attribute VB_Name = "modSpellCheck"
Option Explicit

Dim dbWordList As DAO.Database
Dim cPhoneme As New clsPhoneme

' used to interact with the form frmSpelling
Public sTextBeingChecked As String
Public sOriginalText As String
Public bReplaceText As Boolean

' used to hold words being ignored
Public sTempWords() As String

' used to process text
Dim sWords() As String
Dim iIdx As Integer

' If a word is replaced in text we need to change it in the list of words being spellchecked
' if the replace all option is used
Public Sub ChangeReplaced(ByVal sOriginalStr As String, ByVal sReplaceStr As String)
Dim i As Integer

    For i = iIdx To UBound(sWords)
        If sWords(i) = sOriginalStr Then
            sWords(i) = sReplaceStr
        End If
    Next i

End Sub

' load a list of words into the word database
Public Function LoadWords(ByVal sFileName As String)
Dim f As Integer
Dim tmpStr As String
Dim i As Long
Dim wordRS As DAO.Recordset
Dim checkWordRS As DAO.Recordset
        
    f = FreeFile()
    Set wordRS = dbWordList.OpenRecordset("Words", dbOpenDynaset)
    With wordRS
        Open sFileName For Input As #f
            ' Read words from the file and add to the database
            ' if they are not there
            Do While Not EOF(f)
                ' Prevent the UI from freezing up
                DoEvents
                Line Input #f, tmpStr
                tmpStr = Trim$(tmpStr)
                
                Set checkWordRS = dbWordList.OpenRecordset("SELECT * from Words WHERE " & _
                                                            "[Word] = " & Chr$(34) & Trim(tmpStr) & Chr$(34), _
                                                            dbOpenSnapshot)
                If checkWordRS.RecordCount = 0 Then
                    .AddNew
                    !Word = tmpStr
                    !Soundex = cPhoneme.GetSoundexWord(tmpStr)
                    .Update
                    i = i + 1
                End If
                checkWordRS.Close
                If i Mod 1000 = 0 Then
                    frmLoadWords.sbStatus.Panels(0).Text = "Added..." & i & " words."
                    DoEvents
                End If
            Loop
        Close #f
        frmLoadWords.sbStatus.Panels(0).Text = "Database updated! " & i & " words were added."
        .Close
    End With

End Function

' add a single word to the database of words
Public Function AddWord(ByVal sWord As String, Optional ShowMsg As Boolean = False)
Dim tmpStr As String
Dim i As Long
Dim wordRS As DAO.Recordset
        
    Set wordRS = dbWordList.OpenRecordset("SELECT * from Words WHERE " & _
                                                           "[Word] = " & Chr$(34) & Trim(sWord) & Chr$(34), _
                                                           dbOpenDynaset)
    With wordRS
        If .RecordCount = 0 Then
            .AddNew
            !Word = sWord
            !Soundex = cPhoneme.GetSoundexWord(sWord)
            .Update
            If ShowMsg Then MsgBox sWord & " has been added to word list."
        End If
        .Close
    End With

End Function

' spell check the passed string
Public Function SpellCheck(ByVal sWordsToCheck As String) As String
' sOriginaltext holds an unchanged copy of the text
' sWordsToCheck is used to extract the words for spellcheck
' sTextBeingChecked is used to hold the changed text

Dim sWorkStr As String
Dim rsFindWord As DAO.Recordset

    Set dbWordList = OpenDatabase(App.Path & "\Spell Checker\Words.mdb", False, False)
    
    ' make sure there are some words to check against
    Set rsFindWord = dbWordList.OpenRecordset("SELECT * from Words")
    If rsFindWord.RecordCount = 0 Then
        If MsgBox("Dictionary has no words." & vbCrLf & "Do you want to continue?", vbQuestion & vbYesNo) = vbNo Then
            SpellCheck = sWordsToCheck
            Exit Function
        End If
    End If
    rsFindWord.Close

    ' copy the original text
    sTextBeingChecked = sWordsToCheck
    sOriginalText = sWordsToCheck

    ' default text replacement to true
    bReplaceText = True

    ' remove line separators
    sWordsToCheck = Replace(sWordsToCheck, Chr(10), " ")
    sWordsToCheck = Replace(sWordsToCheck, Chr(13), "")

    ' put the words into an array
    sWords = Split(sWordsToCheck, " ")

    ' initialize the array of words to ignore
    ReDim sTempWords(0)

    For iIdx = 0 To UBound(sWords)
        sWorkStr = RemovePunctuation(Trim$(sWords(iIdx)))
        If Len(sWorkStr) > 0 Then
            If Not IsNumeric(sWorkStr) Then
                ' See if the word exists in the database
                Set rsFindWord = dbWordList.OpenRecordset("SELECT [Word] from Words WHERE " & _
                                                            "[Word] = " & Chr$(34) & Trim(sWorkStr) & Chr$(34), _
                                                            dbOpenSnapshot)
                If rsFindWord.RecordCount = 0 Then
                    ' check that this word is spelt correctly
                    CheckThisWord sWorkStr
                End If
                rsFindWord.Close
            End If
        End If
        If bReplaceText = False Then
            SpellCheck = sOriginalText
            Exit Function
        End If
    Next iIdx

    ReDim sWords(0)
    ReDim sTempWords(0)

    If bReplaceText Then
        SpellCheck = sTextBeingChecked
    Else
        SpellCheck = sOriginalText
    End If

End Function

Private Function DataIsPresentIn(ByRef TestRS As DAO.Recordset) As Boolean
On Error Resume Next
    DataIsPresentIn = True
    TestRS.MoveFirst
    If Err Then DataIsPresentIn = False
    Err.Clear
End Function

Private Function CheckThisWord(ByVal sWord As String)
' Loads the list in the frmSpelling form with the matches and shows the form
' The user can then select an action to perform
Dim sMatch As String, sSoundex As String
Dim SndxMatchRS As DAO.Recordset
Dim lNavIndex As Long, lNavMax As Long, lenTmp As Long, iLevDist As Long
Dim iThreshold As Integer, i As Integer

    iThreshold = 3
    
    If sWord <> vbNullString Then
        For i = 0 To UBound(sTempWords)
            If UCase(sTempWords(i)) = UCase(sWord) Then
                Exit Function
            End If
        Next i
        
        sSoundex = cPhoneme.GetSoundexWord(sWord)
        '// Now find all entries in the database which match the soundex of the input word
        Set SndxMatchRS = dbWordList.OpenRecordset("SELECT [word] from Words WHERE " & _
                                                   "Soundex = " & Chr$(34) & sSoundex & Chr$(34), _
                                                   dbOpenSnapshot)
        '// Populate the Listbox
        frmMain.MousePointer = vbHourglass
        If DataIsPresentIn(SndxMatchRS) Then
            Load frmSpelling
            frmSpelling.txtWord.Text = sWord
            With SndxMatchRS
                .MoveLast
                lNavMax = .RecordCount
                .MoveFirst
                For lNavIndex = 1 To lNavMax
                    DoEvents
                    
                    sMatch = Trim$(!Word)
                    Debug.Print sMatch
                    If sMatch <> vbNullString Then
                        DoEvents
                        If Len(sMatch) < iThreshold Then
                            iThreshold = Len(sMatch) + 1
                        Else
                            iThreshold = 4
                        End If
                        
                        iLevDist = cPhoneme.GetLevenshteinDistance(sWord, UCase$(sMatch))
                        
                        ' Get all Levenshtein distances less than the threshold value
                        ' Add them to the list in Lev. Distance order
                        If iLevDist <= iThreshold Then
                            If iLevDist < frmSpelling.lstMatches.ListCount Then
                                frmSpelling.lstMatches.AddItem sMatch, iLevDist
                            Else
                                frmSpelling.lstMatches.AddItem sMatch
                            End If
                        End If
                    End If ' strMatch <> vbNullString
                    .MoveNext
                Next lNavIndex
            End With ' SndxMatchRS
            frmMain.MousePointer = vbDefault
            frmSpelling.Show vbModal, frmMain
            If bReplaceText = False Then
                Exit Function
            End If
        End If ' DataIsPresentIn(SndxMatchRS)
    End If ' sworkstr <> nullstring
    frmMain.MousePointer = vbDefault
End Function

Private Function RemovePunctuation(ByVal sInStr As String) As String
Dim sPunctuation() As String
Dim sOutStr As String
Dim i As Integer

    ' don't remove single quotes ( ' ) as they are part of the plural form of many words
    ' - are used for hyphenated words
    sPunctuation = Split("~ ` ! @ # $ % ^ & * ( ) _ + = { } [ ] : "" ; < > ? , . / | \", " ")
    sOutStr = sInStr
    
    For i = 0 To UBound(sPunctuation)
        sOutStr = Replace(sOutStr, sPunctuation(i), "")
    Next i
    
    ' replace everything that's not a printable character??????
    sOutStr = Replace(sOutStr, vbTab, " ")
    sOutStr = Replace(sOutStr, vbNullChar, "")
    
    RemovePunctuation = sOutStr

End Function
