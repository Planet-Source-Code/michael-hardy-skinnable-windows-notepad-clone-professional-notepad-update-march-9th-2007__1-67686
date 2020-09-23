VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhoneme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// Soundex

'// From Wikipedia, the free encyclopedia

'// Soundex is a phonetic algorithm for indexing names by their
'// sound when pronounced in English.

'// Soundex is the most widely known of all phonetic algorithms and
'// is often used (incorrectly) as a synonym for "phonetic algorithm".

'// The basic aim is for names with the same pronunciation to be
'// encoded to the same string so that matching can occur despite
'// minor differences in spelling.

'// The Soundex code for a name consists of a letter followed by three
'// numbers: the letter is the first letter of the name, and the numbers
'// encode the remaining consonants.

'// Similar sounding consonants share the same number so, for example,
'// the labial B, F, P and V are all encoded as 1.

'// If two or more letters with the same number were adjacent in the
'// original name, or adjacent except for any intervening vowels, then
'// all are omitted except the first.

'// Vowels can affect the coding, but are never coded directly unless
'// they appear at the start of the name.

'// This particular version of the Soundex algorithm has been adapted
'// from the original design in an attempt to more reliably produce
'// word matching for a generic English language spell checker.

'// Normally, each Soundex begins with the first letter of the given
'// name and only subsequent letters are used to produce the phonetic
'// signature, so only names beginning with the same first letter are
'// compared for similar pronunciation using the standard algorithm.

'// For example, one may seek the correct spelling for "upholstery" and
'// may inadvertently type "apolstry", "apolstery", or even "apholstery"
'// but would still not retrieve the correct spelling for this word.

'// Therefore, this version of the Soundex algorithm has been modified
'// to allow the matching of words that start with differing first
'// letters so as not to assume that the first letter is always known.

'// Consequently, all Soundex's begin with the letter "S", and the
'// encoding always begins with the first letter of the word.

'// Because of this change, many more similarly spelled words are
'// returned as a match, so the Soundex's length has also been
'// extended from three numbers to four to produce a slightly more
'// unique phonetic signature.

'// Returns the 5 letter Soundex for an English word.
Public Function GetSoundexWord(argWord As String) As String
    Dim workStr As String, i As Long
    Dim replaceMask(1& To 6&) As Boolean
    If LenB(argWord) = 0& Then Exit Function

    '// Normalise it to remove ambiguity
    argWord = LCase$(argWord)

    workStr = "S" '+ 4 number encoding

    '// Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6

    For i = 1& To Len(argWord)
        Select Case Mid$(argWord, i, 1&)
            Case "a", "e", "h", "i", "o", "u", "w", "y"
                '// do nothing

            Case "b", "f", "p", "v"
                If replaceMask(1&) = False Then
                    workStr = workStr & Chr$(49) '// 1
                    replaceMask(1&) = True
                End If

            Case "c", "g", "j", "k", "q", "s", "x", "z"
                If replaceMask(2&) = False Then
                    workStr = workStr & Chr$(50) '// 2
                    replaceMask(2&) = True
                End If

            Case "d", "t"
                If replaceMask(3&) = False Then
                    workStr = workStr & Chr$(51) '// 3
                    replaceMask(3&) = True
                End If

            Case "l"
                If replaceMask(4&) = False Then
                    workStr = workStr & Chr$(52) '// 4
                    replaceMask(4&) = True
                End If

            Case "m", "n"
                If replaceMask(5&) = False Then
                    workStr = workStr & Chr$(53) '// 5
                    replaceMask(5&) = True
                End If

            Case "r"
                If replaceMask(6&) = False Then
                    workStr = workStr & Chr$(56) '// 6
                    replaceMask(6&) = True
                End If
        End Select
    Next i

    '// Return the first five chars padded with 0's.
    If Len(workStr) > 5& Then
        GetSoundexWord = Left$(workStr, 5&)
    Else
        GetSoundexWord = workStr & Space$(5& - Len(workStr))
    End If
End Function

'// Returns the Minimum of 3 numbers
Private Function min3(ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long) As Long
    min3 = n1
    If n2 < min3 Then min3 = n2
    If n3 < min3 Then min3 = n3
End Function

'// Levenshtein Distance

'// From Wikipedia, the free encyclopedia

'// In information theory and computer science, the Levenshtein distance
'// or edit distance between two strings is given by the minimum number
'// of operations needed to transform one string into the other, where
'// an operation is an insertion, deletion, or substitution of a single
'// character.

'// It is named after Vladimir Levenshtein, who considered this distance
'// in 1965.

'// It is useful in applications that need to determine how similar two
'// strings are, such as spell checkers.

'// It can be considered a generalisation of the Hamming distance, which
'// is used for strings of the same length and only considers substitution
'// edits.

'// There are also further generalisations of the Levenshtein distance
'// that consider, for example, exchanging two characters as an operation,
'// like in the Damerau-Levenshtein distance algorithm.

'// Returns the Levenshtein Distance between 2 strings.
Public Function GetLevenshteinDistance(argStr1 As String, argStr2 As String) As Long
    Dim LenStr1 As Long, LenStr2 As Long
    Dim editMatrix() As Long, i As Long, j As Long, cost As Long
    Dim str1_i As String, str2_j As String
    Dim p() As Long, q() As Long, r As Long
    Dim x As Long, y As Long

    LenStr1 = Len(argStr1)
    LenStr2 = Len(argStr2)

    If LenStr1 = 0& Then
        '// The length of Str2 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr2

    ElseIf LenStr2 = 0& Then
        '// The length of Str1 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr1

    Else
        '// editMatrix is a table with lenStr1+1 rows and lenStr2+1 columns
        ReDim editMatrix(LenStr1, LenStr2) As Long

        '// i and j are used to iterate over str1 and str2
        For i = 0& To LenStr1
            editMatrix(i, 0&) = i
        Next
    
        For j = 0& To LenStr2
            editMatrix(0&, j) = j
        Next
    
        For i = 1& To LenStr1
            str1_i = Mid$(argStr1, i, 1&)
            For j = 1& To LenStr2
                str2_j = Mid$(argStr2, j, 1&)
                If str1_i = str2_j Then
                    cost = 0&
                Else
                    cost = 1&
                End If

                '//                     deletion,insertion,substitution
                editMatrix(i, j) = min3(editMatrix(i - 1&, j) + 1&, _
                                        editMatrix(i, j - 1&) + 1&, _
                                        editMatrix(i - 1&, j - 1&) + cost)
            Next j
        Next i
    
        GetLevenshteinDistance = editMatrix(LenStr1, LenStr2)
        Erase editMatrix
    End If
End Function
