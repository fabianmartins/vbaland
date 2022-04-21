Function GetBySize(stringA, stringB, smallOrLarger)
    If (Len(stringA) >= Len(stringB)) Then
        LargerString = stringA
        SmallerString = stringB
    Else
        LargerString = stringB
        SmallerString = stringA
    End If
    Result = ""
    If (smallOrLarger >= 0) Then
        Result = LargerString
    Else
        Result = SmallerString
    End If
    GetBySize = Result
End Function

Function SanitizeStr(sourceStr)
    TheStr = sourceStr
    LenStr = Len(TheStr)
    i = 1
    While (i < LenStr)
        TheChar = Mid(TheStr, i, 1)
        AsciiCode = Asc(TheChar)
        If Not ((AsciiCode = 32) Or (AsciiCode >= 97 And AsciiCode <= 122) Or (AsciiCode >= 65 And AsciiCode <= 132) Or (AsciiCode >= 48 And AsciiCode <= 57)) Then
            TheStr = Replace(TheStr, TheChar, "")
        End If
        LenStr = Len(TheStr)
        i = i + 1
    Wend
    SanitizeStr = TheStr
End Function
Function GetWordHitPercentage(wordSourceString, searchWithinString)
    Result = 0
    SpacePosition = 1
    WordHitCount = 0
    TotalWords = 0
    While SpacePosition < Len(wordSourceString)
        NextSpacePosition = InStr(SpacePosition + 1, wordSourceString, Space(1), vbTextCompare)
        If (SpacePosition > NextSpacePosition) Then
            NextSpacePosition = Len(wordSourceString) + 1
        End If
        Word = Mid(wordSourceString, SpacePosition, NextSpacePosition - 1)
        If InStr(1, searchWithinString, Word, vbTextCompare) > 0 Then
            WordHitCount = WordHitCount + 1
        End If
        TotalWords = TotalWords + 1
        SpacePosition = NextSpacePosition + 1
    Wend
    GetWordHitPercentage = WordHitCount / TotalWords
End Function

Function GetCharacterHitPercentage(wordSourceString, searchWithinString)
    Same = 0
    Str1 = searchWithinString
    Str2 = wordSourceString
    For c = 1 To Len(Str2)
        If InStr(1, Str1, Mid(Str2, c, 1), 1) Then
            Same = Same + 1
            Str1 = Replace(Str1, Mid(Str2, c, 1), "*", 1, 1)
        End If
    Next c
    GetCharacterHitPercentage = Same / Len(Str1)
End Function

Sub StrPercent()
    Application.ScreenUpdating = True
    LR = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    With ActiveSheet
         .Range("A1:A" & LR).Value = ActiveSheet.Range("A1:A" & LR).Value
         .Range("B1:B" & LR).Value = ActiveSheet.Range("B1:B" & LR).Value
        For r = 2 To LR
            EqualityPct = 0
            WordHitPct = 0
            CharHitPct = 0
            Str1 = Trim(.Cells(r, "A").Value)
            Str2 = Trim(.Cells(r, "B").Value)
            ''--- Sanitize Strings - Remove non-letters characters
            Str1 = SanitizeStr(Str1)
            Str2 = SanitizeStr(Str2)
            ''--- Check Identicals
            If LCase(Str1) = LCase(Str2) Then
                EqualityPct = 1
                GoTo Done
            End If
            ''--- Check words
            LargerString = GetBySize(Str1, Str2, 1)
            SmallerString = GetBySize(Str1, Str2, -1)
            WordHitPct = GetWordHitPercentage(SmallerString, LargerString)
            ''--- Check per character INEFFICIENT
            ''CharHitPct = GetCharacterHitPercentage(SmallerString, LargerString)
Done:
        If EqualityPct >= WordHitPct Then
            .Cells(r, "C").Value = EqualityPct
        Else
            .Cells(r, "C").Value = WordHitPct
        End If
        Next r
    End With
    Application.ScreenUpdating = True
End Sub