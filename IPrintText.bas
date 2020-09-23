Attribute VB_Name = "IPrintText"
Option Explicit 'Handles the font for the iPod
'The iPod uses the Mac font Chicago, which I was unable to find myself.
'All functions are tailor made to fit a bitmap of the font I made
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SpecialChars As String = "> <up> <down> <play> <pause> : - ( ) . [ ] { } \ &"
Private Const SpecialCharWidth As String = "7 8 8 10 7 3 6 4 4 3 4 4 5 5 5 9"
Public Sub TransBLT(srcHdc As Long, xSrc As Long, ySrc As Long, MaskHDC As Long, Xmsk As Long, Ymsk As Long, Width As Long, height As Long, destHdc As Long, X As Long, Y As Long)
    Const SRCPAINT = &HEE0086 'Assumes the mask matches the source's coordinates
    BitBlt destHdc, X, Y, Width, height, MaskHDC, Xmsk, Ymsk, SRCPAINT
    BitBlt destHdc, X, Y, Width, height, srcHdc, xSrc, ySrc, vbSrcAnd
End Sub
Public Function iPrint(text As String, srcHdc As Long, destHdc As Long, X As Long, ByVal Y As Long, Hi As Boolean)
    If InStr(text, vbNewLine) = 0 Then
        PrintLine text, srcHdc, destHdc, X, Y, Hi
    Else
        Dim temp As Long, tempstr() As String
        tempstr = Split(text, vbNewLine)
        For temp = 0 To UBound(tempstr)
            PrintLine tempstr(temp), srcHdc, destHdc, X, Y, Hi
            Y = Y + StringHeight(tempstr(temp))
        Next
    End If
End Function
Private Function PrintLine(ByVal text As String, srcHdc As Long, destHdc As Long, ByVal X As Long, Y As Long, Hi As Boolean)
    Dim tempstr As String
    Do Until Len(text) = 0
        tempstr = StripWord(text)
        DrawChar tempstr, destHdc, srcHdc, X, Y, Hi
        X = X + CharWidth(tempstr)
    Loop
End Function
Private Function DrawChar(letter As String, destHdc As Long, srcHdc As Long, X As Long, Y As Long, Highlite As Boolean) As Long
    Dim Xany As Long, Ymsk As Long, ySrc As Long, Width As Long, height As Long
    If Len(letter) = 0 Then letter = " "
    Width = CharWidth(letter)
    If letter >= "a" And letter <= "z" Then
        SetLoc Xany, ySrc, Ymsk, height, (Asc(letter) - 97) * 11 + 1, 1, 14, 12
    Else
        If letter >= "A" And letter <= "Z" Then
            SetLoc Xany, ySrc, Ymsk, height, (Asc(letter) - 65) * 11 + 1, 27, 40, 12
        Else
            If letter >= "0" And letter <= "9" Then
                SetLoc Xany, ySrc, Ymsk, height, (Asc(letter) - 48) * 11 + 1, 55, 68, 12
            Else
                letter = LCase(letter)
                Select Case letter
                    Case ">", "<up>", "<down>", "<play>", "<pause>", ":", "-", "(", ")", ".", "[", "]", "{", "}", "\", "&"
                        SetLoc Xany, ySrc, Ymsk, height, 111 + (GetIndex(SpecialChars, letter) * 11), 55, 68, 12
                    Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>"
                        SetLoc Xany, ySrc, Ymsk, height, (Asc(Mid(letter, 3, 1)) - 48) * 19 + 1, 83, 111, 27
                    Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>"
                        SetLoc Xany, ySrc, Ymsk, height, (Asc(Mid(letter, 3, 1)) - 65) * 5 + 212, 83, 90, 6
                    Case " ":           SetLoc Xany, ySrc, Ymsk, height, 275, 83, 96, 12
                    Case "<dir>":       SetLoc Xany, ySrc, Ymsk, height, 275, 109, 122, 12
                    Case "<b:>":        SetLoc Xany, ySrc, Ymsk, height, 191, 83, 111, 27
                    Case "<repeat>":    SetLoc Xany, ySrc, Ymsk, height, 233, 115, 141, 7
                    Case "<shuffle>":   SetLoc Xany, ySrc, Ymsk, height, 254, 115, 141, 7
                    Case "<sun>":       SetLoc Xany, ySrc, Ymsk, height, 212, 99, 125, 7
                    Case "<mon>":       SetLoc Xany, ySrc, Ymsk, height, 233, 99, 125, 7
                    Case "<tue>":       SetLoc Xany, ySrc, Ymsk, height, 254, 99, 125, 7
                    Case "<wed>":       SetLoc Xany, ySrc, Ymsk, height, 212, 107, 133, 7
                    Case "<thu>":       SetLoc Xany, ySrc, Ymsk, height, 233, 107, 133, 7
                    Case "<fri>":       SetLoc Xany, ySrc, Ymsk, height, 254, 107, 133, 7
                    Case "<sat>":       SetLoc Xany, ySrc, Ymsk, height, 212, 115, 141, 7
                End Select
            End If
        End If
    End If
    
    If Not Highlite Then
        TransBLT srcHdc, Xany, ySrc, srcHdc, Xany, Ymsk, Width, height, destHdc, X, Y
    Else
        TransBLT srcHdc, Xany, Ymsk, srcHdc, Xany, ySrc, Width, height, destHdc, X, Y
    End If
    DrawChar = X + Width - 1
End Function
Public Function CharHeight(letter As String) As Long
    'If (letter >= "a" And letter <= "z") Or (letter >= "A" And letter <= "Z") Or (letter >= "0" And letter <= "9") Then
    Select Case letter
        Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>", "<b:>": CharHeight = 27
        Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>": CharHeight = 6
        Case "<repeat>", "<shuffle>", "<sun>", "<mon>", "<tue>", "<wed>", "<thu>", "<fri>", "<sat>": CharHeight = 7
        Case Else: CharHeight = 12
    End Select
End Function

Private Function CharExists(letter As String) As Boolean
    CharExists = CharWidth(letter) > 0
End Function
Private Function CharWidth(ByVal letter As String) As Long
If Left(letter, 1) = "<" And Right(letter, 1) = ">" Then letter = LCase(letter)
Select Case letter
    'Lower Case
    Case "i", "l":                          CharWidth = 3
    Case "t":                               CharWidth = 5
    Case "c", "f", "k", "r", "s", "z":      CharWidth = 6
    Case "q":                               CharWidth = 8
    Case "m", "w":                          CharWidth = 11
    Case "a", "b", "d", "e", "g", "h", "j", "n", "o", "p", "u", "v", "x", "y": CharWidth = 7

    'Upper Case
    Case "E", "F", "L", "S", "Z":       CharWidth = 6
    Case "K", "N", "Q":                 CharWidth = 8
    Case "M", "W":                      CharWidth = 11
    Case "A", "B", "C", "D", "G", "H", "I", "J", "O", "P", "R", "T", "U", "V", "X", "Y": CharWidth = 7

    'Numbers and Special Characters
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": CharWidth = 7 'Normal sized numbers
    Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>": CharWidth = 18  'Large numbers
    Case "<b:>": CharWidth = 8 'Big semicolon (:)
    Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>": CharWidth = 4 'Small numbers
    Case " ": CharWidth = 4 'Space
    Case "<dir>": CharWidth = 11
    Case ">", "<up>", "<down>", "<play>", "<pause>", ":", "-", "(", ")", ".", "[", "]", "{", "}", "\", "&": CharWidth = Val(GetFromIndex(SpecialChars, letter, SpecialCharWidth)) 'Special Chars (Punctuation)
    Case "<repeat>", "<shuffle>", "<sun>", "<mon>", "<tue>", "<wed>", "<thu>", "<fri>", "<sat>": CharWidth = 20 'Special Chars (Days of the week and play modes)
End Select
End Function
Private Sub SetLoc(Xany As Long, ySrc As Long, Ymsk As Long, height As Long, X As Long, y1 As Long, y2 As Long, hit As Long)
    Xany = X
    ySrc = y1
    Ymsk = y2
    height = hit
End Sub
Private Function GetIndex(text As String, word As String, Optional delimeter As String = " ") As Long
    Dim tempstr() As String, temp As Long
    GetIndex = -1
    tempstr = Split(text, delimeter)
    For temp = 0 To UBound(tempstr)
        If tempstr(temp) = word Then
            GetIndex = temp
            Exit For
        End If
    Next
End Function
Private Function GetFromIndex(text As String, word As String, text2 As String, Optional delimeter As String = " ") As String
    Dim temp As Long, tempstr() As String
    temp = GetIndex(text, word, delimeter)
    If temp > -1 Then
        tempstr = Split(text2, delimeter)
        GetFromIndex = tempstr(temp)
    End If
End Function
Private Function WordLength(text As String, start As Long) As Long
    Dim temp As Long, tempstr As String
    WordLength = 1
    If Mid(text, start, 1) = "<" Then
        temp = InStr(start, text, ">")
        If temp > start Then
            tempstr = Mid(text, start, temp - start + 1)
            If CharExists(tempstr) Then WordLength = Len(tempstr)
        End If
    End If
End Function
Private Function StripWord(ByRef text As String) As String
    Dim temp As Long
    temp = WordLength(text, 1)
    StripWord = Left(text, temp)
    text = Right(text, Len(text) - temp)
End Function
Public Function StringWidth(ByVal text As String) As Long
    Dim temp As Long
    Do Until Len(text) = 0
        temp = temp + CharWidth(StripWord(text))
    Loop
    StringWidth = temp
End Function
Public Function StringHeight(ByVal text As String) As Long
    Dim temp As Long, temp2 As Long
    Do Until Len(text) = 0
        temp2 = CharHeight(StripWord(text))
        If temp2 > temp Then temp = temp2
    Loop
    StringHeight = temp
End Function
Public Function GetTime() As String
    Dim temp As String, tempstr As String
    temp = Time
    Do Until Left(temp, 1) = " "
        tempstr = tempstr & "<b" & Left(temp, 1) & ">"
        temp = Right(temp, Len(temp) - 1)
    Loop
    GetTime = tempstr
End Function


