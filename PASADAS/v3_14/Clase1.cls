VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ***    Made By Michael Ciurescu (CVMichael)   ***
'Modified by Paulo Cunha (pcunha) to work with char128.ttf on word or excel on 16/05/2011
'the font at in http://grandzebu.net/index.php?page=/informatique/codbar-en/code128.htm


' References:
' http://www.barcodeman.com/info/c128.php3

'This code has been modified. It erroneously trimmed and deleted characters and spaces.

Private Enum eCode128Type
    eCode128_CodeSetA = 1
    eCode128_CodeSetB = 2
    eCode128_CodeSetC = 3
End Enum

Private Type tCode
    ASet As String
    BSet As String
    CSet As String
    BarSpacePattern As String
End Type

Private CodeArr() As tCode

Private Sub Class_Initialize()
    ReDim CodeArr(106)

    AddEntry 0, " ", " ", "00", Chr(32)
    AddEntry 1, "!", "!", "01", Chr(33)
    AddEntry 2, """", """", "02", Chr(34)
    AddEntry 3, "#", "#", "03", Chr(35)
    AddEntry 4, "$", "$", "04", Chr(36)
    AddEntry 5, "%", "%", "05", Chr(37)
    AddEntry 6, "&", "&", "06", Chr(38)
    AddEntry 7, "'", "'", "07", Chr(39)
    AddEntry 8, "(", "(", "08", Chr(40)
    AddEntry 9, ")", ")", "09", Chr(41)
    AddEntry 10, "*", "*", "10", Chr(42)
    AddEntry 11, "+", "+", "11", Chr(43)
    AddEntry 12, ",", ",", "12", Chr(44)
    AddEntry 13, "-", "-", "13", Chr(45)
    AddEntry 14, ".", ".", "14", Chr(46)
    AddEntry 15, "/", "/", "15", Chr(47)
    AddEntry 16, "0", "0", "16", Chr(48)
    AddEntry 17, "1", "1", "17", Chr(49)
    AddEntry 18, "2", "2", "18", Chr(50)
    AddEntry 19, "3", "3", "19", Chr(51)
    AddEntry 20, "4", "4", "20", Chr(52)
    AddEntry 21, "5", "5", "21", Chr(53)
    AddEntry 22, "6", "6", "22", Chr(54)
    AddEntry 23, "7", "7", "23", Chr(55)
    AddEntry 24, "8", "8", "24", Chr(56)
    AddEntry 25, "9", "9", "25", Chr(57)
    AddEntry 26, ":", ":", "26", Chr(58)
    AddEntry 27, ";", ";", "27", Chr(59)
    AddEntry 28, "<", "<", "28", Chr(60)
    AddEntry 29, "=", "=", "29", Chr(61)
    AddEntry 30, ">", ">", "30", Chr(62)
    AddEntry 31, "?", "?", "31", Chr(63)
    AddEntry 32, "@", "@", "32", Chr(64)
    AddEntry 33, "A", "A", "33", Chr(65)
    AddEntry 34, "B", "B", "34", Chr(66)
    AddEntry 35, "C", "C", "35", Chr(67)
    AddEntry 36, "D", "D", "36", Chr(68)
    AddEntry 37, "E", "E", "37", Chr(69)
    AddEntry 38, "F", "F", "38", Chr(70)
    AddEntry 39, "G", "G", "39", Chr(71)
    AddEntry 40, "H", "H", "40", Chr(72)
    AddEntry 41, "I", "I", "41", Chr(73)
    AddEntry 42, "J", "J", "42", Chr(74)
    AddEntry 43, "K", "K", "43", Chr(75)
    AddEntry 44, "L", "L", "44", Chr(76)
    AddEntry 45, "M", "M", "45", Chr(77)
    AddEntry 46, "N", "N", "46", Chr(78)
    AddEntry 47, "O", "O", "47", Chr(79)
    AddEntry 48, "P", "P", "48", Chr(80)
    AddEntry 49, "Q", "Q", "49", Chr(81)
    AddEntry 50, "R", "R", "50", Chr(82)
    AddEntry 51, "S", "S", "51", Chr(83)
    AddEntry 52, "T", "T", "52", Chr(84)
    AddEntry 53, "U", "U", "53", Chr(85)
    AddEntry 54, "V", "V", "54", Chr(86)
    AddEntry 55, "W", "W", "55", Chr(87)
    AddEntry 56, "X", "X", "56", Chr(88)
    AddEntry 57, "Y", "Y", "57", Chr(89)
    AddEntry 58, "Z", "Z", "58", Chr(90)
    AddEntry 59, "[", "[", "59", Chr(91)
    AddEntry 60, "\", "\", "60", Chr(92)
    AddEntry 61, "]", "]", "61", Chr(93)
    AddEntry 62, "^", "^", "62", Chr(94)
    AddEntry 63, "_", "_", "63", Chr(95)
    AddEntry 64, Chr(0), "`", "64", Chr(96) ' Null
    AddEntry 65, Chr(1), "a", "65", Chr(97) ' SOH
    AddEntry 66, Chr(2), "b", "66", Chr(98) ' STX
    AddEntry 67, Chr(3), "c", "67", Chr(99) ' ETX
    AddEntry 68, Chr(4), "d", "68", Chr(100) ' EOT
    AddEntry 69, Chr(5), "e", "69", Chr(101) ' ENQ
    AddEntry 70, Chr(6), "f", "70", Chr(102) ' ACK
    AddEntry 71, Chr(7), "g", "71", Chr(103) ' BEL
    AddEntry 72, Chr(8), "h", "72", Chr(104) ' BS
    AddEntry 73, Chr(9), "i", "73", Chr(105) ' HT
    AddEntry 74, Chr(10), "j", "74", Chr(106) ' LF
    AddEntry 75, Chr(11), "k", "75", Chr(107) ' VT
    AddEntry 76, Chr(12), "l", "76", Chr(108) ' FF
    AddEntry 77, Chr(13), "m", "77", Chr(109) ' CR
    AddEntry 78, Chr(14), "n", "78", Chr(110) ' SO
    AddEntry 79, Chr(15), "o", "79", Chr(111) ' SI
    AddEntry 80, Chr(16), "p", "80", Chr(112) ' DLE
    AddEntry 81, Chr(17), "q", "81", Chr(113) ' DC1
    AddEntry 82, Chr(18), "r", "82", Chr(114) ' DC2
    AddEntry 83, Chr(19), "s", "83", Chr(115) ' DC3
    AddEntry 84, Chr(20), "t", "84", Chr(116) ' DC4
    AddEntry 85, Chr(21), "u", "85", Chr(117) ' NAK
    AddEntry 86, Chr(22), "v", "86", Chr(118) ' SYN
    AddEntry 87, Chr(23), "w", "87", Chr(119) ' ETB
    AddEntry 88, Chr(24), "x", "88", Chr(120) ' CAN
    AddEntry 89, Chr(25), "y", "89", Chr(121) ' EM
    AddEntry 90, Chr(26), "z", "90", Chr(122) ' SUB
    AddEntry 91, Chr(27), "{", "91", Chr(123) ' ESC
    AddEntry 92, Chr(28), "|", "92", Chr(124) ' FS
    AddEntry 93, Chr(29), "}", "93", Chr(125) ' GS
    AddEntry 94, Chr(30), "~", "94", Chr(126) ' RS
    AddEntry 95, Chr(31), Chr(127), "95", Chr(200) ' US, DEL
    AddEntry 96, "FNC 3", "FNC 3", "96", Chr(201)
    AddEntry 97, "FNC 2", "FNC 2", "97", Chr(202)
    AddEntry 98, "SHIFT", "SHIFT", "98", Chr(203)
    AddEntry 99, "CODE C", "CODE C", "99", Chr(204)
    AddEntry 100, "CODE B", "FNC 4", "CODE B", Chr(205)
    AddEntry 101, "FNC 4", "CODE A", "CODE A", Chr(206)
    AddEntry 102, "FNC 1", "FNC 1", "FNC 1", Chr(207)
    AddEntry 103, "Start A", "Start A", "Start A", Chr(208)
    AddEntry 104, "Start B", "Start B", "Start B", Chr(209)
    AddEntry 105, "Start C", "Start C", "Start C", Chr(210)
    AddEntry 106, "Stop", "Stop", "Stop", Chr(211)
End Sub

Private Sub AddEntry(ByVal Index As Integer, ASet As String, BSet As String, CSet As String, BarSpacePattern As String)
    With CodeArr(Index)
        .ASet = ASet
        .BSet = BSet
        .CSet = CSet
        'The original line of code is also shown.
        .BarSpacePattern = BarSpacePattern '.BarSpacePattern = Replace(BarSpacePattern, " ", "")
    End With
End Sub

Public Function Code128_Str(ByVal Str As String)
    Code128_Str = BuildStr(Str) '= Replace(BuildStr(Str), " ", "")
End Function

Private Function BuildStr(ByVal Str As String) As String
    Dim SCode As eCode128Type
    Dim PrevSCode As eCode128Type
    Dim CurrChar As String
    Dim ArrIndex As Integer
    Dim CharIndex As Long
    Dim CheckDigit As Integer
    Dim CCodeIndex As Integer
    Dim TotalSum As Long

    'Set Scode to the default chaacter set B.
    SCode = eCode128_CodeSetB
    
    'Does the string commence with 2 single digit (0–9) characters?
    'If so then use character set C.
    '************************************************************************************
    'Disabled to force character set B.
    If Str Like "##*" Then SCode = eCode128_CodeSetC
    '************************************************************************************

    TotalSum = 0
    CharIndex = 1

    'Add the initial character to tell the barcode reader what character set is being used.
    Select Case SCode
    '************************************************************************************
    'How do you ever get to use character set A?
    '************************************************************************************
    Case eCode128_CodeSetA
        TotalSum = TotalSum + (103 * CharIndex)
        BuildStr = Trim(BuildStr) & Chr(208)
    Case eCode128_CodeSetB
        TotalSum = TotalSum + (104 * CharIndex)
        BuildStr = Trim(BuildStr) & Chr(209)
    Case eCode128_CodeSetC
        TotalSum = TotalSum + (105 * CharIndex)
        BuildStr = BuildStr & Chr(210) 'BuildStr = Trim(BuildStr) & Chr(210)
    End Select

    PrevSCode = SCode

    Do Until Len(Str) = 0
        'Does the string commence with 4 single digit (0–9) characters?
        'If so then use character set C.
        '************************************************************************************
        'Disabled to force character set B.
        If Str Like "####*" Then SCode = eCode128_CodeSetC
        '************************************************************************************

        If SCode = eCode128_CodeSetC And Mid(Str, 1, 2) Like "##" Then
            CurrChar = Mid(Str, 1, 2)
        Else
            CurrChar = Mid(Str, 1, 1)
        End If

        ArrIndex = GetCharIndex(CurrChar, SCode, True)

        If ArrIndex <> -1 Then
            If CodeArr(ArrIndex).BSet = CurrChar And ((SCode = eCode128_CodeSetC And CodeArr(ArrIndex).CSet <> CurrChar) Or (SCode = eCode128_CodeSetA And CodeArr(ArrIndex).ASet <> CurrChar)) Then
                SCode = eCode128_CodeSetB
            ElseIf CodeArr(ArrIndex).ASet = CurrChar And CodeArr(ArrIndex).BSet <> CurrChar Then
                SCode = eCode128_CodeSetA
            ElseIf CodeArr(ArrIndex).CSet = CurrChar Then
                SCode = eCode128_CodeSetC
            End If

            If PrevSCode <> SCode Then
                Select Case SCode
                Case eCode128_CodeSetA
                    CCodeIndex = GetCharIndex("CODE A", PrevSCode, False)
                Case eCode128_CodeSetB
                    CCodeIndex = GetCharIndex("CODE B", PrevSCode, False)
                Case eCode128_CodeSetC
                    CCodeIndex = GetCharIndex("CODE C", PrevSCode, False)
                End Select

                TotalSum = TotalSum + (CCodeIndex * CharIndex)
                'The original line of code is also shown.
                BuildStr = BuildStr & CodeArr(CCodeIndex).BarSpacePattern 'BuildStr = Trim(BuildStr) & CodeArr(CCodeIndex).BarSpacePattern

                CharIndex = CharIndex + 1
                PrevSCode = SCode
            End If
            
            'The original line of code is also shown.
            BuildStr = BuildStr & CodeArr(ArrIndex).BarSpacePattern 'BuildStr = Trim(BuildStr) & CodeArr(ArrIndex).BarSpacePattern

            TotalSum = TotalSum + (ArrIndex * CharIndex)
            CharIndex = CharIndex + 1
        End If

        If SCode = eCode128_CodeSetC Then
            Str = Mid(Str, 3)
        Else
            Str = Mid(Str, 2)
        End If
    Loop

    'Determine the check digit.
    'ie Remainder after dividing by 103.
    CheckDigit = TotalSum Mod 103

    'Add the CheckDigit to the end of the string.
    'The original line of code is also shown.
    BuildStr = BuildStr & CodeArr(CheckDigit).BarSpacePattern 'BuildStr = Trim(BuildStr) & CodeArr(CheckDigit).BarSpacePattern
    
    'Add the final end character to the barcode.
    'The original line of code is also shown.
    BuildStr = BuildStr & Chr(211) 'BuildStr = Trim(BuildStr) & Chr(211)
End Function

Private Function GetCharIndex(ByVal Char As String, ByVal CodeType As eCode128Type, ByVal Recurse As Boolean) As Integer
    Dim k As Long

    Select Case CodeType
    Case eCode128_CodeSetA
        For k = 0 To UBound(CodeArr)
            If Char = CodeArr(k).ASet Then Exit For
        Next k
    Case eCode128_CodeSetB
        For k = 0 To UBound(CodeArr)
            If Char = CodeArr(k).BSet Then Exit For
        Next k
    Case eCode128_CodeSetC
        For k = 0 To UBound(CodeArr)
            If Char = CodeArr(k).CSet Then Exit For
        Next k
    End Select

    If k = UBound(CodeArr) + 1 Then
        If Not Recurse Then
            GetCharIndex = -1
        Else
            Select Case CodeType
            Case eCode128_CodeSetA
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetC, False)
            Case eCode128_CodeSetB
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetA, False)
            Case eCode128_CodeSetC
                GetCharIndex = GetCharIndex(Char, eCode128_CodeSetB, False)
            End Select

            If GetCharIndex = -1 Then
                Select Case CodeType
                Case eCode128_CodeSetA
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetB, False)
                Case eCode128_CodeSetB
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetC, False)
                Case eCode128_CodeSetC
                    GetCharIndex = GetCharIndex(Char, eCode128_CodeSetA, False)
                End Select
            End If
        End If
    Else
        GetCharIndex = k
    End If
End Function

Public Function Code128_GetWidth(ByVal Str As String, Optional ByVal BarWidth As Integer = 1) As Long
    Dim k As Long, Width As Long

    Str = Replace(Code128_Str(Str), " ", "")
    Debug.Print Str
    For k = 1 To Len(Str)
        Width = Width + Val(Mid(Str, k, 1))
    Next k

    Code128_GetWidth = Width * BarWidth + (28 * BarWidth)
End Function
