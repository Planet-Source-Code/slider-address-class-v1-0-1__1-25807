VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cAddress
' Author:       Slider
' Date:         04/08/2001
' Version:      01.00.01
' Description:  Handles the creation plus breaking down of an address
'               with auto-capitalise (option), Fills ComboBox with all
'               currently known countries and criteria-based validation.
' Edit History: 01.00.00 03/08/01 Initial Release
'               01.00.01 04/08/01 Forgot to check for extra spaces - Fixed.
'               01.00.01 04/08/01 Added support to handle blank lines.
'               01.00.01 04/08/01 Fixed comma handling when breaking out
'                                 City, State & Postcode.
'
'===========================================================================

Option Explicit

Private msAddress     As String

Private msStreet      As String
Private msCity        As String
Private msState       As String
Private msPostcode    As String
Private msCountry     As String

Private msCountries() As String

Private mbAutoCorrect As Boolean

Public Enum eAddressComboType
    eactCountry = 0
End Enum

Public Enum eAddressValidateFields
    eavfStreet = 1
    eavfCity = 2
    eavfState = 4
    eavfPostcode = 8
    eavfCountry = 16
End Enum

'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Property Get Address() As String
    Address = pBuildAddress
End Property

Public Property Let Address(ByVal vNewValue As String)
    msAddress = CompactSpaces(vNewValue)
    If Len(msAddress) Then pExtractAddress
End Property

Public Property Get Street() As String
    Street = msStreet
End Property

Public Property Let Street(ByVal vNewValue As String)
    
    Dim lCount    As Long
    Dim lLoop     As Long
    Dim sTokens() As String
    Dim sTmp      As String
    
    '## Reformat the passed street details
    sTokens = Split(CompactSpaces(vNewValue), vbCrLf)
    lCount = UBound(sTokens)
    For lLoop = 0 To lCount
        sTmp = sTmp + vbCrLf + pStripRightComa(sTokens(lLoop))
    Next
    vNewValue = Trim$(Mid$(sTmp, 3))
    msStreet = IIf(mbAutoCorrect, pProperCasing(vNewValue), vNewValue)

End Property

Public Property Get City() As String
    City = msCity
End Property

Public Property Let City(ByVal vNewValue As String)
    vNewValue = pStripRightComa(CompactSpaces(vNewValue))
    msCity = IIf(mbAutoCorrect, pProperCasing(vNewValue), vNewValue)
End Property

Public Property Get State() As String
    State = msState
End Property

Public Property Let State(ByVal vNewValue As String)
    vNewValue = pStripRightComa(CompactSpaces(vNewValue))
    msState = IIf(mbAutoCorrect, pProperCasing(vNewValue), vNewValue)
End Property

Public Property Get Postcode() As String
    Postcode = msPostcode
End Property

Public Property Let Postcode(ByVal vNewValue As String)
    vNewValue = pStripRightComa(CompactSpaces(vNewValue))
    msPostcode = IIf(mbAutoCorrect, pProperCase(vNewValue), vNewValue)
End Property

Public Property Get Country() As String
    Country = msCountry
End Property

Public Property Let Country(ByVal vNewValue As String)

    Dim lLoop As Long
    Dim bFnd  As Boolean

    '## Validate that it's a recognised country
    For lLoop = 1 To UBound(msCountries)
        If UCase$(vNewValue) = UCase$(msCountries(lLoop)) Then
            If mbAutoCorrect Then
                msCountry = msCountries(lLoop)
            Else
                msCountry = vNewValue
            End If
            bFnd = True
            Exit For
        End If
    Next
    If Not bFnd Then
        msCountry = vNewValue
    End If

End Property

Public Property Get AutoCorrect() As Boolean
    AutoCorrect = mbAutoCorrect
End Property

Public Property Let AutoCorrect(ByVal State As Boolean)
    mbAutoCorrect = State
End Property

Public Sub FillComboBox(ByRef oCombo As VB.ComboBox, _
                        ByVal ComboType As eAddressComboType)

    Dim lLoop As Long
    Dim lMax  As Long

    With oCombo
        .Clear
        Select Case ComboType
            Case eactCountry
                lMax = UBound(msCountries)
                For lLoop = 1 To lMax
                    .AddItem msCountries(lLoop)
                    .ItemData(.NewIndex) = lLoop
                Next

        End Select
        If .ListCount Then
            .ListIndex = 0
        End If
    End With

End Sub

'
'## Will proceess and return (set bits) missing fields from the passed address
'
Public Function ValidateAddress(ByVal Address As String, _
                                ByVal Fields As eAddressValidateFields) As eAddressValidateFields

    Dim stAddress  As String
    Dim stStreet   As String
    Dim stCity     As String
    Dim stState    As String
    Dim stPostcode As String
    Dim stCountry  As String

    Dim eResult    As eAddressValidateFields    '## If valid then 'eResult' should equal 'Fields'.
                                                '   Unset result bits represent missing fields.

    If Len(Address) Then

        '## Save fields
        stAddress = msAddress
        stStreet = msStreet
        stCity = msCity
        stState = msState
        stPostcode = msPostcode
        stCountry = msCountry

        msAddress = Address
        pExtractAddress
        '## Calculate result code
        eResult = (Fields And eavfStreet) * Abs(Len(msStreet) > 0) + _
                  (Fields And eavfCity) * Abs(Len(msCity) > 0) + _
                  (Fields And eavfState) * Abs(Len(msState) > 0) + _
                  (Fields And eavfPostcode) * Abs(Len(msPostcode) > 0) + _
                  (Fields And eavfCountry) * Abs(Len(msCountry) > 0)
        ValidateAddress = eResult

        '## Restore fields
        msAddress = stAddress
        msStreet = stStreet
        msCity = stCity
        msState = stState
        msPostcode = stPostcode
        msCountry = stCountry

    End If

End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Private Sub Class_Initialize()

    ReDim msCountries(1 To 244)

    msCountries(1) = "Afghanistan"
    msCountries(2) = "Albania"
    msCountries(3) = "Algeria"
    msCountries(4) = "American Samoa"
    msCountries(5) = "Andorra"
    msCountries(6) = "Angola"
    msCountries(7) = "Anguilla"
    msCountries(8) = "Antigua and Barbuda"
    msCountries(9) = "Argentina"
    msCountries(10) = "Armenia"
    msCountries(11) = "Ashmore and Cartier Islands"
    msCountries(12) = "Australia"
    msCountries(13) = "Austria"
    msCountries(14) = "Azerbaijan"
    msCountries(15) = "Bahamas"
    msCountries(16) = "Bahrain"
    msCountries(17) = "Baker Island"
    msCountries(18) = "Bangladesh"
    msCountries(19) = "Barbados"
    msCountries(20) = "Belarus"
    msCountries(21) = "Belgium"
    msCountries(22) = "Belize"
    msCountries(23) = "Benin"
    msCountries(24) = "Bermuda"
    msCountries(25) = "Bhutan"
    msCountries(26) = "Bolivia"
    msCountries(27) = "Bosnia -Herzegovina"
    msCountries(28) = "Botswana"
    msCountries(29) = "Bouvet Island"
    msCountries(30) = "Brazil"
    msCountries(31) = "British Indian Ocean Territory"
    msCountries(32) = "Brunei Darussalam"
    msCountries(33) = "Bulgaria"
    msCountries(34) = "Burkina Faso"
    msCountries(35) = "Burundi"
    msCountries(36) = "Cambodia"
    msCountries(37) = "Cameroon"
    msCountries(38) = "Canada"
    msCountries(39) = "Cape Verde"
    msCountries(40) = "Cayman Islands"
    msCountries(41) = "Central African Republic"
    msCountries(42) = "Chad"
    msCountries(43) = "Chile"
    msCountries(44) = "China"
    msCountries(45) = "Christmas Island"
    msCountries(46) = "Cocos (Keeling) Islands"
    msCountries(47) = "Colombia"
    msCountries(48) = "Comoros"
    msCountries(49) = "Congo"
    msCountries(50) = "Cook Islands"
    msCountries(51) = "Coral Sea Islands"
    msCountries(52) = "Costa Rica"
    msCountries(53) = "Cote d 'Ivoire"
    msCountries(54) = "Croatia"
    msCountries(55) = "Cuba"
    msCountries(56) = "Cyprus"
    msCountries(57) = "Czech Republic"
    msCountries(58) = "Denmark"
    msCountries(59) = "Djibouti"
    msCountries(60) = "Dominica"
    msCountries(61) = "Dominican Republic"
    msCountries(62) = "Ecuador"
    msCountries(63) = "Egypt"
    msCountries(64) = "El Salvador"
    msCountries(65) = "Equatorial Guinea"
    msCountries(66) = "Eritrea"
    msCountries(67) = "Estonia"
    msCountries(68) = "Ethiopia"
    msCountries(69) = "F.Y.R.O.Macedonia"
    msCountries(70) = "Faeroe Islands"
    msCountries(71) = "Fiji Islands"
    msCountries(72) = "Finland"
    msCountries(73) = "France"
    msCountries(74) = "French Guiana"
    msCountries(75) = "French Polynesia"
    msCountries(76) = "French Southern and Antarctic Lands"
    msCountries(77) = "Gabon"
    msCountries(78) = "Gambia"
    msCountries(79) = "Georgia"
    msCountries(80) = "Germany"
    msCountries(81) = "Ghana"
    msCountries(82) = "Greece"
    msCountries(83) = "Greenland"
    msCountries(84) = "Grenada"
    msCountries(85) = "Guadeloupe"
    msCountries(86) = "Guam"
    msCountries(87) = "Guatemala"
    msCountries(88) = "Guernsey"
    msCountries(89) = "Guinea"
    msCountries(90) = "Guinea -Bissau"
    msCountries(91) = "Guyana"
    msCountries(92) = "Haiti"
    msCountries(93) = "Heard Island and McDonald Islands"
    msCountries(94) = "Honduras"
    msCountries(95) = "Hong Kong"
    msCountries(96) = "Howland Island"
    msCountries(97) = "Hungary"
    msCountries(98) = "Iceland"
    msCountries(99) = "India"
    msCountries(100) = "Indonesia"
    msCountries(101) = "Iran"
    msCountries(102) = "Iraq"
    msCountries(103) = "Ireland"
    msCountries(104) = "Israel"
    msCountries(105) = "Italy"
    msCountries(106) = "Jamaica"
    msCountries(107) = "Jan Mayen"
    msCountries(108) = "Japan"
    msCountries(109) = "Jarvis Island"
    msCountries(110) = "Jersey"
    msCountries(111) = "Johnston Atoll"
    msCountries(112) = "Jordan"
    msCountries(113) = "Kazakhstan"
    msCountries(114) = "Kenya"
    msCountries(115) = "Kingman Reef"
    msCountries(116) = "Kiribati"
    msCountries(117) = "Kuwait"
    msCountries(118) = "Kyrgyzstan"
    msCountries(119) = "Laos"
    msCountries(120) = "Latvia"
    msCountries(121) = "Lebanon"
    msCountries(122) = "Lesotho"
    msCountries(123) = "Liberia"
    msCountries(124) = "Libya"
    msCountries(125) = "Liechtenstein"
    msCountries(126) = "Lithuania"
    msCountries(127) = "Luxembourg"
    msCountries(128) = "Macau"
    msCountries(129) = "Madagascar"
    msCountries(130) = "Malawi"
    msCountries(131) = "Malaysia"
    msCountries(132) = "Maldives"
    msCountries(133) = "Mali"
    msCountries(134) = "Malta"
    msCountries(135) = "Man (Isle of)"
    msCountries(136) = "Marshall Islands"
    msCountries(137) = "Martinique"
    msCountries(138) = "Mauritania"
    msCountries(139) = "Mauritius"
    msCountries(140) = "Mexico"
    msCountries(141) = "Micronesia (Federated States of)"
    msCountries(142) = "Midway Islands"
    msCountries(143) = "Moldova"
    msCountries(144) = "Monaco"
    msCountries(145) = "Mongolia"
    msCountries(146) = "Montserrat"
    msCountries(147) = "Morocco"
    msCountries(148) = "Mozambique"
    msCountries(149) = "Myanmar"
    msCountries(150) = "Namibia"
    msCountries(151) = "Nauru"
    msCountries(152) = "Navassa Island"
    msCountries(153) = "Nepal"
    msCountries(154) = "Netherlands"
    msCountries(155) = "Netherlands Antilles"
    msCountries(156) = "New Caledonia"
    msCountries(157) = "New Zealand"
    msCountries(158) = "Nicaragua"
    msCountries(159) = "Niger"
    msCountries(160) = "Nigeria"
    msCountries(161) = "Niue"
    msCountries(162) = "Norfolk Island"
    msCountries(163) = "North Korea"
    msCountries(164) = "Northern Mariana Islands"
    msCountries(165) = "Norway"
    msCountries(166) = "Oman"
    msCountries(167) = "Pakistan"
    msCountries(168) = "Palau"
    msCountries(169) = "Palmyra Atoll"
    msCountries(170) = "Panama"
    msCountries(171) = "Papua New Guinea"
    msCountries(172) = "Paraguay"
    msCountries(173) = "Peru"
    msCountries(174) = "Philippines"
    msCountries(175) = "Pitcairn Island"
    msCountries(176) = "Poland"
    msCountries(177) = "Portugal"
    msCountries(178) = "Puerto Rico"
    msCountries(179) = "Qatar"
    msCountries(180) = "Réunion"
    msCountries(181) = "Romania"
    msCountries(182) = "Russia"
    msCountries(183) = "Rwanda"
    msCountries(184) = "Saint Helena"
    msCountries(185) = "Saint Pierre And Miquelon"
    msCountries(186) = "Samoa"
    msCountries(187) = "San Marino"
    msCountries(188) = "Sao Tome And Principe"
    msCountries(189) = "Saudi Arabia"
    msCountries(190) = "Senegal"
    msCountries(191) = "Seychelles"
    msCountries(192) = "Sierra Leone"
    msCountries(193) = "Singapore"
    msCountries(194) = "Slovak Republic"
    msCountries(195) = "Slovenia"
    msCountries(196) = "Solomon Islands"
    msCountries(197) = "Somalia"
    msCountries(198) = "South Africa"
    msCountries(199) = "South Georgia and the South Sandwich Islands"
    msCountries(200) = "South Korea"
    msCountries(201) = "Spain"
    msCountries(202) = "Sri Lanka"
    msCountries(203) = "St. Kitts and Nevis"
    msCountries(204) = "St.Lucia"
    msCountries(205) = "St. Vincent and the Grenadines"
    msCountries(206) = "Sudan"
    msCountries(207) = "Surinam"
    msCountries(208) = "Svalbard"
    msCountries(209) = "Swaziland"
    msCountries(210) = "Sweden"
    msCountries(211) = "Switzerland"
    msCountries(212) = "Syria"
    msCountries(213) = "Taiwan"
    msCountries(214) = "Tajikistan"
    msCountries(215) = "Tanzania"
    msCountries(216) = "Thailand"
    msCountries(217) = "Togo"
    msCountries(218) = "Tokelau"
    msCountries(219) = "Tonga"
    msCountries(220) = "Trinidad and Tobago"
    msCountries(221) = "Tunisia"
    msCountries(222) = "Turkey"
    msCountries(223) = "Turkmenistan"
    msCountries(224) = "Turks and Caicos Islands"
    msCountries(225) = "Tuvalu"
    msCountries(226) = "Uganda"
    msCountries(227) = "Ukraine"
    msCountries(228) = "United Arab Emirates"
    msCountries(229) = "United Kingdom"
    msCountries(230) = "United States of America"
    msCountries(231) = "Uruguay"
    msCountries(232) = "Uzbekistan"
    msCountries(233) = "Vanuatu"
    msCountries(234) = "Vatican City"
    msCountries(235) = "Venezuela"
    msCountries(236) = "Vietnam"
    msCountries(237) = "Virgin Islands(British)"
    msCountries(238) = "Virgin Islands (U.S.)"
    msCountries(239) = "Wake Island"
    msCountries(240) = "Wallis and Futuna"
    msCountries(241) = "Yemen"
    msCountries(242) = "Yugoslavia"
    msCountries(243) = "Zambia"
    msCountries(244) = "Zimbabwe"

End Sub

Private Sub Class_Terminate()
'
End Sub


'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Private Function pBuildAddress() As String

    Dim sFull As String

    sFull = Trim$(IIf(Len(msStreet), msStreet + vbCrLf, "") + _
                  IIf(Len(msCity), msCity + ", ", "") + _
                  IIf(Len(msState), msState + " ", "") + _
                  IIf(Len(msPostcode), msPostcode, ""))
    sFull = Trim$(sFull + IIf(Len(sFull), vbCrLf, "") + _
                          IIf(Len(msCountry), msCountry, ""))
    pBuildAddress = sFull

End Function

Private Sub pExtractAddress()

    Dim lCount    As Long
    Dim lMax      As Long
    Dim lLoop     As Long
    Dim sLines()  As String
    Dim lPtr      As Long
    '
    '## Clear Fields
    '
    msStreet = ""
    msCity = ""
    msState = ""
    msPostcode = ""
    msCountry = ""

    sLines = Split(msAddress, vbCrLf)
    lCount = UBound(sLines)

    '## Remove blank/empty lines
    Do
        If sLines(lPtr) = "" Then
            lCount = lCount - 1
            If lPtr <= lCount Then
                For lLoop = lPtr To lCount
                    sLines(lLoop) = sLines(lLoop + 1)
                Next
            End If
        Else
            '## done with do/loop to handle a series of blank lines
            lPtr = lPtr + 1
        End If
    Loop Until lPtr > lCount
    lMax = lCount

    '## Check for Country
    For lLoop = 1 To UBound(msCountries)
        If UCase$(sLines(lCount)) = UCase$(msCountries(lLoop)) Then
            If mbAutoCorrect Then
                msCountry = msCountries(lLoop)
            Else
                msCountry = sLines(lCount)
            End If
            lCount = lCount - 1
            lMax = lCount
            Exit For
        End If
    Next

    '## Clean up rest of the address details
    If mbAutoCorrect Then
        For lLoop = 0 To lMax
            sLines(lLoop) = pProperCasing(pStripRightComa(sLines(lLoop)))
        Next
    End If

    '## Break out remaining address tokens
    Select Case lCount
        Case 0                          '##
           pGetDetails sLines(lCount)
        Case Is > 0                     '##
           pGetDetails sLines(lCount)
            For lLoop = 0 To lMax - 1
                msStreet = msStreet + vbCrLf + sLines(lLoop)
            Next
            msStreet = Trim$(Mid$(msStreet, 3))
    End Select

End Sub

'
'## Extract City, State, Postcode fields
'
Private Sub pGetDetails(AddressLine As String)

    Dim sTokens() As String
    Dim lCount    As Long
    Dim lLoop     As Long
    Dim lMin      As Long
    Dim bStop     As Boolean

    sTokens = Split(AddressLine, " ")
    lCount = UBound(sTokens)
    '## Extract the city and remove unwanted commas
    For lLoop = 0 To lCount
        If mbAutoCorrect Then
            sTokens(lLoop) = pProperCase(sTokens(lLoop))
        End If
        If InStr(sTokens(lLoop), ",") Then
            sTokens(lLoop) = Left$(sTokens(lLoop), Len(sTokens(lLoop)) - 1)
            If Not bStop Then
                lMin = lMin + 1
                msCity = msCity + " " + sTokens(lLoop)
                bStop = True
            End If
        Else
            If Not bStop Then
                msCity = msCity + " " + sTokens(lLoop)
                lMin = lMin + 1
            End If
        End If
    Next
    msCity = Trim$(msCity)

    '## Now extract State & Postcode
    If lMin < lCount + 1 Then
        Select Case lCount - lMin
            Case Is > 0
                msPostcode = sTokens(lCount)
                For lLoop = lMin To lCount - 1
                    msState = msState + " " + sTokens(lLoop)
                Next
                msState = Trim$(msState)
                
            Case 0
                msState = sTokens(lMin)
        End Select
    End If

End Sub

'
'## Correctly capitalises a line of text
'
Private Function pProperCasing(ByVal Words As String) As String

    Dim lCount    As Long
    Dim lLoop     As Long
    Dim sTokens() As String
    Dim sTmp      As String

    sTokens = Split(Words, " ")
    lCount = UBound(sTokens)
    For lLoop = 0 To lCount
        sTmp = sTmp + " " + pProperCase(sTokens(lLoop))
    Next
    pProperCasing = Trim$(sTmp)

End Function
'
'## First character capitalised, remaining lowercase except hypen'd words, 'Mc' & 'Mac' for a word
'
Private Function pProperCase(ByVal Word As String) As String

    Dim sResult As String
    Dim lSep    As Long
    Dim lLen    As Long
    Dim lPos    As Long

    lLen = Len(Word)
    sResult = UCase$(Left$(Word, 1)) + LCase$(Mid$(Word, 2))
    lSep = InStr(Word, "-")
    If lSep And lSep < lLen Then Mid$(sResult, lSep + 1) = UCase$(Mid$(sResult, lSep + 1, 1))

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "'")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 1, 1) = UCase$(Mid$(sResult, lSep + 1, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "MC")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 2, 1) = UCase$(Mid$(sResult, lSep + 2, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "MAC")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 3, 1) = UCase$(Mid$(sResult, lSep + 3, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    pProperCase = sResult

End Function

'
'## Checks and removes righthand commas
'
Private Function pStripRightComa(Text As String) As String
    If Right$(Text, 1) = "," Then
        pStripRightComa = Left$(Text, Len(Text) - 1)
    Else
        pStripRightComa = Text
    End If
End Function
