Attribute VB_Name = "LocaleHelper"

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Private Const LOCALE_ILANGUAGE = &H1        'Language id
Private Const LOCALE_SLANGUAGE = &H2        'Localized name of language
Private Const LOCALE_SENGLANGUAGE = &H1001  'English name of language

Private Const LOCALE_ICOUNTRY = &H5         'Country code
Private Const LOCALE_SCOUNTRY = &H6         'Name
Private Const LOCALE_SENGCOUNTRY = &H1002   'English name of country

Private Const LOCALE_SDECIMAL = &HE         'Decimal separator
Private Const LOCALE_STHOUSAND = &HF        'Thousand separator

Private Const MAX_BUF As Long = 260


'get country id
Function GetCountryId() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_ICOUNTRY, sBuf, MAX_BUF)
    GetCountryId = Left$(sBuf, lRet - 1)
End Function


'get country name
Function GetCountryName() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SCOUNTRY, sBuf, MAX_BUF)
    GetCountryName = Left$(sBuf, lRet - 1)
End Function


'get country name International
Function GetCountryNameInt() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SENGCOUNTRY, sBuf, MAX_BUF)
    GetCountryNameInt = Left$(sBuf, lRet - 1)
End Function


'get language id
Function GetLocaleLanguageId() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_ILANGUAGE, sBuf, MAX_BUF)
    GetLocaleLanguageId = Left$(sBuf, lRet - 1)
End Function


'get localized language
Function GetLocaleLanguage() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SLANGUAGE, sBuf, MAX_BUF)
    GetLocaleLanguage = Left$(sBuf, lRet - 1)
End Function


'get language International
Function GetLocaleLanguageInt() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SENGLANGUAGE, sBuf, MAX_BUF)
    GetLocaleLanguageInt = Left$(sBuf, lRet - 1)
End Function

'get decimal separator
Function GetDecimalSeparator() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SDECIMAL, sBuf, MAX_BUF)
    GetDecimalSeparator = Left$(sBuf, lRet - 1)
End Function

'get language culture code
Function GetLocaleCulture() As String

' 26.02.2023
' https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a

	Select Case GetUserDefaultLCID

		Case 1026: GetLocaleCulture = "bg-BG"  ' Bulgarian
			
		Case 1027: GetLocaleCulture = "ca-ES"  ' Catalan

		Case 1028: GetLocaleCulture = "zh-TW"  ' Chinese - Taiwan

		Case 1029: GetLocaleCulture = "cs-CZ"  ' Czech

		Case 1030: GetLocaleCulture = "da-DK"  ' Danish

		Case 1031: GetLocaleCulture = "de-DE"  ' German - Germany

		Case 1032: GetLocaleCulture = "el-GR"  ' Greek

		Case 1033: GetLocaleCulture = "en-US"  ' English - United States

		Case 1034: GetLocaleCulture = "es-ES"  ' Spanish - Spain (Traditional Sort)

		Case 1035: GetLocaleCulture = "fi-FI"  ' Finnish

		Case 1036: GetLocaleCulture = "fr-FR"  ' French - France

		Case 1037: GetLocaleCulture = "he-IL"  ' Hebrew

		Case 1038: GetLocaleCulture = "hu-HU"  ' Hungarian

		Case 1039: GetLocaleCulture = "is-IS"  ' Icelandic

		Case 1040: GetLocaleCulture = "it-IT"  ' Italian - Italy

		Case 1041: GetLocaleCulture = "ja-JP"  ' Japanese

		Case 1042: GetLocaleCulture = "ko-KR"  ' Korean

		Case 1043: GetLocaleCulture = "nl-NL"  ' Dutch - Netherlands

		Case 1044: GetLocaleCulture = "nb-NO"  ' Norwegian (Bokmål)

		Case 1045: GetLocaleCulture = "pl-PL"  ' Polish

		Case 1046: GetLocaleCulture = "pt-BR"  ' Portuguese - Brazil

		Case 1047: GetLocaleCulture = "rm-CH"  ' Rhaeto-Romanic

		Case 1048: GetLocaleCulture = "ro-RO"  ' Romanian

		Case 1049: GetLocaleCulture = "ru-RU"  ' Russian

		Case 1050: GetLocaleCulture = "hr-HR"  ' Croatian

		Case 1051: GetLocaleCulture = "sk-SK"  ' Slovak

		Case 1052: GetLocaleCulture = "sq-AL"  ' Albanian - Albania

		Case 1053: GetLocaleCulture = "sv-SE"  ' Swedish

		Case 1054: GetLocaleCulture = "th-TH"  ' Thai

		Case 1055: GetLocaleCulture = "tr-TR"  ' Turkish

		Case 1056: GetLocaleCulture = "ur-PK"  ' Urdu - Pakistan

		Case 1057: GetLocaleCulture = "id-ID"  ' Indonesian

		Case 1058: GetLocaleCulture = "uk-UA"  ' Ukrainian

		Case 1059: GetLocaleCulture = "be-BY"  ' Belarusian

		Case 1060: GetLocaleCulture = "sl-SI"  ' Slovenian

		Case 1061: GetLocaleCulture = "et-EE"  ' Estonian

		Case 1062: GetLocaleCulture = "lv-LV"  ' Latvian

		Case 1063: GetLocaleCulture = "lt-LT"  ' Lithuanian

		Case 1064: GetLocaleCulture = "tg-Cyrl-TJ"  ' Tajik

		Case 1065: GetLocaleCulture = "fa-IR"  ' Persian

		Case 1066: GetLocaleCulture = "vi-VN"  ' Vietnamese

		Case 1067: GetLocaleCulture = "hy-AM"  ' Armenian - Armenia

		Case 1068: GetLocaleCulture = "az-Latn-AZ"  ' Azeri (Latin)

		Case 1069: GetLocaleCulture = "eu-ES"  ' Basque

		Case 1070: GetLocaleCulture = "wen-DE"  ' Sorbian

		Case 1071: GetLocaleCulture = "mk-MK"  ' F.Y.R.O. Macedonian

		Case 1072: GetLocaleCulture = "st-ZA"  ' Sutu

		Case 1073: GetLocaleCulture = "ts-ZA"  ' Tsonga

		Case 1074: GetLocaleCulture = "tn-ZA"  ' Tswana

		Case 1075: GetLocaleCulture = "ven-ZA"  ' Venda

		Case 1076: GetLocaleCulture = "xh-ZA"  ' Xhosa

		Case 1077: GetLocaleCulture = "zu-ZA"  ' Zulu

		Case 1078: GetLocaleCulture = "af-ZA"  ' Afrikaans - South Africa

		Case 1079: GetLocaleCulture = "ka-GE"  ' Georgian

		Case 1080: GetLocaleCulture = "fo-FO"  ' Faroese

		Case 1081: GetLocaleCulture = "hi-IN"  ' Hindi

		Case 1082: GetLocaleCulture = "mt-MT"  ' Maltese

		Case 1083: GetLocaleCulture = "se-NO"  ' Sami

		Case 1084: GetLocaleCulture = "gd-GB"  ' Gaelic (Scotland)

		Case 1085: GetLocaleCulture = "yi"  ' Yiddish

		Case 1086: GetLocaleCulture = "ms-MY"  ' Malay - Malaysia

		Case 1087: GetLocaleCulture = "kk-KZ"  ' Kazakh

		Case 1088: GetLocaleCulture = "ky-KG"  ' Kyrgyz (Cyrillic)

		Case 1089: GetLocaleCulture = "sw-KE"  ' Swahili

		Case 1090: GetLocaleCulture = "tk-TM"  ' Turkmen

		Case 1091: GetLocaleCulture = "uz-Latn-UZ"  ' Uzbek (Latin)

		Case 1092: GetLocaleCulture = "tt-RU"  ' Tatar

		Case 1093: GetLocaleCulture = "bn-IN"  ' Bengali (India)

		Case 1094: GetLocaleCulture = "pa-IN"  ' Punjabi

		Case 1095: GetLocaleCulture = "gu-IN"  ' Gujarati

		Case 1096: GetLocaleCulture = "or-IN"  ' Oriya

		Case 1097: GetLocaleCulture = "ta-IN"  ' Tamil

		Case 1098: GetLocaleCulture = "te-IN"  ' Telugu

		Case 1099: GetLocaleCulture = "kn-IN"  ' Kannada

		Case 1100: GetLocaleCulture = "ml-IN"  ' Malayalam

		Case 1101: GetLocaleCulture = "as-IN"  ' Assamese

		Case 1102: GetLocaleCulture = "mr-IN"  ' Marathi

		Case 1103: GetLocaleCulture = "sa-IN"  ' Sanskrit

		Case 1104: GetLocaleCulture = "mn-MN"  ' Mongolian (Cyrillic)

		Case 1105: GetLocaleCulture = "bo-CN"  ' Tibetan - People's Republic of China

		Case 1106: GetLocaleCulture = "cy-GB"  ' Welsh

		Case 1107: GetLocaleCulture = "km-KH"  ' Khmer

		Case 1108: GetLocaleCulture = "lo-LA"  ' Lao

		Case 1109: GetLocaleCulture = "my-MM"  ' Burmese

		Case 1110: GetLocaleCulture = "gl-ES"  ' Galician

		Case 1111: GetLocaleCulture = "kok-IN"  ' Konkani

		Case 1112: GetLocaleCulture = "mni"  ' Manipuri

		Case 1113: GetLocaleCulture = "sd-IN"  ' Sindhi - India

		Case 1114: GetLocaleCulture = "syr-SY"  ' Syriac

		Case 1115: GetLocaleCulture = "si-LK"  ' Sinhalese - Sri Lanka

		Case 1116: GetLocaleCulture = "chr-US"  ' Cherokee - United States

		Case 1117: GetLocaleCulture = "iu-Cans-CA"  ' Inuktitut

		Case 1118: GetLocaleCulture = "am-ET"  ' Amharic - Ethiopia

		Case 1119: GetLocaleCulture = "tmz"  ' Tamazight (Arabic)

		Case 1120: GetLocaleCulture = "ks-Arab-IN"  ' Kashmiri (Arabic)

		Case 1121: GetLocaleCulture = "ne-NP"  ' Nepali

		Case 1122: GetLocaleCulture = "fy-NL"  ' Frisian - Netherlands

		Case 1123: GetLocaleCulture = "ps-AF"  ' Pashto

		Case 1124: GetLocaleCulture = "fil-PH"  ' Filipino

		Case 1125: GetLocaleCulture = "dv-MV"  ' Divehi

		Case 1126: GetLocaleCulture = "bin-NG"  ' Edo

		Case 1127: GetLocaleCulture = "fuv-NG"  ' Fulfulde - Nigeria

		Case 1128: GetLocaleCulture = "ha-Latn-NG"  ' Hausa - Nigeria

		Case 1129: GetLocaleCulture = "ibb-NG"  ' Ibibio - Nigeria

		Case 1130: GetLocaleCulture = "yo-NG"  ' Yoruba

		Case 1131: GetLocaleCulture = "quz-BO"  ' Quecha - Bolivia

		Case 1132: GetLocaleCulture = "nso-ZA"  ' Sepedi

		Case 1136: GetLocaleCulture = "ig-NG"  ' Igbo - Nigeria

		Case 1137: GetLocaleCulture = "kr-NG"  ' Kanuri - Nigeria

		Case 1138: GetLocaleCulture = "gaz-ET"  ' Oromo

		Case 1139: GetLocaleCulture = "ti-ER"  ' Tigrigna - Ethiopia

		Case 1140: GetLocaleCulture = "gn-PY"  ' Guarani - Paraguay

		Case 1141: GetLocaleCulture = "haw-US"  ' Hawaiian - United States

		Case 1142: GetLocaleCulture = "la"  ' Latin

		Case 1143: GetLocaleCulture = "so-SO"  ' Somali

		Case 1144: GetLocaleCulture = "ii-CN"  ' Yi

		Case 1145: GetLocaleCulture = "pap-AN"  ' Papiamentu

		Case 1152: GetLocaleCulture = "ug-Arab-CN"  ' Uighur - China

		Case 1153: GetLocaleCulture = "mi-NZ"  ' Maori - New Zealand

		Case 2049: GetLocaleCulture = "ar-IQ"  ' Arabic - Iraq

		Case 2052: GetLocaleCulture = "zh-CN"  ' Chinese - People's Republic of China

		Case 2055: GetLocaleCulture = "de-CH"  ' German - Switzerland

		Case 2057: GetLocaleCulture = "en-GB"  ' English - United Kingdom

		Case 2058: GetLocaleCulture = "es-MX"  ' Spanish - Mexico

		Case 2060: GetLocaleCulture = "fr-BE"  ' French - Belgium

		Case 2064: GetLocaleCulture = "it-CH"  ' Italian - Switzerland

		Case 2067: GetLocaleCulture = "nl-BE"  ' Dutch - Belgium

		Case 2068: GetLocaleCulture = "nn-NO"  ' Norwegian (Nynorsk)

		Case 2070: GetLocaleCulture = "pt-PT"  ' Portuguese - Portugal

		Case 2072: GetLocaleCulture = "ro-MD"  ' Romanian - Moldava

		Case 2073: GetLocaleCulture = "ru-MD"  ' Russian - Moldava

		Case 2074: GetLocaleCulture = "sr-Latn-CS"  ' Serbian (Latin)

		Case 2077: GetLocaleCulture = "sv-FI"  ' Swedish - Finland

		Case 2080: GetLocaleCulture = "ur-IN"  ' Urdu - India

		Case 2092: GetLocaleCulture = "az-Cyrl-AZ"  ' Azeri (Cyrillic)

		Case 2108: GetLocaleCulture = "ga-IE"  ' Gaelic (Ireland)

		Case 2110: GetLocaleCulture = "ms-BN"  ' Malay - Brunei Darussalam

		Case 2115: GetLocaleCulture = "uz-Cyrl-UZ"  ' Uzbek (Cyrillic)

		Case 2117: GetLocaleCulture = "bn-BD"  ' Bengali (Bangladesh)

		Case 2118: GetLocaleCulture = "pa-PK"  ' Punjabi (Pakistan)

		Case 2128: GetLocaleCulture = "mn-Mong-CN"  ' Mongolian (Mongolian)

		Case 2129: GetLocaleCulture = "bo-BT"  ' Tibetan - Bhutan

		Case 2137: GetLocaleCulture = "sd-PK"  ' Sindhi - Pakistan

		Case 2143: GetLocaleCulture = "tzm-Latn-DZ"  ' Tamazight (Latin)

		Case 2144: GetLocaleCulture = "ks-Deva-IN"  ' Kashmiri (Devanagari)

		Case 2145: GetLocaleCulture = "ne-IN"  ' Nepali - India

		Case 2155: GetLocaleCulture = "quz-EC"  ' Quecha - Ecuador

		Case 2163: GetLocaleCulture = "ti-ET"  ' Tigrigna - Eritrea

		Case 3073: GetLocaleCulture = "ar-EG"  ' Arabic - Egypt

		Case 3076: GetLocaleCulture = "zh-HK"  ' Chinese - Hong Kong SAR

		Case 3079: GetLocaleCulture = "de-AT"  ' German - Austria

		Case 3081: GetLocaleCulture = "en-AU"  ' English - Australia

		Case 3082: GetLocaleCulture = "es-ES"  ' Spanish - Spain (Modern Sort)

		Case 3084: GetLocaleCulture = "fr-CA"  ' French - Canada

		Case 3098: GetLocaleCulture = "sr-Cyrl-CS"  ' Serbian (Cyrillic)

		Case 3179: GetLocaleCulture = "quz-PE"  ' Quecha - Peru

		Case 4097: GetLocaleCulture = "ar-LY"  ' Arabic - Libya

		Case 4100: GetLocaleCulture = "zh-SG"  ' Chinese - Singapore

		Case 4103: GetLocaleCulture = "de-LU"  ' German - Luxembourg

		Case 4105: GetLocaleCulture = "en-CA"  ' English - Canada

		Case 4106: GetLocaleCulture = "es-GT"  ' Spanish - Guatemala

		Case 4108: GetLocaleCulture = "fr-CH"  ' French - Switzerland

		Case 4122: GetLocaleCulture = "hr-BA"  ' Croatian (Bosnia/Herzegovina)

		Case 5121: GetLocaleCulture = "ar-DZ"  ' Arabic - Algeria

		Case 5124: GetLocaleCulture = "zh-MO"  ' Chinese - Macao SAR

		Case 5127: GetLocaleCulture = "de-LI"  ' German - Liechtenstein

		Case 5129: GetLocaleCulture = "en-NZ"  ' English - New Zealand

		Case 5130: GetLocaleCulture = "es-CR"  ' Spanish - Costa Rica

		Case 5132: GetLocaleCulture = "fr-LU"  ' French - Luxembourg

		Case 5146: GetLocaleCulture = "bs-Latn-BA"  ' Bosnian (Bosnia/Herzegovina)

		Case 6145: GetLocaleCulture = "ar-MO"  ' Arabic - Morocco

		Case 6153: GetLocaleCulture = "en-IE"  ' English - Ireland

		Case 6154: GetLocaleCulture = "es-PA"  ' Spanish - Panama

		Case 6156: GetLocaleCulture = "fr-MC"  ' French - Monaco

		Case 7169: GetLocaleCulture = "ar-TN"  ' Arabic - Tunisia

		Case 7177: GetLocaleCulture = "en-ZA"  ' English - South Africa

		Case 7178: GetLocaleCulture = "es-DO"  ' Spanish - Dominican Republic

		Case 7180: GetLocaleCulture = "fr-029"  ' French - West Indies

		Case 8193: GetLocaleCulture = "ar-OM"  ' Arabic - Oman

		Case 8201: GetLocaleCulture = "en-JM"  ' English - Jamaica

		Case 8202: GetLocaleCulture = "es-VE"  ' Spanish - Venezuela

		Case 8204: GetLocaleCulture = "fr-RE"  ' French - Reunion

		Case 9217: GetLocaleCulture = "ar-YE"  ' Arabic - Yemen

		Case 9225: GetLocaleCulture = "en-029"  ' English - Caribbean

		Case 9226: GetLocaleCulture = "es-CO"  ' Spanish - Colombia

		Case 9228: GetLocaleCulture = "fr-CG"  ' French - Democratic Rep. of Congo

		Case 10241: GetLocaleCulture = "ar-SY"  ' Arabic - Syria

		Case 10249: GetLocaleCulture = "en-BZ"  ' English - Belize

		Case 10250: GetLocaleCulture = "es-PE"  ' Spanish - Peru

		Case 10252: GetLocaleCulture = "fr-SN"  ' French - Senegal

		Case 11265: GetLocaleCulture = "ar-JO"  ' Arabic - Jordan

		Case 11273: GetLocaleCulture = "en-TT"  ' English - Trinidad

		Case 11274: GetLocaleCulture = "es-AR"  ' Spanish - Argentina

		Case 11276: GetLocaleCulture = "fr-CM"  ' French - Cameroon

		Case 12289: GetLocaleCulture = "ar-LB"  ' Arabic - Lebanon

		Case 12297: GetLocaleCulture = "en-ZW"  ' English - Zimbabwe

		Case 12298: GetLocaleCulture = "es-EC"  ' Spanish - Ecuador

		Case 12300: GetLocaleCulture = "fr-CI"  ' French - Cote d'Ivoire

		Case 13313: GetLocaleCulture = "ar-KW"  ' Arabic - Kuwait

		Case 13321: GetLocaleCulture = "en-PH"  ' English - Philippines

		Case 13322: GetLocaleCulture = "es-CL"  ' Spanish - Chile

		Case 13324: GetLocaleCulture = "fr-ML"  ' French - Mali

		Case 14337: GetLocaleCulture = "ar-AE"  ' Arabic - U.A.E.

		Case 14345: GetLocaleCulture = "en-ID"  ' English - Indonesia

		Case 14346: GetLocaleCulture = "es-UY"  ' Spanish - Uruguay

		Case 14348: GetLocaleCulture = "fr-MA"  ' French - Morocco

		Case 15361: GetLocaleCulture = "ar-BH"  ' Arabic - Bahrain

		Case 15369: GetLocaleCulture = "en-HK"  ' English - Hong Kong SAR

		Case 15370: GetLocaleCulture = "es-PY"  ' Spanish - Paraguay

		Case 15372: GetLocaleCulture = "fr-HT"  ' French - Haiti

		Case 16385: GetLocaleCulture = "ar-QA"  ' Arabic - Qatar

		Case 16393: GetLocaleCulture = "en-IN"  ' English - India

		Case 16394: GetLocaleCulture = "es-BO"  ' Spanish - Bolivia

		Case 17417: GetLocaleCulture = "en-MY"  ' English - Malaysia

		Case 17418: GetLocaleCulture = "es-SV"  ' Spanish - El Salvador

		Case 18441: GetLocaleCulture = "en-SG"  ' English - Singapore

		Case 18442: GetLocaleCulture = "es-HN"  ' Spanish - Honduras

		Case 19466: GetLocaleCulture = "es-NI"  ' Spanish - Nicaragua

		Case 20490: GetLocaleCulture = "es-PR"  ' Spanish - Puerto Rico

		Case 21514: GetLocaleCulture = "es-US"  ' Spanish - United States

		Case 58378: GetLocaleCulture = "es-419"  ' Spanish - Latin America

		Case 58380: GetLocaleCulture = "fr-015"  ' French - North Africa

	End Select

End Function