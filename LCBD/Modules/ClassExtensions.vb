
'   file:           classExtensions.vb
'   data:           29/09/2011
'   autore:         Simone Scigliuzzi
'   classe:         
'   descrizione:    modulo che definisce i metodi estesi per la classe String

Imports System.Runtime.CompilerServices
Imports System.Runtime.CompilerServices.ExtensionAttribute
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.IO.Stream
Imports System.Web
Imports System.Web.Script.Serialization                 ' requires System.Web.Extensions references, and .net 4.0 or higher

Public Module classExtensions

    ' estensione metodi classe String

#Region "String Extensions"
    ' String Extensions

    ''' <summary>
    ''' Converts a JSON list into a List
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>requries System.Web.Script.Serialization (add references to System.Web.Extensions)</remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function FromJSON2List(ByVal value As String) As List(Of String)

        Dim lstRetVal As List(Of String) = New List(Of String)
        value = value.Trim()

        Dim jss As New JavaScriptSerializer()
        Try
            lstRetVal = jss.Deserialize(Of List(Of String))(value)
        Catch ex As Exception
        End Try

        Return lstRetVal
    End Function

    ''' <summary>
    ''' Converts an array of strings a List of strings
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>requries System.Web.Script.Serialization (add references to System.Web.Extensions)</remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Array2List(ByVal value() As String) As List(Of String)

        Dim lstRetVal As List(Of String) = New List(Of String)
        lstRetVal.AddRange(value)

        Return lstRetVal
    End Function

    ''' <summary>
    ''' Returns the default string if current string is empty
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Def(ByVal value As String, ByVal DefaultValue As String) As Boolean
        If value.Trim() = "" Then
            Return DefaultValue
        Else
            Return value
        End If
    End Function

    ''' <summary>
    ''' Encodes the string into Base64
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToBase64(ByVal value As String) As String
        Dim arrBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(value)
        Return Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes(value))
    End Function

    ''' <summary>
    ''' Encodes the string into Base64
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function FromBase64(ByVal value As String) As String
        Dim arrBytes() As Byte = System.Convert.FromBase64String(value)
        Return System.Text.ASCIIEncoding.ASCII.GetString(arrBytes)
    End Function

    ''' <summary>
    ''' Turn text into "flat" text
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Flat(ByVal Value As String) As String

        Value = Value.Trim()
        Value = Value.Replace(Chr(9), " ")
        Value = Value.Replace(Chr(34), " ")
        Value = Value.RemoveDuplicatedContiguosChar(" ")
        Value = Value.RemoveDuplicatedContiguosChar(Chr(10))
        Value = Value.RemoveDuplicatedContiguosChar(Chr(13))
        Value = Value.RemoveDuplicatedContiguosChar(vbCrLf)

        ' Removes CRLF at the start and the end of the text
        While Strings.Left(Value, 1) = Chr(10) Or Strings.Left(Value, 1) = Chr(13)
            Value = Strings.Mid(Value, 2).Trim()
        End While
        While Strings.Right(Value, 1) = Chr(10) Or Strings.Right(Value, 1) = Chr(13) Or Strings.Right(Value, 1) = ","
            Value = Strings.Left(Value, Value.Length - 1).Trim()
        End While

        ' Adjusts lines
        Dim arrInLines() As String = Split(Value, vbCrLf)
        Value = ""
        For Each sLine As String In arrInLines
            sLine = sLine.Trim()
            If sLine.Length > 0 Then
                Dim sLC As String = Strings.Right(sLine, 1)             ' makes sure there's a DOT at the end of the line
                Select Case sLC
                    Case ".", "?", "!", ",", ";", ":"
                    Case Else
                        sLine &= "."
                End Select
            End If
            If Value <> "" Then Value &= vbCrLf
            Value &= sLine
        Next

        Return Value
    End Function

    ''' <summary>
    ''' Returns the number of words
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function CountWords(ByVal value As String) As Int32
        value = value.Trim()
        If value = "" Then Return 0
        Dim collection As MatchCollection = Regex.Matches(value, "\S+")
        Return collection.Count
    End Function

    ''' <summary>
    ''' Returns the number of characters excluding tabs and CRLF
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function CountChars(ByVal value As String) As Int32
        value = value.Replace(" ", "").Replace(Chr(9), " ").Replace(Chr(10), " ").Replace(Chr(13), " ").RemoveDuplicatedContiguosChar(" ").Trim()
        If value = "" Then Return 0
        Return value.Length
    End Function

    ''' <summary>
    ''' Returns true if the string represents a valid uri
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function IsValidMailAddress(ByRef value As String) As Boolean
        ' Normalizes arguments
        value = value.Trim().ToLower()
        If value = "" Then Return False

        Dim objRegex As Regex = New Global.System.Text.RegularExpressions.Regex(CType("^[a-z0-9._-]+\@[a-z0-9._-]+\.[a-z0-9]{2,5}$", System.String))
        Dim objMatch As Match = objRegex.Match(value)

        ' return
        Return objMatch.Success
    End Function

    ''' <summary>
    ''' Returns true if the string represents a valid uri
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function IsValidURI(ByRef value As String) As Boolean
        Return System.Uri.IsWellFormedUriString(value, UriKind.Absolute)
    End Function

    ''' <summary>
    ''' Encode url
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function UrlEncode(ByVal value As String) As String
        ' requires System.Web
        Dim sRetVal As String = HttpUtility.UrlEncode(value.Trim())
        sRetVal = sRetVal.Replace("+", "%20")
        Return sRetVal
    End Function

    ''' <summary>
    ''' Decode url
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function UrlDecode(ByVal value As String) As String
        ' requires System.Web
        Return HttpUtility.UrlDecode(value)
    End Function

    ''' <summary>
    ''' Converts current string into a string, also if current strings is nothing
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToString2(ByRef value As String) As String

        If (value Is Nothing) Then value = ""
        value = value.Replace(Chr(0), "")
        If value = Chr(0) Then value = ""
        Return value.ToString()
    End Function

    ''' <summary>
    ''' Converts unicode strings into ascii string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function FromUnicode(ByVal value As String) As String

        ' obtains unicode byte array from string, converts it into ascii char array
        Dim unicodeBytes As Byte() = System.Text.UnicodeEncoding.Unicode.GetBytes(value)
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(System.Text.Encoding.Unicode, System.Text.Encoding.ASCII, unicodeBytes)

        Dim asciiChars(System.Text.Encoding.ASCII.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        System.Text.Encoding.ASCII.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)

        Return asciiString
    End Function

    ''' <summary>
    ''' Converts unicode strings into ascii string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function FromUTF8(ByVal value As String) As String

        ' obtains unicode byte array from string, converts it into ascii char array
        Dim unicodeBytes As Byte() = System.Text.UnicodeEncoding.UTF8.GetBytes(value)
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, unicodeBytes)

        Dim asciiChars(System.Text.Encoding.ASCII.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        System.Text.Encoding.ASCII.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)

        Return asciiString
    End Function

    ''' <summary>
    ''' Converts the current string into a lowerCase string usable on the web
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToWebName(ByVal Name As String) As String

        Dim sWebName As String = ""

        Name = Name.StripHTML()
        Name = Name.ToLower()
        Name = Name.Replace("à", "a")
        Name = Name.Replace("é", "e")
        Name = Name.Replace("è", "e")
        Name = Name.Replace("ì", "i")
        Name = Name.Replace("ò", "o")
        Name = Name.Replace("ù", "u")

        For j As Int32 = 1 To Name.Length
            Dim sChar As String = Strings.Mid(Name, j, 1)
            If InStr(1, "0123456789abcdefghijklmnopqrstuvwxyz", sChar) = 0 Then sChar = "_"
            sWebName &= sChar
        Next

        ' removes duplicated underscores
        ' and unserscrores at the start and the end of the name
        sWebName = sWebName.RemoveDuplicatedContiguosChar("_")

        While Strings.Left(sWebName, 1) = "_"
            sWebName = Strings.Mid(sWebName, 2)
        End While
        While Strings.Right(sWebName, 1) = "_"
            sWebName = Strings.Left(sWebName, sWebName.Length - 1)
        End While

        Return sWebName
    End Function

    ''' <summary>
    ''' Appends a new string to the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="StringToAppend">string to append</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Append(ByRef value As String, ByVal StringToAppend As String) As String

        If (value Is Nothing) Then value = ""
        value = value & StringToAppend
        Return value
    End Function

    ''' <summary>
    ''' Clears the string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Append(ByRef value As String) As String

        If (value Is Nothing) Then value = ""
        Return value
    End Function

    ''' <summary>
    ''' Adds double-quotes from the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function DoubleQuoted(ByVal value As String) As String

        value = Strings.Replace(value, Chr(34), "")
        Return Chr(34) & value & Chr(34)
    End Function

    ''' <summary>
    ''' Removes double-quotes from the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function DeDoubleQuoted(ByVal value As String) As String

        If (value Is Nothing) Then value = ""
        value = value.Trim()
        If value = "" Then Return ""

        While Strings.Left(value.Trim(), 1) = Chr(34)
            value = Strings.Trim(Strings.Mid(value.Trim(), 2))
        End While

        While Strings.Right(value.Trim(), 1) = Chr(34)
            value = Strings.Trim(Strings.Left(value.Trim(), value.Trim().Length - 1))
        End While

        Return value
    End Function

    ''' <summary>
    ''' Removes all duplicated occurrences of the specified substring from the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="FindString">substring to remove</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function RemoveDuplicatedContiguosChar(ByVal value As String, ByVal FindString As String) As String

        If (value Is Nothing) Then value = ""
        While InStr(1, value, FindString & FindString) > 0
            value = Replace(value, FindString & FindString, FindString)
        End While
        Return value
    End Function

    ''' <summary>
    ''' Capitalizes the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToCapital(ByRef value As String) As String

        Dim sRetVal As String = ""
        Dim sArr() As String
        Dim j As Int32
        Dim sCurWord As String

        If (value Is Nothing) Then value = ""
        value = value.Trim()
        If value = "" Then Return ""

        value = value.RemoveDuplicatedContiguosChar(" ")        ' removes duplicated contiguos spaces

        ' splits the string to space
        If InStr(1, value, " ") > 0 Then
            sArr = Split(value, " ")

            For j = 0 To sArr.Length - 1

                sCurWord = sArr(j)
                If sCurWord <> " " Then

                    ' any word shorter than 4 chars
                    If sCurWord.Length <= 3 Then

                        Dim bContainNumber As Boolean = False

                        For jNum As Int32 = 0 To 9
                            If InStr(1, sCurWord, jNum.ToString().Trim(), CompareMethod.Text) > 0 Then
                                bContainNumber = True
                                Exit For
                            End If
                        Next

                        ' if current word contains numbers (probably a code), we don't change it,
                        ' otherwise lower
                        If bContainNumber Then
                            sCurWord = sCurWord
                        Else
                            sCurWord = sCurWord.ToLower()
                        End If
                    Else
                        sCurWord = (Strings.Left(sCurWord, 1)).ToUpper() + (Strings.Mid(sCurWord, 2)).ToLower()
                    End If

                    If sRetVal = "" Then
                        sRetVal = sCurWord
                    Else
                        sRetVal = sRetVal & " " & sCurWord
                    End If
                End If
            Next
        Else
            sRetVal = (Strings.Left(value, 1)).ToUpper() + (Strings.Mid(value, 2)).ToLower()
        End If

        ' sets first char to upper (in case it was set to lower previously)
        sRetVal = Strings.Left(sRetVal, 1).ToUpper() & Strings.Mid(sRetVal, 2)
        Return sRetVal
    End Function

    ''' <summary>
    ''' Capitalizes phrases into the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="Keywords">keywords to keep unchanged</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToCapitalPhrases(ByRef value As String, Optional ByVal Keywords As String = "") As String

        Dim arrPhrases() As String
        Dim sNewValue As String = value.ToLower().Trim()

        sNewValue = sNewValue.RemoveDuplicatedContiguosChar(" ")

        While InStr(1, sNewValue, ". ") > 0
            sNewValue = sNewValue.Replace(". ", ".")
        End While
        While InStr(1, sNewValue, " .") > 0
            sNewValue = sNewValue.Replace(" ", ".")
        End While
        While InStr(1, sNewValue, "? ") > 0
            sNewValue = sNewValue.Replace("? ", "?")
        End While
        While InStr(1, sNewValue, " ?") > 0
            sNewValue = sNewValue.Replace(" ?", "?")
        End While
        While InStr(1, sNewValue, "! ") > 0
            sNewValue = sNewValue.Replace("! ", "!")
        End While
        While InStr(1, sNewValue, " !") > 0
            sNewValue = sNewValue.Replace(" !", "!")
        End While

        ' splits phrases by . and sets starting-letter to uppercase
        arrPhrases = Split(sNewValue, ".")
        For jPhrase As Int32 = 0 To arrPhrases.Length - 1
            If arrPhrases(jPhrase) <> "" Then
                arrPhrases(jPhrase) = Strings.UCase(Strings.Left(arrPhrases(jPhrase), 1)) & Strings.Mid(arrPhrases(jPhrase), 2)
            End If
        Next
        sNewValue = Join(arrPhrases, ". ")

        ' splits phrases by ? and sets starting-letter to uppercase
        arrPhrases = Split(sNewValue, "?")
        For jPhrase As Int32 = 0 To arrPhrases.Length - 1
            If arrPhrases(jPhrase) <> "" Then
                arrPhrases(jPhrase) = Strings.UCase(Strings.Left(arrPhrases(jPhrase), 1)) & Strings.Mid(arrPhrases(jPhrase), 2)
            End If
        Next
        sNewValue = Join(arrPhrases, "?")

        ' splits phrases by ! and sets starting-letter to uppercase
        arrPhrases = Split(sNewValue, "!")
        For jPhrase As Int32 = 0 To arrPhrases.Length - 1
            If arrPhrases(jPhrase) <> "" Then
                arrPhrases(jPhrase) = Strings.UCase(Strings.Left(arrPhrases(jPhrase), 1)) & Strings.Mid(arrPhrases(jPhrase), 2)
            End If
        Next
        sNewValue = Join(arrPhrases, "!")

        Return sNewValue
    End Function

    ''' <summary>
    ''' Capitalizes the current string by phrases
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function ToCapitalProper(ByRef value As String) As String
        If (value Is Nothing) Then value = ""
        value = StrConv(value, VbStrConv.ProperCase)

        Return value
    End Function

    ''' <summary>
    ''' Removes all HTML code from the current string
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function StripHTML(ByRef value As String, Optional ByVal KeepFormatting As Boolean = False) As String

        If (value Is Nothing) Then value = ""
        If value = "" Then Return ""

        ' Keeps formats, by keeping some HTML Tags
        ' "b", "i", "u", "ul", "ol", "li", "strong", "h1", "h2", "h3", "h4", "h5", "h6"
        ' Replaces them by pseudo-tags
        If KeepFormatting Then
            value = value.Replace(Chr(10), "<br>").Replace(Chr(13), "")
            Dim lstAllowedTags As List(Of String) = New List(Of String) From {"b", "i", "u", "ul", "ol", "li", "strong", "h1", "h2", "h3", "h4", "h5", "h6", "br"}
            For Each sTag As String In lstAllowedTags
                value = value.Replace("<" & sTag & ">", "[" & sTag & "]").Replace("</" & sTag & ">", "[/" & sTag & "]")
            Next
        End If

        ' removes HTML tags
        Dim matchpattern As String = "<(?:[^>=]|='[^']*'|=""[^""]*""|=[^'""][^\s>]*)*>"
        value = Regex.Replace(value, "<.*?>", "")
        value = Regex.Replace(value, matchpattern, "")

        ' Keeps formats, by keeping some HTML Tags
        ' "b", "i", "u", "ul", "ol", "li", "strong", "h1", "h2", "h3", "h4", "h5", "h6"
        ' Replaces pseudo-tags with tags
        If KeepFormatting Then
            Dim lstAllowedTags As List(Of String) = New List(Of String) From {"b", "i", "u", "ul", "ol", "li", "strong", "h1", "h2", "h3", "h4", "h5", "h6", "br"}
            For Each sTag As String In lstAllowedTags
                value = value.Replace("[" & sTag & "]", "<" & sTag & ">").Replace("[/" & sTag & "]", "</" & sTag & ">")
            Next
        End If

        value = System.Web.HttpUtility.HtmlDecode(value)        ' decodes html entities. N.B: this requires .net framework 4 + reference to System.Web (available only in .net framework 4 or higher)
        value = value.Replace(vbTab, " ")
        value = value.Replace(Chr(10) & Chr(32), Chr(10))
        value = value.Replace(Chr(13) & Chr(32), Chr(13))
        value = value.RemoveDuplicatedContiguosChar(" ")        ' removes contiguos duplicate spaces
        value = value.RemoveDuplicatedContiguosChar(vbCrLf).RemoveDuplicatedContiguosChar(Chr(10)).RemoveDuplicatedContiguosChar(Chr(13))
        Return value.Trim()
    End Function

    ''' <summary>
    ''' Decode HTML entities
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function DecodeHTMLEntities(ByRef value As String) As String

        If (value Is Nothing) Then value = ""
        If value = "" Then Return ""

        value = System.Web.HttpUtility.HtmlDecode(value)        ' decodes html entities. N.B: this requires .net framework 4 + reference to System.Web (available only in .net framework 4 or higher)
        value = value.RemoveDuplicatedContiguosChar(" ")        ' removes contiguos duplicate spaces
        Return value.Trim()
    End Function

    ''' <summary>
    ''' Converts generic string into a DateTime
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function fromGenericStringToDate(ByRef value As String) As DateTime

        Dim iDay As Int32, iMonth As Int32, iYear As Int32

        If (value Is Nothing) Then value = ""
        value = value.Trim().RemoveDuplicatedContiguosChar(" ").ToLower()

        value = Strings.Replace(value, "-", "/")
        value = Strings.Replace(value, ".", "/")
        value = Strings.Replace(value, "\", "/")
        value = Strings.Replace(value, " ", "")
        value = value.RemoveDuplicatedContiguosChar("/")

        If IsDate(value) Then Return CDate(value)

        Select Case value.Length
            Case 8
                ' 20121231
                iYear = CLng(Strings.Left(value, 4))
                iMonth = CLng(Strings.Mid(value, 5, 2))
                iDay = CLng(Strings.Right(value, 2))

                Return DateSerial(iYear, iMonth, iDay)
        End Select

        Return Date.MinValue
    End Function

    ''' <summary>
    ''' Returns a substring extracted from the source string, between String1 and String2
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="String1">the substring from which to extract the substring, or NullString to start from the first char of the source string</param>
    ''' <param name="String2">the substring to which to extract the substring, or NullString to stop at the last char of the source string</param>
    ''' <param name="BothRequired">true if both substrings are required</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ExtractSubStringBetween(ByVal value As String, Optional ByVal String1 As String = "", Optional ByVal String2 As String = "", Optional ByVal BothRequired As Boolean = False) As String

        If (value Is Nothing) Then value = ""
        Dim sRetVal As String = value

        If (String1 Is Nothing) Then String1 = ""
        If (String2 Is Nothing) Then String2 = ""

        If BothRequired Then
            If InStr(1, value, String1) = 0 Or InStr(1, value, String2) = 0 Then Return ""
        End If

        If String1 <> "" Then
            If InStr(1, sRetVal, String1, CompareMethod.Text) > 0 Then
                sRetVal = Strings.Mid(sRetVal, InStr(1, sRetVal, String1, CompareMethod.Text) + String1.Length)
            End If
        End If
        If String2 <> "" Then
            If InStr(1, sRetVal, String2, CompareMethod.Text) > 0 Then
                sRetVal = Strings.Left(sRetVal, InStr(1, sRetVal, String2, CompareMethod.Text) - 1)
            End If
        End If
        Return sRetVal
    End Function

    ''' <summary>
    ''' Extracts links from an HTML text and returns them into a disctionary (a-name, a-link)
    ''' </summary>
    ''' <param name="value">input string (html code)</param>
    ''' <returns>Returns the list of links into a disctionary (a-name, a-link)</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ExtractLinks(ByRef value As String, Optional ByVal AcceptBrokenLinks As Boolean = False) As Dictionary(Of String, String)

        Dim retVal As New Dictionary(Of String, String)
        Dim re As New Regex("<a.*?href=[""'](?<url>.*?)[""'].*?>(?<name>.*?)</a>", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = re.Matches(value)

        For Each m As Match In matches
            ' extract A tag. So we extract link (url)
            Dim s As String = m.Value.ToString()
            Dim sLink = s.ExtractSubStringBetween(Chr(34), Chr(34), True)
            Dim sName = s.ExtractSubStringBetween(">", "</a>", True)
            If InStr(1, sName, "<") > 0 Then sName = sName.ExtractSubStringBetween("", "<", True)
            If Not AcceptBrokenLinks Then                   ' if doesn't accept broken links, checks if link is a valid uri
                If Not sLink.IsValidURI() Then sLink = ""
            End If
            If sLink <> "" Then
                Dim i As Int32 = 0                          ' creates new key if key exists
                While retVal.ContainsKey(sName)
                    i += 1
                    If i = 1 Then sName &= "-"
                    sName &= i.ToString()
                End While
                retVal.Add(sName, sLink)
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' Extracts iFrames Src(s) from an HTML text and returns them into a List of string
    ''' </summary>
    ''' <param name="value">input string (html code)</param>
    ''' <returns>Returns the list of iFrame sources into a List</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ExtractIFrames(ByRef value As String, Optional ByVal AcceptBrokenLinks As Boolean = False) As List(Of String)

        Dim retVal As New List(Of String)
        Dim re As New Regex("<iframe.*?src=[""'](?<url>.*?)[""'].*?>", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = re.Matches(value)

        For Each m As Match In matches
            ' extract A tag. So we extract link (url)
            Dim s As String = m.Value.ToString()
            Dim sLink = s.ExtractSubStringBetween("src=" & Chr(34), Chr(34), True)
            If Not AcceptBrokenLinks Then                   ' if doesn't accept broken links, checks if link is a valid uri
                If Not sLink.IsValidURI() Then sLink = ""
            End If
            If sLink <> "" Then retVal.Add(sLink)
        Next
        If retVal.Count > 0 Then retVal = retVal.Distinct().ToList() ' distinct the list
        Return retVal
    End Function

    ''' <summary>
    ''' Extracts images url from an HTML text and returns them into list of string
    ''' </summary>
    ''' <param name="value">input string (html code)</param>
    ''' <returns>Returns the listof uri</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ExtracImageUrls(ByRef value As String) As List(Of String)

        Dim retVal As New List(Of String)

        Dim re As New Regex("<img.*?src=[""'](?<url>.*?)[""'].*?>", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = re.Matches(value)

        For Each m As Global.System.Text.RegularExpressions.Match In matches
            ' extract A tag. So we extract link (url)
            Dim s As String = m.Value.ToString()
            Dim sSrc = s.ExtractSubStringBetween("src=""", Chr(34), True).Trim()
            If sSrc.IsValidURI Then
                retVal.Add(sSrc)
            End If
        Next
        If retVal.Count > 0 Then retVal = retVal.Distinct().ToList() ' distinct the list
        Return retVal
    End Function

    ''' <summary>
    ''' Converts a string into boolean
    ''' </summary>
    ''' <param name="value">input string</param>
    ''' <returns>Returns true/false</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToBoolean(ByRef value As String) As Boolean

        If value.Trim.ToUpper() = "Y" Then Return True
        If value.Trim.ToUpper() = "YES" Then Return True
        If value.Trim.ToUpper() = "T" Then Return True
        If value.Trim.ToUpper() = "TRUE" Then Return True
        If value.Trim.ToUpper() = "V" Then Return True
        If value.Trim.ToUpper() = "VERO" Then Return True
        If value.Trim.ToUpper() = "OK" Then Return True
        If value.Trim.ToUpper() = "1" Then Return True

        Return False
    End Function
#End Region

#Region "Double Extensions"
    ' Double Extensions

    ''' <summary>
    ''' Converts a double into hex
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToHex(ByVal value As Double) As String
        Return Hex(value).ToUpper()
    End Function

    ''' <summary>
    ''' Converts the current string into a currency value, by detecting the system decimal separator char
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToCurrency(ByVal value As String) As Double

        Dim sep As String = currentCurrencySeparator()      ' system currency separator

        If (value Is Nothing) Then value = ""
        value = value.Replace("EUR", "") : value = value.Replace("€", "")
        value = value.Replace("USD", "") : value = value.Replace("$", "")
        value = value.Replace("GBP", "") : value = value.Replace("£", "")
        value = value.Trim()

        value = Replace(value, ".", sep)
        value = Replace(value, ",", sep)
        If Strings.Left(value, 1) = sep Then value = "0" & value ' prefix valur by 0 if it start by currency separator

        ' prevents error cause by presence of thousands separator (es: 1.088,93 becomes 1,088,93 so we remove first separator leaving only the last one)
        If InStr(1, value, sep) = 0 Then value &= sep & "0"
        Dim arrValueParts() As String = Split(value, sep)
        value = ""
        For iParts As Int32 = 0 To arrValueParts.Length - 1
            If iParts = arrValueParts.Length - 1 Then value &= sep
            value &= arrValueParts(iParts)
        Next
        If IsNumeric(value) Then
            Return CDbl(value)
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' Converts Currency into string using the current system decimal separator
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToCurrencyStr(ByVal Value As Double) As String

        Dim sep As String = currentCurrencySeparator()

        Dim sValue As String = Value.ToString()

        sValue = sValue.Trim()
        sValue = Replace(sValue, ".", sep)
        sValue = Replace(sValue, ",", sep)
        Return sValue
    End Function

    ''' <summary>
    ''' Converts the current value from/to two different currencies
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="FromCurrency">Original Currency (ex: "EUR")</param>
    ''' <param name="ToCurrency">Destination Currency (ex: "USD")</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ConvertCurrency(ByVal Value As Double, ByVal FromCurrency As String, ByVal ToCurrency As String) As Double
        'TODO:  

        '' 20/11/2013 Google ha cambiato interfaccia.
        '' url del tipo http://www.google.com/ig/calculator?hl=en&q=1USD=?EUR non funzionano più
        '' Utilizziamo altra interfaccia https://www.google.com/finance/converter?a=1&from=EUR&to=GBP

        'FromCurrency = FromCurrency.Trim().ToUpper()          ' normalizza valute
        'ToCurrency = ToCurrency.Trim().ToUpper()

        'If Value = 0 Then Return 0 ' zero: non necessaria conversione
        'If FromCurrency = ToCurrency Then Return Value ' valute uguali: non necessaria conversione
        'If ToCurrency = "" Or FromCurrency = "" Then Return Value ' valute omesse: non può eseguire conversione!!!

        '' ottiene versione stringa del valore con separatore . richiesto da API di Google
        'Dim sValue As String = Value.ToCurrencyStr()
        'sValue = Strings.Replace(sValue, currentCurrencySeparator(), ".")


        '' genera url API Google, esegue webrequest e parsing del json risultante
        'Dim sUrl As String = "https://www.google.com/finance/converter?a={VALUE}&from={FROM_CURRENCY}&to={TO_CURRENCY}"
        'sUrl = sUrl.Replace("{VALUE}", sValue)
        'sUrl = sUrl.Replace("{FROM_CURRENCY}", FromCurrency)
        'sUrl = sUrl.Replace("{TO_CURRENCY}", ToCurrency)

        '' * nuovo codice, vedi OLD
        'Dim sBody As String = Web.getUrl(sUrl)
        'sBody = sBody.ExtractSubStringBetween("<span class=bld>", "</span>", True)              ' estrae prezzo convertito
        'If sBody = "" Then Return Value ' conversione non eseguita
        'sBody = sBody.Replace(ToCurrency, "").Replace(" ", "")                                  ' rimuove codice valuta destinazione e spazi

        'If sBody <> "" Then
        '    ' conversione eseguita
        '    sBody = sBody.Replace(".", currentCurrencySeparator())       ' imposta separatore decimali
        '    Value = CDbl(sBody)
        'End If

        'Return Value
    End Function

    ''' <summary>
    ''' Converts current value from Inches to Centimeters
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function InchesToCm(ByVal value As Double) As Double

        Dim k As Double = 2.54
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value *= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Converts current value from Inches to Millimeters
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function InchesToMm(ByVal value As Double) As Double

        Dim k As Double = 25.4
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value *= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Converts current value from Centimeters to Inches
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function CmToInches(ByVal value As Double) As Double

        Dim k As Double = 2.54
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value /= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Converts current value from Millimeters to Inches
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function MmToInches(ByVal value As Double) As Double

        Dim k As Double = 25.4
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value /= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Converts current value from Pounds (LBS) to Kg
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function PoundsToKg(ByVal value As Double) As Double
        ' converts lbs into Kg

        Dim k As Double = 2.2
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value /= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Converts current value from Kg to Pounds (LBS)
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function KgToPounds(ByVal value As Double) As Double
        ' converts kg into lbs

        Dim k As Double = 2.2
        Dim bNeg As Boolean = False

        If value = 0 Then Return 0
        If value < 0 Then
            bNeg = True
            value = Math.Abs(value)
        End If

        value *= k
        If bNeg Then value = 0 - value

        Return value
    End Function

    ''' <summary>
    ''' Rounds the current value to the specified number of decimal digits
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="Digits">Number of decimal digits. Default value is 2</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function Round(ByVal Value As Double, Optional ByVal Digits As Int32 = 2) As Double

        If Digits < 0 Then Digits = 0
        Return Math.Round(Value, Digits)
    End Function
#End Region

#Region "Int32 Extensions"
    ' Int32 Extensions

    ''' <summary>
    ''' Converts seconds to Time representation hh:mm:ss
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function SecondsToTime(ByVal value As Int32) As String
        If value < 0 Then value = 0
        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(value)
        Dim sTime As String = iSpan.Hours.ToString.PadLeft(2, "0"c) & ":" &
                    iSpan.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                    iSpan.Seconds.ToString.PadLeft(2, "0"c)
        Return sTime
    End Function

    ''' <summary>
    ''' Converts a Int32 into hex
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToHex(ByVal value As Int32) As String
        Return Hex(value).ToUpper()
    End Function

    ''' <summary>
    ''' Converts a Int32 into a version number (such as 1.0.1)
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Function ToVersionNumber(ByVal value As Int32) As String
        '******************************************************
        ' Rende un numero di versione (es. 100) in formato stringa (1.0.0)
        '******************************************************

        Dim sRetVal As String = ""
        Dim sInt As String = "", sDec As String = ""
        Dim sDummy As String = "", dDummy As Double = 0, arrDummy() As String
        Dim lDots, jC, sC

        ' default
        sRetVal = ""
        lDots = 0

        ' normalizza dati
        If value < 0 Then value = Math.Abs(value)
        If value = 0 Then                   ' valore predefinito
            Return "0.0.1"
        End If

        If Fix(value) <> value Then
            ' il num.versione ha una parte decimale
            ' es:   0.91        =>  0.9.1
            '       17.2545     =>  17.2.545
            dDummy = Round(value - Fix(value), 14)  ' estrae parte decimale. NOTA: arrotonda a 14 decimali per evitare numeri lunghi
            sInt = CStr(Fix(value))
            sDec = Replace(CStr(dDummy), "0.", "")
            sDummy = sInt
            For jC = 1 To Len(sDec)
                If lDots = 1 Then
                    sC = Mid(sDec, jC)
                    sDummy = sDummy & "." & sC
                    Exit For
                Else
                    sC = Mid(sDec, jC, 1)
                    sDummy = sDummy & "." & sC
                    lDots = lDots + 1
                End If
            Next
        Else
            ' il num.versione non ha una parte decimale
            ' es:   101         =>  1.0.1
            '       7543        =>  7.5.43
            sInt = CStr(value)
            For jC = 1 To Len(sInt)
                If lDots = 2 Then
                    sC = Mid(sInt, jC)
                    sDummy = sDummy & "." & sC
                    Exit For
                Else
                    sC = Mid(sInt, jC, 1)
                    If sDummy <> "" Then sDummy = sDummy & "."
                    sDummy = sDummy & sC
                    lDots = lDots + 1
                End If
            Next
        End If

        ' Conta parti del numero di versione. Richiede almeno 3 parti
        arrDummy = Split(sDummy, ".")
        If UBound(arrDummy) = 1 Then sDummy = sDummy & ".0"
        sRetVal = sDummy

        ' return
        Return sRetVal
    End Function
#End Region

#Region "List Extensions"
    ' List Extensions

    <Extension()>
    Public Sub Add(ByRef list As List(Of String), ParamArray values As String())
        For Each s As String In values
            list.Add(s)
        Next
    End Sub
#End Region

#Region "Date Extensions"
    ' Date Extensions

    ''' <summary>
    ''' Returns a timestamp from date/time
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks>Test by https://www.epochconverter.com/</remarks>
    ''' <returns></returns>
    <Extension()>
    Public Function ToUnix(ByVal value As Date) As String
        Dim dtStart As Date = New Global.System.DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc)
        Dim ts As TimeSpan = (value - dtStart)
        Return ts.TotalSeconds.ToString()
    End Function

    ''' <summary>
    ''' Converts a timestamp into date/time
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks>Test by https://www.epochconverter.com/</remarks>
    ''' <returns></returns>
    <Extension()>
    Public Function FromUnix(ByRef value As DateTime, ByVal TimeStamp As String)
        Dim dtStart As Date = New Global.System.DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc)
        Dim lSeconds As Int64 = CLng(TimeStamp)
        value = dtStart.AddSeconds(lSeconds)
    End Function
#End Region

#Region "Privates"
    '
    Private Function currentCurrencySeparator() As String
        ' restituisce il separatore di decimali in base alle impostazioni internazioni di sistema

        Dim cDot As Double = CDbl("9.9")
        Dim cComma As Double = CDbl("9,9")

        If cDot = 9.9 Then
            Return "."
        Else
            Return ","
        End If
    End Function
#End Region
End Module