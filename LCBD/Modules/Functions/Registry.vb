''' <summary>
''' Registry.vb
''' reads/writes registry into "HKEY_CURRENT_USER\Software\VB and VBA Program Settings"
''' </summary>
Module Registry
#Region "Public Members"
    ' Public Members

    ''' <summary>
    ''' Reads and returns a setting from the windows registry
    ''' </summary>
    ''' <param name="Section">Windows Registry Section</param>
    ''' <param name="ParamName">Windows Registry Parameter name</param>
    ''' <param name="DefaultValue">The default value to read</param>
    ''' <param name="Encode">True if the values has been encoded when written</param>
    ''' <returns>The read value or the default value, if missing</returns>
    ''' <remarks></remarks>
    Function ReadSetting(ByVal Section As String, ByVal ParamName As String, Optional ByVal DefaultValue As String = "", Optional ByVal Encode As Boolean = False) As String

        Dim sRetVal As String

        sRetVal = GetSetting(AppInfo.AppName, Section, ParamName, "")
        If sRetVal = "" Then
            sRetVal = DefaultValue          ' returns the default value
        Else

            If Encode Then sRetVal = RegDecode(sRetVal)
        End If
        Return sRetVal
    End Function

    ''' <summary>
    ''' Writes a setting into the Windows Registry
    ''' </summary>
    ''' <param name="Section">Windows Registry Section</param>
    ''' <param name="ParamName">Windows Registry Parameter name</param>
    ''' <param name="Value">The value to write</param>
    ''' <param name="Encode">True to encode the value</param>
    ''' <remarks></remarks>
    Sub WriteSetting(ByVal Section As String, ByVal ParamName As String, ByVal Value As String, Optional ByVal Encode As Boolean = False)

        If Encode Then Value = RegEncode(Value) ' encodes?
        SaveSetting(AppInfo.AppName, Section, ParamName, Value)
    End Sub

    ''' <summary>
    ''' Removes a stored value from the Windows Registry
    ''' </summary>
    ''' <param name="Section">Windows Registry Section</param>
    ''' <param name="ParamName">the Windows Registry parameter name to remove</param>
    ''' <remarks></remarks>
    Sub RemoveSetting(ByVal Section As String, ByVal ParamName As String)
        DeleteSetting(AppInfo.AppName, Section, ParamName)
    End Sub

    ''' <summary>
    ''' Deletes all values and all sub-keys into a Windows Registry section
    ''' </summary>
    ''' <param name="Section">Windows Registry Section</param>
    ''' <remarks></remarks>
    Sub DeleteSection(ByVal Section As String)
        Try
            DeleteSetting(AppInfo.AppName, Section)
        Catch ex As Exception
            ' errore
        End Try
    End Sub
#End Region

#Region "Private Members"
    ' Private Members

    ''' <summary>
    ''' Encodes a value
    ''' </summary>
    ''' <param name="value">value to encode</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RegEncode(ByVal value As String) As String
        If value.Length = 0 Then Return ""
        Dim sRetVal As String = ""

        For j As Int32 = 1 To value.Length
            sRetVal = sRetVal & Strings.Right("00" & Hex(Asc(Strings.Mid(value, j, 1))), 2)
        Next
        Return sRetVal.ToLower()
    End Function

    ''' <summary>
    ''' Decodes an encoded value
    ''' </summary>
    ''' <param name="value">value to decode</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RegDecode(ByVal value As String) As String
        If value.Length = 0 Then Return ""
        Dim sRetVal As String = ""

        For j As Int32 = 1 To value.Length Step 2

            Dim sHex As String = Strings.Mid(value, j, 2)
            Dim lHex As Int32 = "&H" & sHex

            sRetVal = sRetVal & Chr(lHex)
        Next
        Return sRetVal
    End Function
#End Region
End Module