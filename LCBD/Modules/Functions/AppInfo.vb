''' <summary>
''' AppInfo.vb
''' </summary>
Module AppInfo

#Region "Public Properties"
    ' Public Methods

    ''' <summary>
    ''' Returns the AppName
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AppName() As String
        Get
            Return My.Application.Info.AssemblyName
        End Get
    End Property

    ''' <summary>
    ''' Returns the ApPShortName
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AppShortName() As String
        Get
            Dim sASN As String = AppInfo.AppName
            If InStr(1, sASN, "-") > 0 Then sASN = Strings.Left(sASN, InStr(1, sASN, "-") - 1)
            Return sASN
        End Get
    End Property

    ''' <summary>
    ''' Returns the AppVersion
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AppVersion() As String
        Get
            Return My.Application.Info.Version.ToString().Trim()
        End Get
    End Property

    ''' <summary>
    ''' Returns the AppVersionNum
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AppVersionNum() As String
        Get
            Return AppInfo.AppVersion.Replace(".", "")
        End Get

    End Property
    ''' <summary>
    ''' Returns the application path
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Path() As String
        Get
            Dim sRetVal As String = My.Application.Info.DirectoryPath
            If Strings.Right(sRetVal, 1) <> "\" Then sRetVal &= "\"
            Return sRetVal
        End Get
    End Property

    Public ReadOnly Property AppSavePath() As String
        Get
            Return My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & AppInfo.AppName
        End Get
    End Property

    ''' <summary>
    ''' Returns the current startup CommandLine parameters
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CommandLineParams() As List(Of String)
        Get
            Dim lstParams As List(Of String) = New List(Of String)
            Dim arrParams() As String = Global.System.Environment.GetCommandLineArgs()

            If arrParams.Length > 0 Then
                lstParams.AddRange(arrParams)
                lstParams.RemoveAt(0)           ' removes App path/filename
            End If

            Return lstParams
        End Get
    End Property
#End Region
End Module