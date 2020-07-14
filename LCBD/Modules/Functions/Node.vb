''' <summary>
''' Node.vb
''' </summary>
Module Node
#Region "Public Properties"

#End Region

#Region "Public Methods"
    ' Public Methods

    ''' <summary>
    ''' Returns true if Node.js is installed
    ''' </summary>
    ''' <returns>DO NOT USE IT! WE CANNOT READ THE PROGRAM OUTPUT!!!</returns>
    Public Function GetVersion() As String
        'Dim sRetVal As String = ""
        'Shell.ShellAndWait("cmd.exe", "/K node -v", False, True, ProcessWindowStyle.Hidden, True, 0, sRetVal)
    End Function

    ''' <summary>
    ''' Returns true if Node.js is installed
    ''' </summary>
    ''' <returns></returns>
    Public Function HasNode() As Boolean
        Return Shell.CheckCommand("cmd.exe", "/C node -v")
    End Function

    ''' <summary>
    ''' Returns true if NPM (Node Package Manager) is installed
    ''' </summary>
    ''' <returns></returns>
    Public Function HasNPM() As Boolean
        Return Shell.CheckCommand("npm")
    End Function

    ''' <summary>
    ''' Downloads and install Node.js v.10.16.3 (x86/x64, depending on the current operating system)
    ''' </summary>
    ''' <remarks>
    ''' Downloads https://nodejs.org/dist/v10.16.3/node-v10.16.3-x86.msi
    ''' or https://nodejs.org/dist/v10.16.3/node-v10.16.3-x64.msi
    ''' </remarks>
    ''' <returns></returns>
    Public Function Install(Optional ByRef Version As String = "", Optional ByRef ErrorDescription As String = "") As Boolean

        Dim bRetVal As Boolean = False
        Dim sTemplateUrl As String = "https://nodejs.org/dist/v{version}/node-v{version}-{osx}.msi"
        Dim sTemplateFileName As String = "node-v{version}-{osx}.msi"
        Version = "10.16.3"
        Dim sOSX As String = "x86"
        ErrorDescription = ""

        Dim sUrl As String = ""
        Dim sFileName As String = ""

        If SystemInfo.Is64Bit Then sOSX = "x64"                           ' detects if OS is 64 or 32 bits
        sUrl = sTemplateUrl.Replace("{version}", Version).Replace("{osx}", sOSX)                ' calculates url to download
        sFileName = sTemplateFileName.Replace("{version}", Version).Replace("{osx}", sOSX)      ' calculates download filename

        Call FileSystem.CreateFolderTree("temp\downloads\")
        If Web.DownloadFile(sUrl, "temp\downloads\" & sFileName) Then

            Dim sCmd As String = "/i " & Chr(34) & "temp\downloads\" & sFileName & Chr(34) & " /qn"
            If Shell.ShellAndWait("msiexec.exe", sCmd, True, True) Then
                bRetVal = True
            Else
                ErrorDescription = "Unable to install 'Node.js'. Please, proceed to manually download and install Node.js"
            End If
        Else
            ' Download error!
            ErrorDescription = "Unable to download '" & sUrl & "'. Please, proceed to manually download and install Node.js"
        End If

        ' return
        Install = bRetVal
    End Function

    ''' <summary>
    ''' Exec concurrently a list of node.js scripts with their command-lines, and waits until the end of all the processes
    ''' </summary>
    ''' <param name="ScriptCommandsAndFileName">Dictionary of command-lines (keys) and script-filenames (values)</param>
    ''' <returns>true in case of success</returns>
    ''' <remarks></remarks>
    Public Function ExecNodeScripts(ByVal ScriptCommandsAndFileName As Dictionary(Of String, String)) As Boolean
        Dim threadList As List(Of Threading.Thread) = New List(Of Threading.Thread)

        For Each kvp As KeyValuePair(Of String, String) In ScriptCommandsAndFileName
            Dim sScriptArguments As String = kvp.Key
            Dim sScriptFileName As String = kvp.Value

            Dim t = New Global.System.Threading.Thread(CType(Sub() Node.ExecNodeScript(CType(sScriptFileName, System.String), CType(sScriptArguments, System.String)), Threading.ThreadStart))
            t.IsBackground = True
            t.Start()
            threadList.Add(t)
        Next

        ' Wait for all threads to finish.
        ' The loop will only exit once all threads have completed their work.
        For Each t In threadList
            t.Join()
        Next
        Return True
    End Function

    ''' <summary>
    ''' Exec a single node.js scripts with command-lines
    ''' </summary>
    ''' <param name="ScriptFileName">Script Filename</param>
    ''' <param name="ScriptArguments">arguments to send to the script</param>
    ''' <returns>true in case of success</returns>
    ''' <remarks></remarks>
    Public Function ExecNodeScript(ByVal ScriptFileName As String, Optional ByVal ScriptArguments As String = "") As Boolean

        Dim sCmdLine As String = ""
        ScriptArguments = ScriptArguments.Trim()

        sCmdLine = ScriptFileName
        sCmdLine = sCmdLine & " " & ScriptArguments
        sCmdLine = sCmdLine.Trim()
        sCmdLine = "/C node " & sCmdLine        ' /C Command and then terminate;    /K Run Command and then return to the CMD prompt. This Is useful For testing, to examine variables

        Dim bRetVal As Boolean = Shell.ShellAndWait("cmd.exe", sCmdLine, True, True, ProcessWindowStyle.Hidden)
        Return bRetVal
    End Function
#End Region
End Module