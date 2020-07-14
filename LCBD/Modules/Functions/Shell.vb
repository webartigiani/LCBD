''' <summary>
''' Shell.vb
''' </summary>
Module Shell
#Region "Public Members"
    ' Public Members

    ''' <summary>
    ''' Executes a command
    ''' </summary>
    ''' <param name="ProcessPath">application path</param>
    ''' <param name="Arguments">command line arguments</param>
    ''' <param name="Wait">true to make code wait</param>
    ''' <param name="SetWorkingDirectory">Sets/not the application directory as working directory</param>
    ''' <param name="WindowStyle">WindowStyle: Normal, Maximized, Minimized, Hidden</param>
    ''' <param name="ExitCode">Returns the process exit code when "wait" is true</param>
    ''' <param name="GetOutput">True to get the application output</param>
    ''' <param name="Output">Returns the application output</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShellAndWait(ByVal ProcessPath As String, Optional ByVal Arguments As String = "", Optional ByVal Wait As Boolean = False, Optional ByVal SetWorkingDirectory As Boolean = False, Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional ByVal GetOutput As Boolean = False, Optional ByRef ExitCode As Int32 = 0, Optional ByRef Output As String = "") As Boolean

        ' default
        Dim bRetVal As Boolean = False
        ExitCode = 0
        Output = ""

        ' Normalizes arguments
        Arguments = Arguments.Trim()

        Dim objProcess As Global.System.Diagnostics.Process
        Try
            objProcess = New Global.System.Diagnostics.Process()
            If SetWorkingDirectory Then objProcess.StartInfo.WorkingDirectory = AppInfo.Path
            objProcess.StartInfo.WindowStyle = WindowStyle
            objProcess.StartInfo.FileName = ProcessPath
            If GetOutput Then                           ' if we have to capture program output
                objProcess.StartInfo.UseShellExecute = False
                objProcess.StartInfo.RedirectStandardOutput = True
            End If
            If Arguments.Trim() <> "" Then objProcess.StartInfo.Arguments = Arguments.Trim()
            objProcess.Start()
            If Wait Then
                objProcess.WaitForExit()             ' Wait until the process passes back an exit code
                ExitCode = objProcess.ExitCode       ' Returns process exit code and the output
                If GetOutput Then Output = objProcess.StandardOutput.ReadToEnd()
                objProcess.Close()                   ' Free resources associated with this process
            Else
                ExitCode = objProcess.ExitCode       ' Returns process exit code and the output
                If GetOutput Then Output = objProcess.StandardOutput.ReadToEnd()
                objProcess.Close()                   ' Free resources associated with this process
            End If

            bRetVal = True
        Catch ex As Exception
            ' Error
            Debug.Print(ex.Message)
        End Try

        ' return
        ShellAndWait = bRetVal
    End Function

    ''' <summary>
    ''' Checks if application or command is supported
    ''' </summary>
    ''' <param name="ProcessPathOrName">application path or name</param>
    ''' <param name="Arguments">command line arguments</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckCommand(ByVal ProcessPathOrName As String, Optional ByVal Arguments As String = "") As Boolean

        Dim bRetVal As Boolean = False

        Dim objProcess As Global.System.Diagnostics.Process
        Try
            objProcess = New Global.System.Diagnostics.Process()
            objProcess.StartInfo.WorkingDirectory = AppInfo.Path
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            objProcess.StartInfo.FileName = ProcessPathOrName
            If Arguments.Trim() <> "" Then objProcess.StartInfo.Arguments = Arguments.Trim()
            objProcess.Start()
            objProcess.Close()
            bRetVal = True
        Catch ex As Exception
        End Try

        Return bRetVal
    End Function
#End Region
End Module