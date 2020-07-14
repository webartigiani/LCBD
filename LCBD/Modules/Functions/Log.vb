''' <summary>
''' Log.vb
''' </summary>
Module Log

#Region "Variables"
    ' Variables

    Private msLogFile As String = "log.txt"
#End Region

#Region "Public Members"
    ' Public Members

    ''' <summary>
    ''' Writes a line into log
    ''' </summary>
    ''' <param name="LogLine"></param>
    ''' <returns></returns>
    Public Function Log(ByVal LogLine As String) As Boolean

        LogLine = LogLine.Replace(Chr(10), " ").Replace(Chr(13), " ").Replace(vbTab, " ").RemoveDuplicatedContiguosChar(" ")
        LogLine = LogLine.Trim()
        If LogLine = "" Then Return True

        Dim sRecord As String = Now().ToString() & vbTab
        sRecord &= AppInfo.AppName & " v." & AppInfo.AppVersion & vbTab
        sRecord &= LogLine

        FileSystem.appendTextFile(msLogFile, sRecord,, True)
        Return True
    End Function

    ''' <summary>
    ''' Write an error line into log
    ''' </summary>
    ''' <param name="ErrorDescription"></param>
    ''' <returns></returns>
    Public Function Err(ByVal ErrorDescription As String) As Boolean

        ErrorDescription = ErrorDescription.Replace(Chr(10), " ").Replace(Chr(13), " ").Replace(vbTab, " ").RemoveDuplicatedContiguosChar(" ")
        ErrorDescription = ErrorDescription.Trim()
        If ErrorDescription = "" Then Return True

        Dim sRecord As String = Now().ToString() & " ERROR" & vbTab
        sRecord &= AppInfo.AppName & " v." & AppInfo.AppVersion & vbTab
        sRecord &= ErrorDescription

        FileSystem.appendTextFile(msLogFile, sRecord,, True)
        Return True
    End Function

    ''' <summary>
    ''' Reads the content of the log
    ''' </summary>
    ''' <returns></returns>
    Public Function Read() As String
        Dim sRetVal As String = FileSystem.readTextFile(msLogFile)
        Return sRetVal
    End Function

    ''' <summary>
    ''' Clears the content of the log
    ''' </summary>
    ''' <returns></returns>
    Public Function Clear() As Boolean
        FileSystem.DeleteFile(msLogFile)
    End Function
#End Region
End Module