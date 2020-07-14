''' <summary>
''' Environment.vb
''' </summary>
Module Environment

#Region "Public Properties"
    ' Public Methods

    ''' <summary>
    ''' Returns the environment Root folder
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property SourcesFolder() As String
        Get
            Return "sources"
        End Get
    End Property

    ''' <summary>
    ''' Returnshe a specific environment folder
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Folder(ByVal FolderName As String) As String
        Get
            Dim sRetVal As String = ""

            FolderName = FolderName.Replace("/", "\").Trim()
            If FolderName = "" Then Exit Property
            If Left(FolderName, 1) = "\" Then FolderName = Mid(FolderName, 2)
            If Right(FolderName, 1) = "\" Then FolderName = Left(FolderName, Len(FolderName) - 1)
            If FolderName <> "" Then Return FolderName
        End Get
    End Property
#End Region
End Module
