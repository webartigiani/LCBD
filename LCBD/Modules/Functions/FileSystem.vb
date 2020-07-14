''' <summary>
''' FileSystem.vb
''' </summary>
Module FileSystem

#Region "File Functions"
    ' File Functions

    ''' <summary>
    ''' Returns true if the specified file exists
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FileExists(ByVal FileName As String) As Boolean
        Try
            Return System.IO.File.Exists(FileName)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns a temp-unique filename
    ''' </summary>
    ''' <param name="Extension"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTempFileName(Optional ByVal Extension As String = "") As String

        Dim sRetVal As String = Guid.NewGuid().ToString().ToLower()

        Extension = Extension.Trim().ToLower()
        Extension = Extension.Replace(".", "").Trim()
        If Extension = "" Then Extension = ".tmp"
        If Strings.Left(Extension, 1) <> "." Then Extension = "." & Extension

        Return sRetVal & Extension
    End Function

    ''' <summary>
    ''' Returns the FileName file size
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileSize(ByVal FileName As String) As Int32

        Try
            Dim oFI As Global.System.IO.FileInfo = New Global.System.IO.FileInfo(CType(FileName,System.String))
            Return oFI.Length
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Copies a file from source to destination
    ''' </summary>
    ''' <param name="Source"></param>
    ''' <param name="Destination"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CopyFile(ByVal Source As String, ByVal Destination As String) As Boolean

        Source = Source.Trim()
        Destination = Destination.Trim()
        Try
            System.IO.File.Copy(Source, Destination, True)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Moves a file from source to destination
    ''' </summary>
    ''' <param name="Source"></param>
    ''' <param name="Destination"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MoveFile(ByVal Source As String, ByVal Destination As String) As Boolean

        Source = Source.Trim().ToLower()
        Destination = Destination.Trim().ToLower()
        If Source = Destination Then Return False
        Try
            System.IO.File.Copy(Source, Destination, True)
            FileSystem.DeleteFile(Source)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Tryes to delete a file
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteFile(ByVal FileName As String) As Boolean
        If System.IO.File.Exists(FileName) Then
            Try
                System.IO.File.Delete(FileName)
                Return True
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End If
        Return False
    End Function

    ''' <summary>
    ''' Converts the input file into a Base64 string
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    Public Function FileToBase64(ByVal FileName As String) As String
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(FileName))
    End Function
#End Region

#Region "Text File Functions"
    ' Text File Functions

    ''' <summary>
    ''' Read the specified text file and returns its content
    ''' </summary>
    ''' <param name="FileName">filename of the file to read</param>
    ''' <returns>returns the content of the file, or null string if the file doesn't esists or is empty</returns>
    ''' <remarks></remarks>
    Public Function readTextFile(ByVal FileName As String, Optional ByVal Encoding As String = "UTF-8") As String

        If System.IO.File.Exists(FileName) Then
            Dim strContents As String = ""
            Dim objReader As System.IO.StreamReader
            Try
                ' opens and reads the downloaded file
                Select Case Encoding.ToLower().Trim()
                    Case "utf7", "utf-7" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.UTF7)
                    Case "utf8", "utf-8" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.UTF8)
                    Case "utf16", "utf-16" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.BigEndianUnicode)
                    Case "utf32", "utf-32" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.UTF32)
                    Case "unicode" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.Unicode)
                    Case "ascii" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.ASCII)
                    Case "ansii", "ansi" : objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.Default)
                    Case Else
                        ' ascii
                        objReader = New System.IO.StreamReader(FileName, System.Text.Encoding.Default)
                End Select
                strContents = objReader.ReadToEnd()
                objReader.Close()

                ' return
                Return strContents
            Catch ex As Exception
                Return ""
            End Try
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Read the specified text file and returns its content
    ''' </summary>
    ''' <param name="PathName">Path from wicth to read text files</param>
    ''' <returns>returns the content of the file, or null string if the file doesn't esists or is empty</returns>
    ''' <remarks></remarks>
    Public Function readTextFiles(ByVal PathName As String, Optional ByVal Encoding As String = "UTF-8") As List(Of String)
        Dim lstRetVal As List(Of String) = New List(Of String)

        For Each sFileName As String In System.IO.Directory.GetFiles(PathName)
            Debug.Print("FileSystem.readTextFiles   Loading File '" & sFileName & "'...")
            lstRetVal.Add(FileSystem.readTextFile(sFileName, Encoding))
        Next
        ' return
        Return lstRetVal
    End Function

    ''' <summary>
    ''' Writes the specified text into the specified file
    ''' </summary>
    ''' <param name="FileName">filename of the file to write</param>
    ''' <param name="Content">Content to write</param>
    ''' <param name="IncludeNewLine">true to include new line at the end of file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function writeTextFile(ByVal FileName As String, ByVal Content As String, Optional ByVal Encoding As String = "UTF-8", Optional ByVal IncludeNewLine As Boolean = False) As Boolean

        If (Content Is Nothing) Then Content = ""

        While Strings.Right(Content, 1) = Chr(10) Or Strings.Right(Content, 1) = Chr(13)
            Content = Strings.Left(Content, Content.Length - 1)
        End While

        If System.IO.File.Exists(FileName) Then FileSystem.DeleteFile(FileName)
        Try
            Dim objWriter As Global.System.IO.StreamWriter

            Select Case Encoding.ToLower().Trim()
                Case "utf7", "utf-7" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.UTF7, Text.Encoding))
                Case "utf8", "utf-8" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.UTF8, Text.Encoding))
                Case "utf16", "utf-16" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.BigEndianUnicode, Text.Encoding))
                Case "utf32", "utf-32" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.UTF32, Text.Encoding))
                Case "unicode" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.Unicode, Text.Encoding))
                Case "ascii" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.ASCII, Text.Encoding))
                Case "ansii", "ansi" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.Default, Text.Encoding))
                Case Else
                    ' ascii
                    objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(False), CType(Global.System.Text.Encoding.Default, Text.Encoding))
            End Select

            ' Appends/Not new line
            If IncludeNewLine Then
                objWriter.WriteLine(Content)
            Else
                objWriter.Write(Content)
            End If
            objWriter.Close()
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Appends the specified text into the specified file
    ''' </summary>
    ''' <param name="FileName">filename of the file to write</param>
    ''' <param name="Content">Content to write</param>
    ''' <param name="IncludeNewLine">true to include new line ad the end of file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function appendTextFile(ByVal FileName As String, ByVal Content As String, Optional ByVal Encoding As String = "utf8", Optional ByVal IncludeNewLine As Boolean = False) As Boolean

        If (Content Is Nothing) Then Content = ""

        While Strings.Right(Content, 1) = Chr(10) Or Strings.Right(Content, 1) = Chr(13)
            Content = Strings.Left(Content, Content.Length - 1)
        End While

        Try
            Dim objWriter As Global.System.IO.StreamWriter

            Select Case Encoding.ToLower().Trim()
                Case "utf7" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.UTF7, Text.Encoding))
                Case "utf8" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.UTF8, Text.Encoding))
                Case "utf16" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.BigEndianUnicode, Text.Encoding))
                Case "utf32" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.UTF32, Text.Encoding))
                Case "unicode" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.Unicode, Text.Encoding))
                Case "ascii" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.ASCII, Text.Encoding))
                Case "ansii" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.Default, Text.Encoding))
                Case "ansi" : objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.Default, Text.Encoding))
                Case Else
                    ' ascii
                    objWriter = New Global.System.IO.StreamWriter(CType(FileName,System.String), CBool(True), CType(Global.System.Text.Encoding.Default, Text.Encoding))
            End Select

            ' Appends/Not new line
            If IncludeNewLine Then
                objWriter.WriteLine(Content)
            Else
                objWriter.Write(Content)
            End If

            objWriter.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "Directory Functions"
    ' Directory Functions

    ''' <summary>
    ''' Returns true if the specified folder exists
    ''' </summary>
    ''' <param name="DirectoryName">directory to create</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FolderExists(ByVal DirectoryName As String) As Boolean
        Return System.IO.Directory.Exists(DirectoryName)
    End Function

    ''' <summary>
    ''' Creates the specified directory
    ''' </summary>
    ''' <param name="DirectoryName">directory to create</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateFolder(ByVal DirectoryName As String) As Boolean
        Try
            System.IO.Directory.CreateDirectory(DirectoryName)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Creates the specified directory tree
    ''' </summary>
    ''' <param name="DirectoryName">directory to create</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateFolderTree(ByVal DirectoryName As String) As Boolean
        Try
            Dim arrPaths() As String = Split(DirectoryName, "\")
            Dim sTmpPath As String = ""
            For Each sCurFolder As String In arrPaths
                If sCurFolder <> "" Then
                    If sTmpPath <> "" Then sTmpPath &= "\"
                    sTmpPath &= sCurFolder
                    If sCurFolder.Length = 2 And Strings.Right(sCurFolder, 1) = ":" Then
                        ' Drive path
                    Else
                        FileSystem.CreateFolder(sTmpPath)
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Deletes all files from the specified directory, and the specified folder
    ''' </summary>
    ''' <param name="DirectoryName">Directory name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteDir(ByVal DirectoryName As String) As Boolean
        Try
            Call System.IO.Directory.Delete(DirectoryName, True)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Tryes to delete all files in a diretory
    ''' </summary>
    ''' <param name="DirectoryName">Diretory to delete</param>
    ''' <param name="FileNamePrefix">Filename prefix to delete</param>
    ''' <param name="Except">File name not to delete</param>
    ''' <returns>Returns true if deletes all file or the folder is empty</returns>
    ''' <remarks></remarks>
    Public Function DeleteAllFiles(ByVal DirectoryName As String, Optional ByVal Extension As String = "*.*", Optional ByVal FileNamePrefix As String = "", Optional ByVal Except As String = "") As Boolean

        Dim lFoundedFiles As Int32 = 0
        Dim lDeletedFiles As Int32 = 0

        Extension = Extension.Trim().ToLower()
        If Extension = "" Then Extension = "*.*"
        FileNamePrefix = FileNamePrefix.Trim().ToLower().Replace(DirectoryName.ToLower(), "")
        Except = Except.Trim().ToLower().Replace(DirectoryName.ToLower(), "")

        For Each f In System.IO.Directory.GetFiles(DirectoryName, Extension, System.IO.SearchOption.TopDirectoryOnly)

            Dim bRemoveIt As Boolean = False

            ' Filters on filename
            If FileNamePrefix <> "" Then
                Dim sFindFileName As String = f.ToLower().Replace(DirectoryName.ToLower(), "")
                If Len(sFindFileName) > Len(FileNamePrefix) Then
                    bRemoveIt = (Strings.Left(sFindFileName, FileNamePrefix.Length) = FileNamePrefix)
                End If
            Else
                bRemoveIt = True
            End If

            ' Except this file?
            If f.ToLower() = DirectoryName.ToLower() & Except Then bRemoveIt = False

            If bRemoveIt Then
                lFoundedFiles += 1
                If FileSystem.DeleteFile(f) Then lDeletedFiles += 1
            End If
        Next

        ' return
        Return (lFoundedFiles = lDeletedFiles)
    End Function

    ''' <summary>
    ''' Copies the source folder into the destination folder
    ''' </summary>
    ''' <param name="SourceDirectory">source foler</param>
    ''' <param name="DestinationDirectory">destination dolfer</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CopyFolder(ByVal SourceDirectory As String, ByVal DestinationDirectory As String) As Boolean
        Try
            FileIO.FileSystem.CopyDirectory(SourceDirectory, DestinationDirectory, True)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns a list of the specified Directory subfolders
    ''' </summary>
    ''' <param name="DirectoryName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ListSubFolders(ByVal DirectoryName As String) As List(Of String)

        Dim oRetVal As List(Of String) = New List(Of String)

        DirectoryName = DirectoryName.Trim()
        If Len(DirectoryName) = 0 Then Return oRetVal ' no directory specified
        If Strings.Right(DirectoryName, 1) <> "\" Then DirectoryName &= "\"

        Try
            For Each sSubFolder As String In System.IO.Directory.GetDirectories(DirectoryName)
                sSubFolder = sSubFolder.Replace(DirectoryName, "")
                oRetVal.Add(sSubFolder)
            Next
        Catch ex As Exception
        End Try

        ' return
        Return oRetVal
    End Function

    ''' <summary>
    ''' Lists all files into the specified directory
    ''' </summary>
    ''' <param name="DirectoryName"></param>
    ''' <param name="Pattern">search pattern</param>
    ''' <param name="OrderByDate">1: order by date increasing; -1 order by date decreasing; 0:order by name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ListFiles(ByVal DirectoryName As String, Optional ByVal Pattern As String = "*.*", Optional ByVal OrderByDate As Int32 = 0) As List(Of String)

        Dim oRetVal As List(Of String) = New List(Of String)
        Dim lCounter As Int32 = 0

        ' Normalizes parameters
        DirectoryName = DirectoryName.Trim()
        If Len(DirectoryName) = 0 Then DirectoryName = AppInfo.Path
        If Strings.Right(DirectoryName, 1) <> "\" Then DirectoryName &= "\"

        Pattern = Pattern.Trim().ToLower()
        If Pattern = "" Then Pattern = "*.*"

        Select Case OrderByDate
            Case 0
            Case 1
            Case -1
            Case Else
                OrderByDate = 0
        End Select

        ' inizializza recordset per ordinamenti
        Dim oList = CreateObject("ADOR.Recordset")
        oList.Fields.Append("name", 200, 1024)
        oList.Fields.Append("date", 7)
        oList.Open

        Dim objDI = New Global.System.IO.DirectoryInfo(CType(DirectoryName,System.String))
        Try
            For Each objFI As IO.FileInfo In objDI.GetFiles(Pattern)
                Dim sFileName As String = objFI.Name
                oList.AddNew
                oList("name").Value = sFileName
                oList("date").Value = objFI.LastWriteTime
                oList.Update
                lCounter = lCounter + 1
            Next
        Catch ex As Exception
        End Try

        If lCounter > 0 Then
            ' Applica ordinamento file
            Select Case OrderByDate
                Case 0 : oList.Sort = "name"
                Case 1 : oList.Sort = "date"
                Case -1 : oList.Sort = "date desc"
            End Select

            ' Genera lista file
            Do Until oList.EOF
                oRetVal.Add(oList("name").Value)
                Call oList.MoveNext()
            Loop
            Call oList.Close()
        End If

        ' return
        Return oRetVal
    End Function

    ''' <summary>
    ''' Returns the number of files into the specified directory
    ''' </summary>
    ''' <param name="DirectoryName"></param>
    ''' <param name="Pattern">search pattern</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDirectoryFileCount(ByVal DirectoryName As String, Optional ByVal Pattern As String = "*.*") As Int32

        Dim oRetVal As List(Of String) = New List(Of String)

        DirectoryName = DirectoryName.Trim()
        If Len(DirectoryName) = 0 Then Return 0 ' no directory specified
        If Strings.Right(DirectoryName, 1) <> "\" Then DirectoryName &= "\"

        Pattern = Pattern.Trim().ToLower()
        If Pattern = "" Then Pattern = "*.*"

        Return System.IO.Directory.GetFiles(DirectoryName, Pattern).Count
    End Function

    ''' <summary>
    ''' Returns the total size of the specified directory
    ''' </summary>
    ''' <param name="DirectoryName">Directory name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDirectorySize(ByVal DirectoryName As String) As Int64
        Return (From strFile In My.Computer.FileSystem.GetFiles(DirectoryName.Trim(), FileIO.SearchOption.SearchAllSubDirectories) Select New System.IO.FileInfo(strFile).Length).Sum()
    End Function
#End Region
End Module
