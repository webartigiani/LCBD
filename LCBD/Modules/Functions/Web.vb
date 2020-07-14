Imports System.Net
''' <summary>
''' Web.vb
''' </summary>
Module Web
#Region "Public Methods"
    ' Public Methods

    ''' <summary>
    ''' Returns the content of the speficied url
    ''' </summary>
    ''' <param name="Url">url (http/https/ftp) to read</param>
    ''' <returns>the url content</returns>
    ''' <remarks>for ftp access, url must specify the required username and password</remarks>
    Public Function GetURL(ByVal Url As String) As String

        Dim sRetVal As String = ""

        Dim objStream As System.IO.Stream
        Dim objStreamReader As System.IO.StreamReader


        ' se https attiva la callback per gli eventi del gestore di certificati
        If (Strings.Left(Url.ToLower(), 8) = "https://") Then
            System.Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf SSLCertificateHandler
            System.Net.ServicePointManager.Expect100Continue = True                             ' 2017-11-16
            System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        End If

        Dim objHttWebpRequest As System.Net.HttpWebRequest
        Dim objHttpWebResponse As System.Net.HttpWebResponse

        Try
            objHttWebpRequest = System.Net.WebRequest.Create(Url)
            objHttWebpRequest.AllowAutoRedirect = True

            With objHttWebpRequest
                .UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.63 Safari/535.7"
                .ContentType = "application/x-www-form-urlencoded"
                .KeepAlive = False
                .Method = "GET"
                .Timeout = 3000
                .AutomaticDecompression = Net.DecompressionMethods.GZip

                ' legge response
                objHttpWebResponse = .GetResponse()
                If objHttpWebResponse.StatusCode = 200 Then
                    objStream = objHttpWebResponse.GetResponseStream()
                    ' nessun formato precedentemente valutato è valido: gestisce come csv
                    objStreamReader = New System.IO.StreamReader(objStream, System.Text.Encoding.UTF8)
                    sRetVal = objStreamReader.ReadToEnd()
                End If
                objHttpWebResponse.Close()
            End With

        Catch ex As Exception
            ' errore http
            Debug.Print(ex.Message)
        End Try

        ' return
        Return sRetVal
    End Function

    Public Function GetURLBin(ByVal URL As String) As Byte()

        Dim objHttWebpRequest As System.Net.HttpWebRequest
        Dim objHttpWebResponse As System.Net.HttpWebResponse
        Dim objStream As System.IO.Stream
        Dim objStreamReader As System.IO.StreamReader

        Try
            ' se https attiva la callback per gli eventi del gestore di certificati
            If (Strings.Left(URL.ToLower(), 8) = "https://") Then
                System.Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf SSLCertificateHandler
                System.Net.ServicePointManager.Expect100Continue = True                             ' 2017-11-16
                System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
            End If

            objHttWebpRequest = HttpWebRequest.Create(URL)
            objHttpWebResponse = objHttWebpRequest.GetResponse()
            objStream = objHttpWebResponse.GetResponseStream()
            Dim Buffer(4096) As Byte, BlockSize As Integer          'SourceStream has no ReadAll, so we must read data block-by-block  

            'Memory stream to store data  
            Dim TempStream As New System.IO.MemoryStream
            Do
                BlockSize = objStream.Read(Buffer, 0, 4096)
                If BlockSize > 0 Then TempStream.Write(Buffer, 0, BlockSize)
            Loop While BlockSize > 0

            'return the document binary data  
            Return TempStream.ToArray()
        Catch ex As Exception
            Debug.Print("GetURLDataBin", ex.Message, URL)
        Finally
            'grrr... Using is great, but the command is not in VB.Net  
            objStream.Close()
            objHttpWebResponse.Close()
        End Try
    End Function

    ''' <summary>
    ''' Downloads a file from the specified url and saves it in local
    ''' </summary>
    ''' <param name="DownloadUrl">Url to download</param>
    ''' <param name="SaveFileName">filename to save</param>
    ''' <returns>true in case of success</returns>
    ''' <remarks></remarks>
    Public Function DownloadFile(ByVal DownloadUrl As String, ByVal SaveFileName As String, Optional ByVal UserName As String = "", Optional ByVal Password As String = "") As Boolean
        ' default
        UserName = UserName.Trim()
        Password = Password.Trim()

        ' If https, enables TLS 1.2 callback for certificates validation event listener
        If (Strings.Left(DownloadUrl.ToLower(), 8) = "https://") Then
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Global.System.Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf SSLCertificateHandler
        End If

        Dim oClient = New Global.System.Net.WebClient()
        Try
            Using oClient
                If UserName <> "" And Password <> "" Then oClient.Credentials = New Global.System.Net.NetworkCredential(CType(UserName,System.String), CType(Password,System.String))
                oClient.DownloadFile(DownloadUrl, SaveFileName)
                Return True
            End Using
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Clears Internet Explorer temporary files
    ''' </summary>
    Public Sub ClearTemporaryInternetFiles()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 8")
        Shell.ShellAndWait("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 8", True,, ProcessWindowStyle.Hidden)
    End Sub

    ''' <summary>
    ''' Clears Internet Explorer cookies
    ''' </summary>
    Public Sub ClearCookies()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2")
        Shell.ShellAndWait("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2", True,, ProcessWindowStyle.Hidden)
    End Sub

    ''' <summary>
    ''' Clears Internet Explorer Hostiry
    ''' </summary>
    Public Sub ClearHistory()
        'System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1")
        Shell.ShellAndWait("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1", True,, ProcessWindowStyle.Hidden)
    End Sub

    ''' <summary>
    ''' ExtractsLinks from an html code and returns thems into a list of string
    ''' </summary>
    ''' <param name="HTML">HTML code from wich extract links</param>
    ''' <param name="Links">returned list of extracted links</param>
    ''' <returns></returns>
    ''' <remarks>Returns list of extracted links</remarks>
    Public Function ExtractLinksFromHTML(ByVal HTML As String, ByRef Links As List(Of String), Optional ByVal Distinct As Boolean = True) As Boolean

        Dim bRetVal As Boolean = False
        Dim sLastLink As String = ""
        Links = New List(Of String)

        Dim colMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(HTML, "<a[^>]*>")
        If colMatches.Count > 0 Then

            For Each objMatch As System.Text.RegularExpressions.Match In colMatches
                Dim sLink As String = objMatch.Value.ToString()
                sLink = sLink.ExtractSubStringBetween("href=""", Chr(34)).ToLower()

                If sLink <> "" Then
                    If sLink <> sLastLink Then Links.Add(sLink)
                    sLastLink = sLink
                    bRetVal = True
                End If
            Next
        End If

        If (Distinct) Then Links = Links.Distinct().ToList() ' esegue distinct sulla lista
        Return bRetVal
    End Function
#End Region

#Region "Private Members"
    ' Private Members

    ''' <summary>
    ''' Gestore di eventi del gestore dei certificati
    ''' </summary>
    ''' <returns>restituisce sempre true per by-passare eventuali errori di certificato</returns>
    ''' <remarks></remarks>
    Private Function SSLCertificateHandler(ByVal sender As Object, ByVal certificate As Global.System.Security.Cryptography.X509Certificates.X509Certificate _
        , ByVal chain As Global.System.Security.Cryptography.X509Certificates.X509Chain, ByVal SSLerror As Global.System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function
#End Region
End Module