''' <summary>
''' LCBD - Local Carrier Bulk Downloader
''' WebArtigiani
''' </summary>
''' 
Public Class frmMain
#Region "Form Events"

    ' Form Events
    ''' <summary>
    ''' Load
    ''' </summary>
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim sFileCSV As String = Registry.ReadSetting("settings", "input-csv")
        Dim sOutputFolder As String = Registry.ReadSetting("settings", "output-folder")

        With Me
            .Text = AppInfo.AppName & " v." & AppInfo.AppVersion
            .picHeader.BackColor = ColorTranslator.FromHtml("#0580ac")
            .picLogo.BackColor = ColorTranslator.FromHtml("#0580ac")
            .picBottom.BackColor = ColorTranslator.FromHtml("#0580ac")

            .lnk1.BackColor = ColorTranslator.FromHtml("#0580ac")
            .lnk2.BackColor = ColorTranslator.FromHtml("#0580ac")
            .pnlWip.Visible = False
            .lblWip.Text = ""
            .btnStop.Left = .btnExecute.Left
            .btnStop.Top = .btnExecute.Top
            .btnStop.Visible = False

            .txtFileName.Text = sFileCSV
            .txtSaveFolder.Text = sOutputFolder
        End With
    End Sub

    ''' <summary>
    ''' Browse for CSV file
    ''' </summary>
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click

        With Me.OpenFileDialog1
            .Filter = "File CSV (*.csv)|*.csv"
            .Title = "Seleziona il file CSV da elaborare"

            Dim r As DialogResult = .ShowDialog()
            If r = DialogResult.OK Then
                Dim sInput As String = .FileName

                txtFileName.Text = sInput
                Call Registry.WriteSetting("settings", "input-csv", sInput)
            End If
        End With
    End Sub

    ''' <summary>
    ''' Browse for output folder
    ''' </summary>
    Private Sub btnOutputFolder_Click(sender As Object, e As EventArgs) Handles btnOutputFolder.Click
        ' output-folder"

        With Me.FolderBrowserDialog1
            .Description = "Seleziona la directory in cui scaricare le immagini"

            Dim r As DialogResult = .ShowDialog()
            If r = DialogResult.OK Then
                Dim sInput As String = .SelectedPath

                txtSaveFolder.Text = sInput
                Call Registry.WriteSetting("settings", "output-folder", sInput)
            End If
        End With
    End Sub

    ''' <summary>
    ''' Executes
    ''' </summary>
    Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click

        Dim sFile As String = txtFileName.Text.Trim()
        Dim sOutputFolder As String = txtSaveFolder.Text.Trim()
        Dim sBaseUrl As String = "https://logo.clearbit.com/"
        Dim sErrorFile As String = "errors.txt"

        Dim iLine As Int32 = 0
        Dim lTotal As Int32 = 0, lDone As Int32 = 0, lErrors As Int32 = 0


        If sFile = "" Then
            Call MessageBox.Show("Prego, selezionare il file CSV da elaborare.", "File mancante", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        If sOutputFolder = "" Then
            Call MessageBox.Show("Prego, selezionare la directory di destinazione.", "Directory mancante", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        If Not FileSystem.FileExists(sFile) Then
            Call MessageBox.Show("Impossibile trovare il file '" & sFile & "'.", "File non trovato", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        If Not FileSystem.FolderExists(sOutputFolder) Then
            Call MessageBox.Show("Impossibile trovare la directory '" & sOutputFolder & "'.", "Percorso non trovato", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        If Strings.Right(sOutputFolder, 1) <> "\" Then sOutputFolder &= "\"


        ' Checks if output folder is empty
        If FileSystem.ListFiles(sOutputFolder).Count > 0 Then
            If MessageBox.Show("La directory '" & sOutputFolder & "' non è vuota." & vbCrLf & "Si desidera procedere comunque?", "Conferma Directory", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Return
        End If


        ' Reads the source file
        Dim sFileCnt As String = FileSystem.readTextFile(sFile).Trim()
        If sFileCnt = "" Then
            Call MessageBox.Show("Il file selezionato non contiene righe." & vbCrLf & "Prego, selezionare un'altro file CSV.", "File vuoto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        Dim arrLines() As String = Split(sFileCnt, vbCrLf)          ' Splits file lines
        If arrLines.Count = 0 Then
            Call MessageBox.Show("Il file selezionato non contiene righe." & vbCrLf & "Prego, selezionare un'altro file CSV.", "File vuoto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        ' Deletes errors.txt
        Call FileSystem.DeleteFile(sErrorFile)


        ' Executes
        With Me
            .Cursor = Cursors.WaitCursor
            .btnBrowse.Enabled = False
            .btnOutputFolder.Enabled = False
            .btnExecute.Enabled = False
            .lblWip.Text = "Attendere..."
            .pgbWip.Minimum = 0
            .pgbWip.Maximum = arrLines.Count
            .pgbWip.Value = 0
            .pnlWip.Visible = True
            .btnExecute.Visible = False
            .btnStop.Tag = ""
            .btnStop.Visible = True


            For Each sLine As String In arrLines

                ' Checks if stop
                If btnStop.Tag <> "" Then Exit For

                Application.DoEvents()
                iLine += 1

                With Me
                    .pgbWip.Value = iLine
                    .lblWip.Text = "Lettura riga " & iLine.ToString() & " di " & arrLines.Count.ToString & "..."
                End With

                sLine = sLine.Trim()
                If sLine <> "" And iLine > 1 Then

                    Console.Write(iLine.ToString() & " ")
                    Dim arrFields() As String = Split(sLine, ",")
                    Dim sCarrier As String = arrFields(0).Replace(Chr(34), "").Trim()
                    Dim sCountry As String = arrFields(1).Trim().ToLower()

                    If sCarrier <> "" And sCountry <> "" Then
                        lTotal += 1

                        Dim sUrl As String = sBaseUrl & sCarrier & "." & sCountry
                        Dim sImgFile As String = sOutputFolder & sCarrier & "." & sCountry & ".png"

                        sUrl = sUrl.Replace("&", "")
                        sImgFile = sImgFile.Replace("&", "").Replace("/", "")

                        ' Downloads file
                        If Not FileSystem.FileExists(sImgFile) Then
                            If Web.DownloadFile(sUrl, sImgFile) Then
                                lDone += 1
                            Else
                                Call Application.DoEvents()
                                If Web.DownloadFile(sBaseUrl & sCarrier.ToLower().Replace("&", "") & ".com", sImgFile) Then
                                    lDone += 1
                                Else

                                    ' Creates Mockup Image
                                    ' ex https://via.placeholder.com/128x128.png?text=TIM%20II
                                    Call Application.DoEvents()
                                    sUrl = "https://via.placeholder.com/128x128.png?text=" & sCarrier.UrlEncode()
                                    If Web.DownloadFile(sUrl, sImgFile) Then
                                        lDone += 1
                                    Else
                                        ' Error
                                        lErrors += 1
                                        Call FileSystem.appendTextFile(sErrorFile, sUrl & " (url not found)",, True)
                                    End If


                                    ' Checks downloaded file
                                    Call Application.DoEvents()
                                    If FileSystem.FileExists(sImgFile) Then
                                        If FileSystem.GetFileSize(sImgFile) = 0 Then
                                            ' File is empty
                                            lErrors += 1
                                            lDone -= 1
                                            Call FileSystem.appendTextFile(sErrorFile, sUrl & " (0 byte file)",, True)
                                            Call FileSystem.DeleteFile(sImgFile)
                                        Else

                                            ' Creates Mockup Image
                                            ' ex https://via.placeholder.com/128x128.png?text=TIM%20II
                                            sUrl = "https://via.placeholder.com/128x128.png?text=" & sCarrier.Replace(" ", "%20")
                                            If Web.DownloadFile(sUrl, sImgFile) Then
                                                lDone += 1
                                            Else
                                                ' Error
                                                lErrors += 1
                                                Call FileSystem.appendTextFile(sErrorFile, sUrl & " (url not found)",, True)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            If lDone < 0 Then lDone = 0
            If lErrors > 0 Then
                If MessageBox.Show("Scaricate " & lDone.ToString() & " immagini di " & lTotal.ToString() & "." & vbCrLf & vbCrLf & "Si desidera visualizzare il log degli errori per le " & lErrors.ToString() & " immagini non scaricate?", "Finito", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                    Process.Start(sErrorFile)
                End If
            Else
                Call MessageBox.Show("Scaricate " & lDone.ToString() & " immagini di " & lTotal.ToString() & ".", "Finito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            .btnStop.Tag = ""
            .btnStop.Visible = False
            .btnExecute.Visible = True
            .btnBrowse.Enabled = True
            .btnOutputFolder.Enabled = True
            .btnExecute.Enabled = True
            .pnlWip.Visible = False
            .Cursor = Cursors.Default
        End With
    End Sub

    ''' <summary>
    ''' Stops
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        If MessageBox.Show("Confermi di volere interrompere il download delle immagini?", "Conferma interruzione", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then btnStop.Tag = "Y"
    End Sub

    ''' <summary>
    ''' Open link
    ''' </summary>
    Private Sub lnk1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnk1.LinkClicked, lnk2.LinkClicked

        Dim lnk As LinkLabel = CType(sender, LinkLabel)
        Dim sUrl As String = lnk.Tag
        Process.Start(sUrl)
    End Sub
#End Region
End Class
