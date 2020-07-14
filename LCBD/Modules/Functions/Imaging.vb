Imports System.Drawing
Imports System.IO
''' <summary>
''' Imaging.vb
''' </summary>
Module Imaging

#Region "Public Members"
    ' Public Members

    ''' <summary>
    ''' Returns true if the specified file is an image
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <returns></returns>
    Public Function IsImage(ByVal FileName As String) As Boolean
        Try
            Using fs As New FileStream(FileName, FileMode.Open, FileAccess.Read)
                Using b As New Bitmap(fs)
                    b.Dispose()
                End Using
                fs.Close()
            End Using
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns the size of the specified image filename
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <param name="Width"></param>
    ''' <param name="Height"></param>
    ''' <returns></returns>
    Public Function GetImageSize(ByVal FileName As String, ByRef Width As Int32, ByRef Height As Int32) As Boolean

        ' default
        Width = 0
        Height = 0
        Try
            Using fs As New FileStream(FileName, FileMode.Open, FileAccess.Read)
                Using b As New Bitmap(fs)
                    Width = b.Width
                    Height = b.Height
                    b.Dispose()
                End Using
                fs.Close()
            End Using
            Return True
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Resizes image from file, and saves it as new file
    ''' </summary>
    ''' <param name="SourceFileName"></param>
    ''' <param name="SaveFileName"></param>
    ''' <param name="TargetWidth"></param>
    ''' <param name="TargetHeight"></param>
    ''' <returns></returns>
    Public Function ResizeImageFile(ByVal SourceFileName As String, ByVal SaveFileName As String, ByVal TargetWidth As Int32, ByVal TargetHeight As Int32) As Boolean
        Try
            Using fs As New FileStream(SourceFileName, FileMode.Open, FileAccess.Read)
                Using b As New Bitmap(fs)
                    Dim oDestBmp As Drawing.Bitmap = ResizeImage(b, TargetWidth, TargetHeight)     ' resizes the bitmap
                    oDestBmp.Save(SaveFileName)
                    oDestBmp.Dispose()
                    b.Dispose()
                End Using
                fs.Close()
            End Using

            'Dim oSourceImage As Image = Image.FromFile(SourceFileName)                              ' loads source-filename
            'Dim oSourceBmp = New Drawing.Bitmap(oSourceImage)                                       ' converts it into a bitmap
            'Dim oDestBmp As Drawing.Bitmap = ResizeImage(oSourceBmp, TargetWidth, TargetHeight)     ' resizes the bitmap
            'oDestBmp.Save(SaveFileName)
            'oSourceImage.Dispose()
            'oSourceBmp.Dispose()
            'oDestBmp.Dispose()

            'oSourceImage = Nothing
            'oSourceBmp = Nothing
            'oDestBmp = Nothing
            Return True
        Catch ex As Exception
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Resizes an image from Bitmap and returns it as new Bitmap
    ''' </summary>
    ''' <param name="bmSource"></param>
    ''' <param name="TargetWidth"></param>
    ''' <param name="TargetHeight"></param>
    ''' <returns></returns>
    Public Function ResizeImage(ByVal bmSource As Drawing.Bitmap, ByVal TargetWidth As Int32, ByVal TargetHeight As Int32) As Drawing.Bitmap

        Dim bmDest As New Drawing.Bitmap(TargetWidth, TargetHeight, Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim nSourceAspectRatio = bmSource.Width / bmSource.Height
        Dim nDestAspectRatio = bmDest.Width / bmDest.Height

        Dim NewX = 0
        Dim NewY = 0
        Dim NewWidth = bmDest.Width
        Dim NewHeight = bmDest.Height

        If nDestAspectRatio = nSourceAspectRatio Then
            'same ratio
        ElseIf nDestAspectRatio > nSourceAspectRatio Then
            'Source is taller
            NewWidth = Convert.ToInt32(Math.Floor(nSourceAspectRatio * NewHeight))
            NewX = Convert.ToInt32(Math.Floor((bmDest.Width - NewWidth) / 2))
        Else
            'Source is wider
            NewHeight = Convert.ToInt32(Math.Floor((1 / nSourceAspectRatio) * NewWidth))
            NewY = Convert.ToInt32(Math.Floor((bmDest.Height - NewHeight) / 2))
        End If

        Using grDest = Drawing.Graphics.FromImage(bmDest)
            With grDest
                .CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
                .InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = Drawing.Drawing2D.PixelOffsetMode.HighQuality
                .SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
                .CompositingMode = Drawing.Drawing2D.CompositingMode.SourceOver

                ' Sets a white background (DOESN'T WORK !!!)
                'Dim oBlankBmp As Drawing.Bitmap = New Drawing.Bitmap(TargetWidth, TargetHeight)
                'Dim gr As Graphics = Graphics.FromImage(oBlankBmp)
                'gr.DrawLine(Pens.White, 0, 0, TargetWidth, TargetHeight)
                '.DrawImage(oBlankBmp, 0, 0, TargetWidth, TargetHeight)
                '.Save()

                .DrawImage(bmSource, NewX, NewY, NewWidth, NewHeight)
                .Save()
            End With
        End Using

        Return bmDest
    End Function

    ''' <summary>
    ''' Converts a Base64 string into an image
    ''' </summary>
    ''' <param name="base64string"></param>
    ''' <returns></returns>
    Function Base64ToBitmap(ByVal base64string As String) As Global.System.Drawing.Bitmap

        ' Setup image and get data stream together
        Dim b64 As String = base64string.Replace(" ", "+")
        Dim b() As Byte
        b = Convert.FromBase64String(b64)                   ' Converts the base64 encoded msg to image data
        Dim MS As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream(CType(b,System.Byte()))
        Dim img As Global.System.Drawing.Image = Global.System.Drawing.Image.FromStream(MS)

        Return New Bitmap(img)
    End Function

    ''' <summary>
    ''' Converts a Base64 string into an image
    ''' </summary>
    ''' <param name="base64string"></param>
    ''' <returns></returns>
    Function Base64ToImage(ByVal base64string As String) As Global.System.Drawing.Image

        ' Setup image and get data stream together
        Dim b64 As String = base64string.Replace(" ", "+")
        Dim b() As Byte
        b = Convert.FromBase64String(b64)                   ' Converts the base64 encoded msg to image data
        Dim MS As Global.System.IO.MemoryStream = New Global.System.IO.MemoryStream(CType(b,System.Byte()))
        Dim objRetIMG As Global.System.Drawing.Image = Global.System.Drawing.Image.FromStream(MS)

        Return objRetIMG
    End Function
#End Region
End Module
