''' <summary>
''' SystemInfo.vb
''' </summary>
Module SystemInfo
#Region "Public Properties"
    ' Public Methods

    ''' <summary>
    ''' Returns true if we're in a 64 bit operating System
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Is64Bit() As Boolean
        Get
            Return Global.System.Environment.Is64BitOperatingSystem
        End Get
    End Property

    ''' <summary>
    ''' Returns true if we're in a 32 bit operating System
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Is32Bit() As Boolean
        Get
            Return Not Global.System.Environment.Is64BitOperatingSystem
        End Get
    End Property

    Public ReadOnly Property TimeStamp() As Int64
        Get
            Return CLng(DateTime.UtcNow.Subtract(New DateTime(1970, 1, 1)).TotalMilliseconds)
        End Get
    End Property

#End Region
End Module