Public Class ClassTripFile
    Private _fullpath As String

    Public Sub New(ByVal strFullPath As String)
        _fullpath = strFullPath
    End Sub

    Public ReadOnly Property FullPath() As String
        Get
            Return _fullpath
        End Get
    End Property

    Public ReadOnly Property Extension() As String
        Get
            Return System.IO.Path.GetExtension(_fullpath)
        End Get
    End Property

    Public ReadOnly Property Filename() As String
        Get
            Return System.IO.Path.GetFileName(_fullpath)
        End Get
    End Property

    Public ReadOnly Property FilenameWithoutExtension() As String
        Get
            Return System.IO.Path.GetFileNameWithoutExtension(_fullpath)
        End Get
    End Property
End Class
