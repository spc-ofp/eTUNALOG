Public Class ClassTemplateFilesList

    Private m_TempFiles(-1) As ClassTripFile
    Private m_count As Integer

    Public Sub New()
        m_count = 0
    End Sub

    Public Function Item(ByVal Index As Integer) As ClassTripFile
        If (Index + 1) <= m_count Then
            Item = m_TempFiles(Index)
        Else
            ' raise error
            Item = Nothing
        End If
    End Function

    Public ReadOnly Property Count() As Integer
        Get
            Return m_count
        End Get
    End Property

    Public Sub AddItem(ByVal strFilename As String)
        m_count = m_count + 1
        ReDim Preserve m_TempFiles(m_count - 1)
        m_TempFiles(m_count - 1) = New ClassTripFile(strFilename)
    End Sub

    Public Sub AddItemsToComboBox(ByVal oComboBox As ComboBox)
        Dim i As Integer

        oComboBox.Items.Clear()

        For i = 1 To m_count
            oComboBox.Items.Add(m_TempFiles(i - 1).Filename)
        Next
    End Sub

    Public Sub Clear()
        ReDim m_TempFiles(-1)
        m_count = 0
    End Sub

End Class
