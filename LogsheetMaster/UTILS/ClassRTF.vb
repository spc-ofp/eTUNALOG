Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.IO
Imports System.Diagnostics

Public Class MyRtb
    Inherits RichTextBox

    <Editor(GetType(RtfEditor), GetType(UITypeEditor))> _
    Public Property RichText() As String
        Get
            Return MyBase.Rtf
        End Get
        Set(ByVal value As String)
            MyBase.Rtf = value
        End Set
    End Property

End Class

Friend Class RtfEditor
    Inherits UITypeEditor

    Public Overrides Function GetEditStyle(ByVal context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return UITypeEditorEditStyle.Modal
    End Function

    Public Overrides Function EditValue(ByVal context As ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
        Dim fname As String = Path.Combine(Path.GetTempPath, "text.rtf")
        File.WriteAllText(fname, CStr(value))
        Process.Start("wordpad.exe", fname).WaitForExit()
        value = File.ReadAllText(fname)
        File.Delete(fname)
        Return value
    End Function
End Class
