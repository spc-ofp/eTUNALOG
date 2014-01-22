Imports FDFApp
Imports FDFApp.FDFApp_Class
Imports FDFApp.FDFDoc_Class

Public Class ClassFDF

    Public Sub New()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        'cFDFDoc = cFDFApp.FDFOpenFromFile("C:\tmp\Logsheet_PS_2009.pdf", cFDFApp.FDFType.PDF)
        cFDFDoc = cFDFApp.PDFOpenFromFile("C:\tmp\Logsheet_PS_2009.pdf")

    End Sub


End Class
