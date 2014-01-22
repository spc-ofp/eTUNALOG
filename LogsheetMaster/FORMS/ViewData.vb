Imports System.IO
Imports System.Xml
Imports System.Data.SQLite


Public Class ViewData

    Dim conn As New OleDb.OleDbConnection

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        conn = New OleDb.OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Application.StartupPath & "\Logsheet_XML_Import.accdb"
        '
        'get data into list
        Me.RefreshData()
    End Sub


    Private Sub RefreshData()
        If Not conn.State = ConnectionState.Open Then
            'open connection
            conn.Open()
        End If

        Dim da As New OleDb.OleDbDataAdapter("SELECT FirstName, LastName FROM test ORDER BY FirstName", conn)
        Dim dt As New DataTable
        'fill data to datatable
        da.Fill(dt)

        'offer data in data table into datagridview
        Me.DataGridView1.DataSource = dt

        'close connection
        conn.Close()
    End Sub


    Private Sub CmdInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdInsert.Click
        Dim cmd As New OleDb.OleDbCommand

        If Not conn.State = ConnectionState.Open Then
            'open connection if it is not yet open
            conn.Open()
        End If

        cmd.Connection = conn
        'check whether add new or update
        If Me.txtFirstname.Text & "" <> "" Then
            'add new 
            'add data to table
            cmd.CommandText = "INSERT INTO test(FirstName, LastName, Address) " & _
                            " VALUES('" & Me.txtFirstname.Text & "','" & Me.txtLastname.Text & "','" & Me.txtAddress.Text & "')"

            cmd.ExecuteNonQuery()
        End If

        'refresh data in list
        RefreshData()
        'clear form
        'Me.btnClear.PerformClick()

        'close connection
        conn.Close()

    End Sub

    Private Sub CmdImportXml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdImportXml.Click
        ParseXML()
    End Sub

    Private Sub ParseXML()
        Try
            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode

            'Create the XML Document
            m_xmld = New XmlDocument()
            m_xmld.Load("C:\TMP\test.xml")

            'Get the list of name nodes 
            'm_nodelist = m_xmld.SelectNodes("/PSLOGSHEET2009/LOGSHEET/WholePage/Trip_Header")
            m_nodelist = m_xmld.SelectNodes("/PSLOGSHEET2009/LOGSHEET/WholePage/Daily_Log/Daily_Details/DAILYLOG/Day")
            'Loop through the nodes
            For Each m_node In m_nodelist
                'Get the Gender Attribute Value
                'Dim genderAttribute = m_node.Attributes.GetNamedItem("gender").Value
                Dim val1 = m_node.ChildNodes.Item(0).InnerText
                Dim val2 = m_node.ChildNodes.Item(1).InnerText
                MsgBox("1: " & val1 & " 2: " & val2)
            Next
        Catch errorVariable As Exception
            'Error trapping
            MsgBox(errorVariable.ToString(), MsgBoxStyle.OkOnly)
        End Try
    End Sub

End Class