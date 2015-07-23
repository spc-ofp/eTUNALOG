Imports System.IO
Imports System.Net.Mail
Imports FDFApp
Imports FDFApp.FDFApp_Class
Imports FDFApp.FDFDoc_Class

Public Class FormMain
    Dim TripFileslist As New ClassTripFileslist
    Dim CommentsFileslist As New ClassTripFileslist
    Dim FADFileslist As New ClassTripFileslist

    Dim TemplateFileslist As New ClassTemplateFilesList
    Dim LLTemplateFileslist As New ClassTemplateFilesList
    Dim FADTemplateFileslist As New ClassTemplateFilesList
    Dim AboutTabNumber As Integer

    Private Sub FormMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Logsheet_XML_ImportDataSet.test' table. You can move, or remove it, as needed.
        'Me.TestTableAdapter.Fill(Me.Logsheet_XML_ImportDataSet.test)

        CheckLanguage()
        CheckFADTAB(False)
        CheckLOGSHEETTAB(False)

        ' temporarily removes the Testing TAB
        TabControl1.Controls.Remove(TabControl1.TabPages(5))

        ' removing TABs according to option settings
        If CheckShowFAD.Checked = False Then
            TabControl1.Controls.Remove(TabControl1.TabPages(2))
        End If
        If CheckShowLogsheet.Checked = False Then
            TabControl1.Controls.Remove(TabControl1.TabPages(1))
        End If

        AboutTabNumber = 3

        ReadExistingTripFiles()
        ReadExistingFADFiles()
        ReadTemplateFiles_PS()
        ReadTemplateFiles_LL()
        ReadFADTemplateFiles()

    End Sub

    Public Sub ReadExistingTripFiles()
        ' reads trip files in the 'trips' folder and adds to the existing trips list
        Dim tripFilesPath As String
        Dim tripFiles As String()

        TripFileslist.Clear()

        Try
            tripFilesPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\Trips"

            If Directory.Exists(tripFilesPath) Then
                tripFiles = Directory.GetFiles(tripFilesPath)

                For Each tripFile In tripFiles
                    If Path.GetExtension(tripFile) = ".pdf" Then
                        TripFileslist.AddItem(tripFile)
                    End If
                Next

                FillTripList()

            Else
                Directory.CreateDirectory(tripFilesPath)
            End If
        Catch Ex As Exception
            MessageBox.Show("There was an error reading the existing trip files :-\")
        End Try
    End Sub

    Public Sub ReadExistingFADFiles()
        ' reads FAD files in the 'FAD_Reports' folder and adds to the existing FAD file list
        Dim FADFilesPath As String
        Dim FADFiles As String()

        FADFileslist.Clear()

        Try
            FADFilesPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\FAD_Reports"

            If Directory.Exists(FADFilesPath) Then
                FADFiles = Directory.GetFiles(FADFilesPath)

                For Each FADFile In FADFiles
                    If Path.GetExtension(FADFile) = ".pdf" Then
                        FADFileslist.AddItem(FADFile)
                    End If
                Next

                FillFADList()

            Else
                Directory.CreateDirectory(FADFilesPath)
            End If
        Catch Ex As Exception
            MessageBox.Show("There was an error reading the existing FAD files :-\")
        End Try
    End Sub

    Public Sub ReadTemplateFiles_PS()
        ' reads all template PDFs files in the 'template' folder and adds to the existing template combobox
        Dim templateFilesPath As String
        Dim templateFiles As String()

        TemplateFileslist.Clear()

        Try
            templateFilesPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\Template"

            If Directory.Exists(templateFilesPath) Then
                templateFiles = Directory.GetFiles(templateFilesPath)

                For Each tripFile In templateFiles
                    If Path.GetExtension(tripFile) = ".pdf" And tripFile.Contains("PS_") Then
                        TemplateFileslist.AddItem(tripFile)
                    End If
                Next
                FillTemplateList()
            Else
                Directory.CreateDirectory(templateFilesPath)
            End If
        Catch Ex As Exception
            MessageBox.Show("There was an error reading the existing template files :-\")
        End Try
    End Sub

    Public Sub ReadTemplateFiles_LL()
        ' reads all template PDFs files in the 'template' folder and adds to the existing template combobox
        Dim templateFilesPath As String
        Dim templateFiles As String()

        LLTemplateFileslist.Clear()

        Try
            templateFilesPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\Template"

            If Directory.Exists(templateFilesPath) Then
                templateFiles = Directory.GetFiles(templateFilesPath)

                For Each tripFile In templateFiles
                    If Path.GetExtension(tripFile) = ".pdf" And tripFile.Contains("LL_") Then
                        LLTemplateFileslist.AddItem(tripFile)
                    End If
                Next
                FillLLTemplateList()
            Else
                Directory.CreateDirectory(templateFilesPath)
            End If
        Catch Ex As Exception
            MessageBox.Show("There was an error reading the existing template files :-\")
        End Try
    End Sub

    Public Sub ReadFADTemplateFiles()
        ' reads all template PDFs files in the 'template' folder and adds to the existing template combobox
        Dim FADtemplateFilesPath As String
        Dim FADtemplateFiles As String()

        FADTemplateFileslist.Clear()

        Try
            FADtemplateFilesPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\Template"

            If Directory.Exists(FADtemplateFilesPath) Then
                FADtemplateFiles = Directory.GetFiles(FADtemplateFilesPath)

                For Each FADFile In FADtemplateFiles
                    If Path.GetExtension(FADFile) = ".pdf" And FADFile.Contains("FAD_") Then
                        FADTemplateFileslist.AddItem(FADFile)
                    End If
                Next
                FillFADTemplateList()
            Else
                Directory.CreateDirectory(FADtemplateFilesPath)
            End If
        Catch Ex As Exception
            MessageBox.Show("There was an error reading the existing template files :-\")
        End Try
    End Sub


    Public Sub FillTripList()
        TripFileslist.AddItemsToListBox(ListBoxTrips)
    End Sub

    Public Sub FillFADList()
        FADFileslist.AddItemsToListBox(ListBoxFAD)
    End Sub

    Public Sub FillTemplateList()
        TemplateFileslist.AddItemsToComboBox(ComboBoxTemplate)
    End Sub

    Public Sub FillLLTemplateList()
        LLTemplateFileslist.AddItemsToComboBox(ComboBoxLLTemplate)
    End Sub

    Public Sub FillFADTemplateList()
        FADTemplateFileslist.AddItemsToComboBox(ComboBoxFADTemplate)
    End Sub


    Private Sub ButtonOpenTrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenTrip.Click
        If ListBoxTrips.SelectedIndex < 0 Then
            MessageBox.Show("Please select a trip first...")
        Else
            'MessageBox.Show("You are opening " + TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
            Try
                System.Diagnostics.Process.Start(TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the file :-\")
            End Try
        End If
    End Sub


    Private Sub ButtonNewTripFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNewTripFile.Click
        Dim strNewFilename As String, strNewPath As String, strTemplate As String
        Dim StrGearType As String

        If RadioSeiner.Checked = True Then
            StrGearType = "PS"
        ElseIf RadioLongliner.Checked = True Then
            StrGearType = "LL"
        Else
            StrGearType = ""
        End If


        ' Check that we have the basic parameters - trip date, vesselname, ffavid
        If TextBoxVesselname.TextLength = 0 Or TextBoxFFAVID.TextLength = 0 Then
            MessageBox.Show("Please complete the VESSEL PARAMETERS before opening a new trip...")
            TabControl1.SelectTab(AboutTabNumber)
            TextBoxVesselname.Focus()
        ElseIf RadioSeiner.Checked = False And RadioLongliner.Checked = False Then
            MessageBox.Show("Please complete the VESSEL GEAR TYPE parameters before opening a new trip...")
            TabControl1.SelectTab(AboutTabNumber)
            RadioSeiner.Focus()
        ElseIf ComboBoxTemplate.SelectedIndex < 0 And RadioSeiner.Checked = True Then
            MessageBox.Show("Please select a PURSE SEINE LOGSHEET TEMPLATE before creating a new trip...")
            TabControl1.SelectTab(AboutTabNumber)
            ComboBoxTemplate.Focus()
        ElseIf ComboBoxLLTemplate.SelectedIndex < 0 And RadioLongliner.Checked = True Then
            MessageBox.Show("Please select a LONGLINE LOGSHEET TEMPLATE before creating a new trip...")
            TabControl1.SelectTab(AboutTabNumber)
            ComboBoxLLTemplate.Focus()
        Else
            strNewFilename = TextBoxVesselname.Text.Trim + "_" + StrGearType + "_" + TextBoxFFAVID.Text.Trim + IIf(TripNumber.TextLength <> 0, "_" + Trim(TripNumber.Text), "") + "_" + Format(DateTimePickerTripStartDate.Value, "yyyy-MM-dd") + ".pdf"
            strNewPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\trips\" + strNewFilename
            If RadioSeiner.Checked = True Then
                strTemplate = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\template\" + ComboBoxTemplate.SelectedItem
            Else
                strTemplate = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\template\" + ComboBoxLLTemplate.SelectedItem
            End If

            If System.IO.File.Exists(strNewPath) Then
                MessageBox.Show("A file with this name already exists...")
            ElseIf Not System.IO.File.Exists(strTemplate) Then
                MessageBox.Show("The logsheet template file could not be found in the template folder")
            Else
                If MessageBox.Show("Do you want to create " + strNewFilename + " ?", "TRIP CREATION", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    System.IO.File.Copy(strTemplate, strNewPath)
                    System.IO.File.SetAttributes(strNewPath, FileAttribute.Normal)

                    'MessageBox.Show("open file '" + strNewPath + "' here")
                    Try
                        System.Diagnostics.Process.Start(strNewPath)
                    Catch ex As Exception
                        MessageBox.Show("There was an error opening the file :-\")
                    End Try
                    ReadExistingTripFiles()
                End If
            End If
        End If
    End Sub

    Private Sub ListBoxTrips_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBoxTrips.DoubleClick
        If ListBoxTrips.SelectedIndex >= 0 Then
            Try
                System.Diagnostics.Process.Start(TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the file :-\")
            End Try
        End If
    End Sub
   
    Private Sub ComboBoxTemplate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'MessageBox.Show(ComboBoxTemplate.SelectedItem)
    End Sub

    Private Sub LinkLabelManu_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabelManu.LinkClicked
        System.Diagnostics.Process.Start("mailto:" + "emmanuels@spc.int" + "?subject=SPC%20Purse%20Seine%20Electronic%20Logsheet%20(pdf)")
    End Sub

    Private Sub ButtonDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDelete.Click
        If ListBoxTrips.SelectedIndex < 0 Then
            MessageBox.Show("Please select a trip to DELETE first...")
        Else
            'MessageBox.Show("You are opening " + TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
            Try
                If MessageBox.Show("Do you want to DELETE trip '" + ListBoxTrips.SelectedItem + "' ?", "TRIP DELETION", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    System.IO.File.Delete(TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
                    ReadExistingTripFiles()
                End If
            Catch ex As Exception
                MessageBox.Show("There was an error deleting the file :-\")
            End Try
        End If

    End Sub

    Private Sub TextBoxVesselname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBoxVesselname.Text.ToUpper()
        TextBoxVesselname.Refresh()
    End Sub

    Private Sub LinkLabelCopyright_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabelCopyright.LinkClicked
        System.Diagnostics.Process.Start("http://www.spc.int/oceanfish")
    End Sub

    Private Sub ButtonSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSendEmail.Click

        ' Check that we have the basic parameters - trip date, vesselname, ffavid
        If TextBoxCompanyEmail.TextLength = 0 Then
            MessageBox.Show("Please complete the company's email for emailing the trip...")
            TabControl1.SelectTab(AboutTabNumber)
            TextBoxCompanyEmail.Focus()
        ElseIf ListBoxTrips.SelectedIndex < 0 Then
            MessageBox.Show("Please select a trip first...")
        Else
            'MessageBox.Show("You are opening " + TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
            Try
                If MessageBox.Show("Do you want to send the trip '" + ListBoxTrips.SelectedItem + "' to " + TextBoxCompanyEmail.Text + " ?", "EMAIL TRIP", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    LabelWait.Visible = True
                    Me.Refresh()
                    MailHelper.SendMailMessage("ofpweb@gmail.com", TextBoxCompanyEmail.Text, "", "emmanuels@spc.int", "Smart purse seine PDF trip submission", "Please find a PDF trip attached. " & TextBoxCaptain.Text & " (" & TextBoxVesselname.Text & ").", TripFileslist.Item(ListBoxTrips.SelectedIndex).FullPath)
                    LabelWait.Visible = False
                    Me.Refresh()
                    MessageBox.Show("PDF trip form " + ListBoxTrips.SelectedItem + " sent to " + TextBoxCompanyEmail.Text + ".")
                End If
            Catch ex As Exception
                LabelWait.Visible = False
                Me.Refresh()
                MessageBox.Show("There was an error sending the file :-\")
            End Try
        End If
    End Sub

 

    Private Sub ButtonManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonManual.Click
        Dim manual As String
        manual = Directory.GetCurrentDirectory + "\Manual\eTunaLog_operation_manual.pdf"
        MessageBox.Show("You are opening " + manual)
        If Not File.Exists(manual) Then
            MessageBox.Show("Sorry but the manual is not available in this version of the system...")
        Else
            Try
                System.Diagnostics.Process.Start(manual)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the manual :-\")
            End Try
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ViewData.Show()
    End Sub

    Public Sub InjectPDF()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        'cFDFDoc = cFDFApp.FDFOpenFromFile("C:\tmp\Logsheet_PS_2009.pdf", cFDFApp.FDFType.PDF)
        'cFDFDoc = cFDFApp.PDFOpenFromFile("C:\tmp\Logsheet_PS_2009.pdf")
        cFDFDoc = cFDFApp.PDFOpenFromFile("j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPAddForm("LOGSHEET", "j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPSetValue("VesselName1", "Emmanuel vessel", True, False)
        'cFDFDoc.XDPAddField("FullName", "Bo Duke", "C:\tmp\tmp.pdf")
        'cFDFDoc.PDFMergeFDF2Buf("c:\tmp\tmp.pdf", False, "")
        cFDFDoc.PDFMergeXDP2File("j:\pdf_forms\tmp\NewLogsheet1_new.pdf", "j:\pdf_forms\tmp\NewLogsheet1.pdf", False, "")
        ' to be completed
    End Sub

    Public Sub InjectPDF2()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        cFDFDoc = cFDFApp.PDFOpenFromFile("c:\tmp\Logsheet.pdf")
        cFDFDoc.XDPSetValue("VesselName1", "SOLOMON RUBY", True, False)
        cFDFDoc.XDPSetValue("Company1", "NFD", True, False)
        cFDFDoc.PDFMergeXDP2File("c:\tmp\NewLogsheet.pdf", "c:\tmp\Logsheet.pdf", False, "")
    End Sub

    Public Sub InjectPDF3()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        cFDFDoc = cFDFApp.FDFCreate
        cFDFDoc.FDFSetFile("j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPAddForm("LOGSHEET", "j:\pdf_forms\tmp\NewLogsheet1.pdf")
        'cFDFDoc.XDPAddSubForm("WholePage", "j:\pdf_forms\tmp\NewLogsheet1.pdf")
        'cFDFDoc.XDPAddSubForm("Trip_Header", "j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPSetValue("VesselName1", "ManuVessel")
        'cFDFDoc.XDPSetValue("Company", "ManuCompany")
        cFDFDoc.FDFSavetoFile("j:\pdf_forms\tmp\logsheet1-xdp.xdp", FDFApp.FDFDoc_Class.FDFType.XDP, True)
    End Sub

    Public Sub InjectPDF4()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        cFDFDoc = cFDFApp.FDFCreate
        cFDFDoc.FDFSetFile("http://www.nk-inc.com/logsheet-new.pdf")
        cFDFDoc.XDPAddForm("LOGSHEET", "http://www.nk-inc.com/logsheet-new.pdf")
        cFDFDoc.XDPAddSubForm("LOGSHEET", "http://www.nk-inc.com/logsheet-new.pdf")
        cFDFDoc.XDPSetValue("VesselName1", "Nick TEST2")
        cFDFDoc.XDPSetValue("Company1", "Nick TEST2")
        cFDFDoc.FDFSavetoFile("c:\tmp\logsheet-new.xdp", FDFApp.FDFDoc_Class.FDFType.XDP, True)
    End Sub

    Public Sub InjectPDF5()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        cFDFDoc = cFDFApp.FDFCreate
        cFDFDoc.FDFSetFile("j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPAddForm("LOGSHEET", "j:\pdf_forms\tmp\NewLogsheet1.pdf")
        cFDFDoc.XDPSetValue("VesselName1", "Emmanuel vessel", True, False)
        cFDFDoc.PDFMergeXDP2File("j:\pdf_forms\tmp\NewLogsheet1_new.pdf", "j:\pdf_forms\tmp\NewLogsheet1.pdf", False, "")
    End Sub

    Public Sub OUTPUTFDFDoc()
        Dim cFDFApp As New FDFApp.FDFApp_Class
        Dim cFDFDoc As New FDFApp.FDFDoc_Class
        cFDFDoc = cFDFApp.FDFCreate
        cFDFDoc.FDFSetFile("j:\pdf_forms\tmp\NewLogsheet1.pdf")

        For Each frm As FDFApp.FDFDoc_Class.FDFDoc_Class In cFDFDoc.XDPGetForms
            Console.WriteLine("Form Level:" & frm.FormLevel & "[" & frm.FormNumber & "]")
            For Each fld As FDFApp.FDFDoc_Class.FDFField In frm.struc_FDFFields
                Console.WriteLine(" FieldName:" & fld.FieldName)
            Next
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OUTPUTFDFDoc()
        'InjectPDF5()
    End Sub


    Private Sub ButtonNewFAD_Click(sender As Object, e As EventArgs) Handles ButtonNewFAD.Click
        Dim strNewFilename As String, strNewPath As String, strTemplate As String

        ' Check that we have the basic parameters - trip date, vesselname, ffavid
        If TextBoxVesselname.TextLength = 0 Or TextBoxFFAVID.TextLength = 0 Then
            MessageBox.Show("Please complete the parameters details before opening a new FAD form...")
            TabControl1.SelectTab(AboutTabNumber)
            TextBoxVesselname.Focus()
        ElseIf ComboBoxFADTemplate.SelectedIndex < 0 Then
            MessageBox.Show("Please select a FAD template before creating a new trip comments logsheet...")
            TabControl1.SelectTab(AboutTabNumber)
            ComboBoxFADTemplate.Focus()
        Else
            strNewFilename = TextBoxVesselname.Text.Trim + "_" + TextBoxFFAVID.Text.Trim + "_" + Format(DateTimePickerTripStartDate.Value, "yyyy-MM-dd") + "_FAD.pdf"

            strNewPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\FAD_Reports\" + strNewFilename

            strTemplate = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\Template_FAD\" + ComboBoxFADTemplate.SelectedItem

            If System.IO.File.Exists(strNewPath) Then
                MessageBox.Show("A file with this name already exists...")
            ElseIf Not System.IO.File.Exists(strTemplate) Then
                MessageBox.Show("The FAD template file could not be found in the TEMPLATE_FAD folder")
            Else
                If MessageBox.Show("Do you want to create " + strNewFilename + " ?", "NEW FAD REPORT", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    System.IO.File.Copy(strTemplate, strNewPath)
                    System.IO.File.SetAttributes(strNewPath, FileAttribute.Normal)

                    'MessageBox.Show("open file '" + strNewPath + "' here")
                    Try
                        System.Diagnostics.Process.Start(strNewPath)
                    Catch ex As Exception
                        MessageBox.Show("There was an error opening the file :-\")
                    End Try
                    ReadExistingFADFiles()
                End If
            End If
        End If

    End Sub

    Private Sub ButtonOpenFAD_Click(sender As Object, e As EventArgs) Handles ButtonOpenFAD.Click
        If ListBoxFAD.SelectedIndex < 0 Then
            MessageBox.Show("Please select a FAD report first...")
        Else
            'MessageBox.Show("You are opening " + CommentsFileslist.Item(ListBoxComments.SelectedIndex).FullPath)
            Try
                System.Diagnostics.Process.Start(FADFileslist.Item(ListBoxFAD.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the file :-\")
            End Try
        End If

    End Sub

    Private Sub ButtonDeleteFAD_Click(sender As Object, e As EventArgs) Handles ButtonDeleteFAD.Click
        If ListBoxFAD.SelectedIndex < 0 Then
            MessageBox.Show("Please select a FAD report to DELETE first...")
        Else
            'MessageBox.Show("You are deleting " + FADFileslist.Item(ListBoxFAD.SelectedIndex).FullPath)
            Try
                If MessageBox.Show("Do you want to DELETE FAD report '" + ListBoxFAD.SelectedItem + "' ?", "DELETE FAD REPORT", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    System.IO.File.Delete(FADFileslist.Item(ListBoxFAD.SelectedIndex).FullPath)
                    ReadExistingFADFiles()
                End If
            Catch ex As Exception
                MessageBox.Show("There was an error deleting the FAD report file :-\")
            End Try
        End If
    End Sub

    Private Sub ListBoxFAD_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxFAD.DoubleClick
        If ListBoxFAD.SelectedIndex >= 0 Then
            Try
                System.Diagnostics.Process.Start(FADFileslist.Item(ListBoxFAD.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the FAD report file :-\")
            End Try
        End If

    End Sub

    Private Sub ButtonVersion_Click(sender As Object, e As EventArgs) Handles ButtonVersion.Click
        Dim edoc As String
        edoc = Directory.GetCurrentDirectory + "\Manual\eTunaLog_Version_Summary.pdf"
        MessageBox.Show("You are opening " + edoc)
        If Not File.Exists(edoc) Then
            MessageBox.Show("Sorry but the document is not available in this version of the system...")
        Else
            Try
                System.Diagnostics.Process.Start(edoc)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the docuement :-\")
            End Try
        End If

    End Sub

    Private Sub LogEdit_Click(sender As Object, e As EventArgs)
        If ComboBoxTemplate.SelectedIndex < 0 Then
            MessageBox.Show("Please select a template first...")
        Else
            Try
                MessageBox.Show("YOU ARE EDITING THE MAIN PS LOGSHEET TEMPLATE" + Chr(13) + Chr(13) + "Please fill in the main header details, those you would like to be kept for each new trip:" + Chr(13) + Chr(13) + " - vessel," + Chr(13) + " - registration," + Chr(13) + " - FFAVid," + Chr(13) + " - WCPFCId," + Chr(13) + " - Company," + Chr(13) + " - IRC" + Chr(13) + Chr(13) + "When finished, save and close the template." + Chr(13) + Chr(13) + "(DO NOT ENTER ANY TRIP DATA INTO THIS TEMPLATE)", "Template EDIT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                System.Diagnostics.Process.Start(TemplateFileslist.Item(ComboBoxTemplate.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the template file :-\")
            End Try
        End If
    End Sub

    Private Sub FadEdit_Click(sender As Object, e As EventArgs)
        If ComboBoxFADTemplate.SelectedIndex < 0 Then
            MessageBox.Show("Please select a FAD template first...")
        Else
            Try
                MessageBox.Show("YOU ARE EDITING THE MAIN FAD TEMPLATE" + Chr(13) + Chr(13) + "Please fill in the VESSEL name and FFAVID (if any) in order to keep these parameters filled for each new FAD report" + Chr(13) + Chr(13) + "When finished, save and close the template." + Chr(13) + Chr(13) + "(DO NOT ENTER ANY FAD DATA INTO THIS TEMPLATE)", "FAD Template EDIT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                System.Diagnostics.Process.Start(FADTemplateFileslist.Item(ComboBoxFADTemplate.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the FAD template file :-\")
            End Try
        End If

    End Sub

   
    Private Sub ButtonHelp_Click(sender As Object, e As EventArgs) Handles ButtonHelp.Click
        TabControl1.SelectTab(AboutTabNumber)
        TextBoxVesselname.Focus()
        FirstSteps.Show()
    End Sub

  
    Private Sub LogLLEdit_Click(sender As Object, e As EventArgs)
        If ComboBoxLLTemplate.SelectedIndex < 0 Then
            MessageBox.Show("Please select a Longline template first...")
        Else
            Try
                MessageBox.Show("YOU ARE EDITING THE MAIN LONGLINE LOGSHEET TEMPLATE" + Chr(13) + Chr(13) + "Please fill in the main header details (first page), those you would like to be kept for each new trip:" + Chr(13) + Chr(13) + " - vessel," + Chr(13) + " - registration," + Chr(13) + " - FFAVid," + Chr(13) + " - WCPFCId," + Chr(13) + " - Company," + Chr(13) + " - IRC" + Chr(13) + Chr(13) + "When finished, save and close the template." + Chr(13) + Chr(13) + "(DO NOT ENTER ANY TRIP DATA INTO THIS TEMPLATE)", "Template EDIT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                System.Diagnostics.Process.Start(LLTemplateFileslist.Item(ComboBoxLLTemplate.SelectedIndex).FullPath)
            Catch ex As Exception
                MessageBox.Show("There was an error opening the template file :-\")
            End Try
        End If
    End Sub

    Private Sub CheckShowLogsheet_CheckedChanged(sender As Object, e As EventArgs) Handles CheckShowLogsheet.Click
        CheckLOGSHEETTAB(True)
    End Sub

    Private Sub CheckShowFAD_CheckedChanged(sender As Object, e As EventArgs) Handles CheckShowFAD.Click
        CheckFADTAB(True)
    End Sub

    Private Sub RadioFrench_click(sender As Object, e As EventArgs) Handles RadioFrench.Click
        CheckLanguage()
    End Sub

    Private Sub RadioEnglish_click(sender As Object, e As EventArgs) Handles RadioEnglish.Click
        If RadioEnglish.Checked = True Then
            MessageBox.Show("English will show up at next eTUNALOG startup...")
        End If
    End Sub

    Private Sub CheckFADTAB(FireMsg As Boolean)
        ' removing TABs according to option settings
        If CheckShowFAD.Checked = False Then
            TabControl1.Controls.Remove(TabControl1.TabPages(2))
        Else
            If FireMsg = True Then
                MessageBox.Show("Your FAD tab will appear on next eTUNALOG startup...")
                CheckShowFAD.Enabled = False
            End If
        End If
    End Sub

    Private Sub CheckLOGSHEETTAB(FireMsg As Boolean)
        ' removing TABs according to option settings
        If CheckShowLogsheet.Checked = False Then
            TabControl1.Controls.Remove(TabControl1.TabPages(1))
        Else
            If FireMsg = True Then
                MessageBox.Show("Your LOGSHEET tab will appear on next eTUNALOG startup...")
                CheckShowLogsheet.Enabled = False
            End If
        End If
    End Sub

    Private Sub CheckLanguage()
        If RadioFrench.Checked = True Then
            LabelTitle.Text = "Système de Gestion de Fiches de Pêche Electroniques"
            LabelTitle.Location.X.Equals(5)
            TabControl1.TabPages(1).Text = "Fiches de Pêche"
            TabControl1.TabPages(2).Text = "Fiches DCP"
            TabControl1.TabPages(3).Text = "Paramètres"
            TabControl1.TabPages(4).Text = "A Propos d'eTUNALOG"
            ButtonHelp.Text = "Nouveau sur eTUNALOG ? Suivez le guide.. (Anglais)"
            ButtonNewTripFile.Text = "Nouveau Voyage"
            ButtonOpenTrip.Text = "Ouvrir un voyage"
            ButtonDelete.Text = "Effacer"
            ButtonNewFAD.Text = "Nouvelle Fiche DCP"
            ButtonOpenFAD.Text = "Ouvrir Fiche DCP"
            ButtonDeleteFAD.Text = "Effacer"
            Parameters.TabPages(0).Text = "Détails Navire"
            Parameters.TabPages(1).Text = "Modèles PDF"
            Parameters.TabPages(2).Text = "Options"
            GroupBoxVessel.Text = "Paramètres Navire"
            LabelVesselName.Text = "Nom du navire (*)"
            GroupBoxGear.Text = "Type d'engin de pêche (*)"
            RadioSeiner.Text = "Senneur"
            RadioLongliner.Text = "Palangrier"
            LabelCompany.Text = "Société de pêche"
            LabelCountry.Text = "Pays"
            LabelCapt.Text = "Nom du Capitaine mon capitaine"
            LabelEmail.Text = "Email Société de Pêche"
            LabelCompuls.Text = "(*) Paramètre obligatoire (utilisé pour nommer la fiche de pêche)"
            LabelPS.Text = "Modèle Fiche SENNEUR"
            LabelLL.Text = "Modèle Fiche PALANGRIER"
            LabelFAD.Text = "Modèle Fiche DCP"
            GroupShow.Text = "Voir"
            CheckShowLogsheet.Text = "Fiches de Pêche"
            CheckShowFAD.Text = "Fiches DCP"
            GroupLanguage.Text = "Langage"
            GroupBoxTemplates.Text = "Modèles Fiches PDF"
            Me.Text = "Système de Gestion de Fiches de Pêche Electroniques"
        End If
    End Sub

End Class
