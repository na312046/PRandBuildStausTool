Imports System.Data.OleDb
Imports System.IO
Imports System.Text.RegularExpressions
Imports mshtml  'C:\Program Files (x86)\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll
Imports SHDocVw 'E:\Workflow\GetPRStatus\GetPRStatus\obj\Debug\Interop.SHDocVw.dll
Imports System.Configuration
Public Class frmHome

    Dim strFailedLst As String = ""
    Dim strXLSFile As String = ""
    Dim fileExt As String = ".XLS"
    Dim OcnStr As String = ""
    Dim ods As DataSet
    Dim odtbl As DataTable
    Dim CRMServerURL As String = ConfigurationManager.AppSettings("CRMServerURL").ToString()

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub
    Private Delegate Sub addlistbox(ByVal msg As String)
    Private Sub addlog(ByVal msg As String)
        If Me.lstBxLog.InvokeRequired Then
            Dim mydel As New addlistbox(AddressOf addlog)
            Me.Invoke(mydel, New Object() {msg})
        Else
            Me.lstBxLog.Items.Add(msg)
        End If

    End Sub
    Public Function getStatus(prLink As String) As String
        Dim strStatus As String = ""
        Try
            Dim value As Object = Nothing
            Dim doc As New HTMLDocument()
            Dim ie As New InternetExplorer()
            Dim isOpened As Boolean = False
            doc = IsIEWindowOpenedByUrl(prLink, isOpened)
            If isOpened = False Then
                'ie.Navigate("https://dynamicscrm.visualstudio.com/DefaultCollection/_git/CRM/pullrequest/144381?_a=overview", value, value, value, value)
                'ie.Navigate("https://dynamicscrm.visualstudio.com/DefaultCollection/_git/CRM/pullrequest/144381?_a=overview")
                ie.Navigate(prLink)
                ie.Visible = False
                System.Threading.Thread.Sleep(5000)
                Do
                Loop Until Not ie.Busy
                'Store HTML Document object
                doc = ie.Document
                System.Threading.Thread.Sleep(1000)
            End If
            Dim span As Object
            span = CountClassOccur("span", "status-indicator abandoned", doc)
            If (span(0) = 1) Then
                strStatus = "abandoned"
            End If

            'Get all the anchor tags
            If (strStatus IsNot "abandoned") Then

                Dim action As Object
                Dim status As Object

                action = CountClassOccur("a", "actionLink", doc)
                If (action(0) = 1) Then

                    If (String.Equals(action(1), "Build succeeded")) Then
                        strStatus = action(1)
                    Else
                        strStatus = Rebuild(ie)
                    End If
                ElseIf (action(0) = 0) Then
                    status = CountClassOccur("span", "statusText", doc)
                    strStatus = status(1)
                End If
            End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            strStatus = ex.Message
        Finally
            CloselEWindowsByURL(prLink)
        End Try
        Return strStatus
    End Function
    Private Sub BuildStatusInExcelOld(strStatus As String, strPr As String)
        Dim ocn As New OleDbConnection(OcnStr)
        Try
            If (ocn.State = ConnectionState.Closed) Then
                ocn.Open()
            End If
            Dim Query As String = "Update [Sheet1$] set [Status]='" & strStatus & "' where [PR Link] ='" & strPr & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(Query, ocn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            WriteLog("Error in opening connection to update PR Status in Excel: " & vbCrLf & ex.Message)
        Finally
            If ocn IsNot Nothing Then
                If ocn.State = ConnectionState.Open Then
                    ocn.Close()
                End If
                ocn = Nothing
            End If
        End Try
    End Sub

    Private Sub BuildStatusInExcel(myds As DataSet)
        Dim cn As OleDbConnection = New OleDbConnection(OcnStr)
        Dim cmd As OleDbCommand
        Try
            Dim query As String = "Update [Sheet1$] set [Status]=@statusval where [PR Link]=@prlink"
            cmd = New OleDbCommand(query, cn)
            cmd.Parameters.Add("@statusval", OleDbType.LongVarChar, 201, "Status")
            cmd.Parameters.Add("@prlink", OleDbType.LongVarChar, 201, "PR link")
            Using odas As New OleDbDataAdapter("Select [PR Link],[Status] from [Sheet1$]", cn)
                odas.UpdateCommand = cmd
                odas.Update(myds, "tbldata")
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Try
            txtFileName.Text = String.Empty
            Dim dr As DialogResult = OpenFileDlg.ShowDialog()
            If dr <> DialogResult.OK Then
                Return
            End If
            txtFileName.Text = OpenFileDlg.FileName
        Catch ex As Exception
            WriteLog(ex.Message)
        End Try
    End Sub
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        Dim ocn As OleDbConnection = Nothing
        Try
            Try
                strXLSFile = txtFileName.Text.Trim()
                If String.IsNullOrEmpty(strXLSFile) Then
                    MsgBox("Plese select PRs list file.", MsgBoxStyle.Exclamation, strCaption)
                    Exit Sub
                End If
                If Not File.Exists(strXLSFile) Then
                    MsgBox("Excel file not exists in physica location.", MsgBoxStyle.Exclamation, strCaption)
                    Exit Sub
                End If
                fileExt = Path.GetExtension(strXLSFile)

                If fileExt.ToUpper() = ".XLS" Then
                    OcnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strXLSFile + ";Extended Properties=Excel 8.0;" 'HDR=Yes;IMEX=2"
                ElseIf fileExt.ToUpper() = ".XLSX" Then
                    'OcnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strXLSFile & ";Extended Properties=Excel 12.0 Xml;HDR=Yes;IMEX=2"
                    OcnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strXLSFile & ";Extended Properties=Excel 12.0 ;"
                End If
                ocn = New OleDbConnection(OcnStr)
                Try
                    If (ocn.State = ConnectionState.Closed) Then
                        ocn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Error in opening connection to read data from Excel: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, strCaption)
                    Exit Sub
                End Try
                Dim Query As String = "Select [PR Link],[Status] from [Sheet1$]"
                'WriteLog(Query)
                Using oda As New OleDbDataAdapter(Query, ocn)
                    Try
                        ods = New DataSet()
                        oda.Fill(ods, "tblData")
                        'odtbl = New DataTable("tblData")
                        'oda.Fill(odtbl)
                    Catch ex As Exception
                        MessageBox.Show("Error in reading data from Excel: " & vbCrLf & ex.Message, strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End Try
                End Using
                'writeLog("Columns count in Excel file: " & CInt(ods.Tables("tblData").Columns.Count))
                Dim recCount As Integer = 0
                recCount = ods.Tables("tblData").Rows.Count
                If recCount = 0 Then
                    MessageBox.Show("No records in Excel file", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Else
                    'Login to the CRM server
                    If Not OpenDynamicsCRMUrl(CRMServerURL) Then
                        MessageBox.Show("Login failed, unable to open https://dynamicscrm.visualstudio.com in IE." & vbCrLf & "Please login with valid credentials and try again.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    lstBxLog.Items.Clear()
                    btnBrowse.Enabled = False
                    btnStart.Enabled = False
                    btnClose.Enabled = False
                    btnClear.Enabled = False
                    ProgressBar1.Visible = True
                    ProgressBar1.Maximum = recCount
                    ProgressBar1.Value = 0
                    lblProgressPercent.Visible = True
                    BackgroundWorker1.RunWorkerAsync()
                End If
            Catch ex As Exception
                WriteLog(ex.Message)
            End Try
        Catch ex As Exception
            WriteLog("btnStart_Click-->" + ex.Message)
        Finally
            If ocn IsNot Nothing Then
                If ocn.State = ConnectionState.Open Then
                    ocn.Close()
                End If
                ocn = Nothing
            End If
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            If Not ods Is Nothing AndAlso Not ods.Tables("tblData") Is Nothing AndAlso ods.Tables("tblData").Rows.Count > 0 Then
                odtbl = ods.Tables(0)
                Dim i As Integer = 0
                Dim recCount As Integer = 0
                recCount = odtbl.Rows.Count
                addlog("######## Total PRs count: " & recCount)
                Dim strPR As String = ""
                For Each dr As DataRow In odtbl.Rows
                    i += 1
                    Dim strUrl As String = dr(0).ToString
                    strPR = Regex.Split(strUrl.ToLower, "pullrequest/")(1)
                    strPR = strPR.Substring(0, strPR.IndexOf("?"))
                    Dim strStatus As String = getStatus(strUrl)
                    dr(1) = strStatus
                    'BuildStatusInExcelOld(strStatus, strUrl)
                    addlog(i & ") PR#" & strPR & "  Status: " & strStatus)

                    BackgroundWorker1.ReportProgress(i)
                    System.Threading.Thread.Sleep(100)
                Next
                BuildStatusInExcel(ods)
            Else
                MsgBox("No data found")
                Exit Sub
            End If
        Catch ex As Exception
            WriteLog("BackgroundWorker1_DoWork-->" + ex.Message)
        Finally
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
        Dim percent As Integer = Math.Round((ProgressBar1.Value / ProgressBar1.Maximum) * 100)
        lblProgressPercent.Text = percent.ToString() & "%"
        'lblProgressPercent.BackColor = Color.Transparent
        lblProgressPercent.BackColor = SystemColors.ControlLight
        If percent > 40 Then
            lblProgressPercent.BackColor = Color.LightGreen
        End If
        lblProgressPercent.Refresh()
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            btnBrowse.Enabled = True
            btnStart.Enabled = True
            btnClose.Enabled = True
            btnClear.Enabled = True
            If String.IsNullOrEmpty(strFailedLst) Then
                addlog("########## PR status verification completed successfully. ########")
                MessageBox.Show("PR status verification completed successfully. Please check logs for more details.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                WriteLog("####Status verification failed PR list####" + vbCrLf + strFailedLst)
                addlog("##### Status verification failed for some PRs. Please check logs for more details.")
                MessageBox.Show("Status verification failed for some PRs. Please check logs for more details.", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            WriteLog("BackgroundWorker1_RunWorkerCompleted-->" + ex.Message)
        End Try
    End Sub
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtFileName.Text = ""
        lstBxLog.Items.Clear()
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        lblProgressPercent.Visible = False
        lblProgressPercent.Text = ""
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub tsmItemSave_Click(sender As Object, e As EventArgs) Handles tsmItemSave.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Text Files (*.txt*)|*.txt"
        saveFileDialog1.Title = "Save PR status log"
        saveFileDialog1.FileName = "PRStatusLog_" & DateTime.Now.ToString("ddMMyyyy_hhmmss")
        If saveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK AndAlso Not String.IsNullOrEmpty(saveFileDialog1.FileName) Then
            Using streamwriter As New IO.StreamWriter(saveFileDialog1.FileName)
                For Each str As String In lstBxLog.Items
                    streamwriter.WriteLine(str)
                Next
            End Using
        End If
    End Sub

    Private Sub lstBxLog_MouseDown(sender As Object, e As MouseEventArgs) Handles lstBxLog.MouseDown
        If e.Button = MouseButtons.Right AndAlso lstBxLog.Items.Count > 0 AndAlso btnStart.Enabled Then
            lstBxLog.ContextMenuStrip = ContextMenuStrip1
        Else
            lstBxLog.ContextMenuStrip = Nothing
        End If
    End Sub
End Class