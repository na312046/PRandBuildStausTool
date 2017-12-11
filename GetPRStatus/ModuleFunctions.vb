Imports System.IO
Imports System.Configuration
Imports mshtml 'C:\Program Files (x86)\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll
Imports SHDocVw '\\v-gadidd-dev\workflow\GetPRStatus\GetPRStatus\obj\Debug\Interop.SHDocVw.dll
Module ModuleFunctions
    Public strCaption As String = "CRM :: PRStatus"
    Public Function IsIEWindowOpened(ByVal strWindowTitle As String, Optional ByVal closeWindow As Boolean = False) As Boolean
        Dim isWindowOpened As Boolean = False
        Try
            Dim shellWindows As ShellWindows = New ShellWindowsClass()
            For Each ie As IWebBrowser2 In shellWindows
                If strWindowTitle IsNot Nothing AndAlso strWindowTitle <> String.Empty AndAlso ie.Name.Contains("Internet Explorer") Then
                    Dim htmlDoc As IHTMLDocument2 = TryCast(ie.Document, IHTMLDocument2)
                    If htmlDoc IsNot Nothing AndAlso htmlDoc.title IsNot Nothing AndAlso htmlDoc.title.StartsWith(strWindowTitle) Then
                        If htmlDoc IsNot Nothing AndAlso htmlDoc.title IsNot Nothing AndAlso htmlDoc.title.Contains(strWindowTitle) Then
                            isWindowOpened = True
                            If closeWindow Then
                                ie.Quit()
                            End If
                            Exit For
                        End If
                        End
                    End If
                End If
            Next
        Catch
        End Try
        Return isWindowOpened
    End Function

    Public Function IsIEWindowOpenedByUrl(ByVal strURL As String, ByRef isOpened As Boolean) As HTMLDocument
        Dim isWindowOpened As New HTMLDocument()
        isOpened = False
        Try
            Dim shellWindows As ShellWindows = New ShellWindowsClass()
            For Each ie As IWebBrowser2 In shellWindows
                If strURL IsNot Nothing AndAlso strURL <> String.Empty AndAlso ie.Name.Contains("Internet Explorer") Then
                    Dim htmlDoc As HTMLDocument = TryCast(ie.Document, HTMLDocument)
                    If htmlDoc IsNot Nothing AndAlso htmlDoc.url IsNot Nothing AndAlso htmlDoc.url.StartsWith(strURL) Then
                        'If htmlDoc IsNot Nothing AndAlso htmlDoc.title IsNot Nothing AndAlso htmlDoc.title.Contains(strURL) Then
                        isWindowOpened = htmlDoc
                        isOpened = True
                        Exit For
                    End If
                End If
            Next
        Catch
        End Try
        Return isWindowOpened
    End Function
    Public Function CountClassOccur(tag As String, className As String, doc As IHTMLDocument) As Object
        Dim iHTMLCol As IHTMLElementCollection
        Dim iHTMLEle As IHTMLElement


        iHTMLCol = doc.getElementsByTagName(tag)
        Dim count As Int32 = 0
        Dim reviewer As String = ""
        For Each iHTMLEle In iHTMLCol
            If (iHTMLEle.className IsNot Nothing) Then
                If (String.Equals(iHTMLEle.className, className)) Then
                    count += 1
                    reviewer = iHTMLEle.innerText
                End If
            End If
        Next
        Return {count, reviewer, True}
    End Function

    Public Function Rebuild(ie As InternetExplorer) As String
        Dim doc As IHTMLDocument = ie.Document
        Dim strStatus As String = ""
        Dim iHTMLCol As IHTMLElementCollection = doc.getElementsByTagName("i")
        Dim count As Int32 = 0
        Try
            For Each iHTMLEle As IHTMLElement In iHTMLCol
                If (iHTMLEle.getAttribute("data-icon-name") IsNot Nothing AndAlso iHTMLEle.getAttribute("data-icon-name").ToString = "More") Then
                    If (count > 0) Then
                        iHTMLEle.click()
                        Dim HTMLDoc As New HTMLDocument
                        HTMLDoc = ie.Document
                        iHTMLCol = HTMLDoc.getElementsByTagName("button")
                        For Each iHTMLE As IHTMLElement In iHTMLCol
                            If (iHTMLE.getAttribute("name") IsNot Nothing AndAlso iHTMLE.getAttribute("name").ToString = "Queue build") Then
                                iHTMLEle.click()
                                strStatus = "Build in progress (RE-QUEUED)"
                                Exit For
                            End If
                        Next
                        Exit For
                    End If
                    count += 1
                End If
            Next
        Catch ex As Exception
            strStatus = ex.Message
        End Try
        Return strStatus
    End Function
    Public Function OpenDynamicsCRMUrl(ByVal strURL As String) As Boolean
        Dim IsLoggedIn As Boolean = False
        Try
            Dim value As Object = Nothing
            Dim doc As New HTMLDocument()
            Dim ie As New InternetExplorer()
            Dim isOpened As Boolean = False
            doc = IsIEWindowOpenedByUrl(strURL, isOpened)
            If isOpened Then
                IsLoggedIn = True
            Else
                'ie.Navigate(strURL, value, value, value, value)
                ie.Navigate(strURL)
                ie.Visible = True
                System.Threading.Thread.Sleep(5000)

                Do
                Loop Until Not ie.Busy

                'Store HTML Document object
                doc = ie.Document
                System.Threading.Thread.Sleep(1200)

                Dim htmlDoc As IHTMLDocument2 = TryCast(ie.Document, IHTMLDocument2)
                If htmlDoc IsNot Nothing AndAlso htmlDoc.title IsNot Nothing Then
                    Dim pageTitle As String = htmlDoc.title.ToLower()
                    If pageTitle.Contains("sign in to your account") OrElse pageTitle.Contains("authentication options") OrElse pageTitle.Contains("multi-factor authentication ") OrElse pageTitle.Contains("waiting for response") OrElse pageTitle.Contains("sign out") OrElse pageTitle.Contains("error") Then
                        MessageBox.Show("Please login through microsoft credentials and phone authentication in the opened IE browser." & vbCrLf & "After successfull login please click on Ok to proceed", strCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        System.Threading.Thread.Sleep(100)
                    End If
                    'For Each browser As InternetExplorer In New ShellWindows()
                    '    If browser.LocationURL.ToString().Contains(strURL) Then
                    '        strIeStatus = "Pass"
                    '        Exit For
                    '    End If
                    'Next
                    Dim shellWindows As ShellWindows = New ShellWindowsClass()
                    For Each IeBrowser As IWebBrowser2 In shellWindows
                        If strURL IsNot Nothing AndAlso strURL.Trim() <> String.Empty AndAlso ie.Name.Contains("Internet Explorer") Then
                            Dim IehtmlDoc As IHTMLDocument2 = TryCast(IeBrowser.Document, IHTMLDocument2)
                            If IehtmlDoc IsNot Nothing AndAlso IehtmlDoc.url IsNot Nothing AndAlso IehtmlDoc.url.Contains(strURL) Then
                                IsLoggedIn = True
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, strCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
            WriteLog("OpenDynamicsCRMUrl()-->" & ex.Message)
            IsLoggedIn = False
        Finally
        End Try
        Return IsLoggedIn
    End Function
    Public Sub CloselEWindowsByTitle(ByVal strWindowTitle As String)
        Try
            Dim shellWindows As ShellWindows = New ShellWindowsClass()
            For Each ie As IWebBrowser2 In shellWindows
                If strWindowTitle IsNot Nothing AndAlso strWindowTitle.Trim() <> String.Empty AndAlso ie.Name.Contains("Internet Explorer") Then
                    Dim htmlDoc As IHTMLDocument2 = TryCast(ie.Document, IHTMLDocument2)
                    If htmlDoc IsNot Nothing AndAlso htmlDoc.title IsNot Nothing AndAlso htmlDoc.title.Contains(strWindowTitle) Then
                        ie.Quit()
                    End If
                End If
            Next
        Catch

        End Try
    End Sub
    Public Sub CloselEWindowsByURL(ByVal strURL As String)
        Try
            Dim shellWindows As ShellWindows = New ShellWindowsClass()
            For Each ie As IWebBrowser2 In shellWindows
                If strURL IsNot Nothing AndAlso strURL.Trim() <> String.Empty AndAlso ie.Name.Contains("Internet Explorer") Then
                    Dim htmlDoc As IHTMLDocument2 = TryCast(ie.Document, IHTMLDocument2)
                    If htmlDoc IsNot Nothing AndAlso htmlDoc.url IsNot Nothing AndAlso htmlDoc.url.Contains(strURL) Then
                        ie.Quit()
                    End If
                End If
            Next
        Catch

        End Try
    End Sub

    Public Sub OpenCurrentPage(ByVal accountNumber As String, ByVal pageTitle As String)
        Try
            Dim shellWindows As SHDocVw.ShellWindows = New SHDocVw.ShellWindowsClass()
            For Each ie As SHDocVw.InternetExplorer In shellWindows
                If ie.Name.Contains("Internet Explorer") Then
                    Dim htmlDoc As IHTMLDocument2 = ie.Document
                    If (htmlDoc IsNot Nothing And htmlDoc.title IsNot Nothing And htmlDoc.title.ToLower().Contains(pageTitle.ToLower())) Then
                        System.Threading.Thread.Sleep(300)
                        Dim labels As IHTMLElementCollection = htmlDoc.getElementsByTagName("Label")
                        Dim btns As IHTMLElementCollection = htmlDoc.getElementsByTagName("A")
                        Dim header As IHTMLElement = htmlDoc.getElementsByTagName("h1")(0)
                        If (header Is Nothing) Then
                            If (labels.length > 1) Then
                                For Each Lbl As IHTMLElement In labels
                                    If (Lbl.innerText.Trim.ToUpper.Equals("PRODUCT")) Then
                                        Dim txtProduct As IHTMLInputElement = Lbl.parentNode.nextSibling.children(0)
                                        txtProduct.value = accountNumber
                                        Exit For
                                        'htmlDoc.page.ExecuteJScript(BMMHelper.GetBMMMenuClicOS(BMMChieldPageMenu.Continu))  
                                    End If
                                Next
                            Else
                                Dim grid As IHTMLElement2 = htmlDoc.getElementsByTagName("Select")(0)
                                Dim options As IHTMLElementCollection = grid.getElementsByTagName("Option")
                                For Each item As IHTMLElement In options
                                    If (item.innerText.Contains(accountNumber)) Then
                                        item.selected = True
                                        Exit For
                                    End If
                                Next
                            End If
                            For Each Lbl As IHTMLElement In labels
                                If (Lbl.innerText.Trim.ToUpper.Equals("PRODUCT")) Then
                                    Dim txtProduct As IHTMLInputElement = Lbl.parentNode.nextSibling.children(0)
                                    txtProduct.value = accountNumber
                                    Exit For
                                    'htm1Doc.page.ExecuteJScript(BMMHelper.GetBMMMenuClick)S(BMMChieldPageMenu.Continu))
                                End If
                            Next
                            For Each btn As IHTMLElement In btns
                                If (btn.innerText.Trim.ToUpper.Equals("CONTINUE")) Then
                                    btn.click()
                                    Exit For
                                End If
                            Next

                        Else
                            If (header.innerText.Trim.ToUpper.Contains("LOCATE")) Then
                                If (labels.length > 1) Then
                                    For Each Lbl As IHTMLElement In labels
                                        If (Lbl.innerText.Trim.ToUpper.Equals("PRODUCT")) Then
                                            Dim txtProduct As IHTMLInputElement = Lbl.parentNode.nextSibling.children(0)
                                            txtProduct.value = accountNumber
                                            Exit For
                                            'htmlDoc.page.ExecuteJScript(BMMHelper.GetBMMMenuClicOS(BMMChieldPageMenu.Continu))  
                                        End If
                                    Next
                                Else
                                    Dim grid As IHTMLElement2 = htmlDoc.getElementsByTagName("Select")(0)
                                    Dim options As IHTMLElementCollection = grid.getElementsByTagName("Option")
                                    For Each item As IHTMLElement In options
                                        If (item.innerText.Contains(accountNumber)) Then
                                            item.selected = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                For Each btn As IHTMLElement In btns
                                    If (btn.innerText.Trim.ToUpper.Equals("CONTINUE")) Then
                                        btn.click()
                                        Exit For
                                    End If
                                Next
                            Else
                                Exit For
                            End If
                            Exit For
                        End If
                        System.Threading.Thread.Sleep(100)
                    End If
                End If
            Next
        Catch ex As Exception
            'write log
        End Try

    End Sub
    Public Function WordExists(ByVal searchString As String, ByVal findString As String) As Boolean
        Dim returnValue As Boolean = False
        If System.Text.RegularExpressions.Regex.Matches(searchString, "\b" & findString & "\b").Count > 0 Then
            returnValue = True
        End If
        Return returnValue
    End Function
    Public Sub WriteLog(strError As String)
        Try
            Dim AppPath As String = AppDomain.CurrentDomain.BaseDirectory
            Dim strLog As String = "LOG\"
            Dim strFilePath As String = AppPath & strLog

            If Not (Directory.Exists(strFilePath)) Then
                Directory.CreateDirectory(strFilePath)
            End If
            Dim fn As String = String.Format("{0}{1}.txt", strFilePath, DateTime.Now.ToString("ddMMyyyy"))
            Dim fs As New FileStream(fn, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)

            Dim writer As New StreamWriter(fs)
            writer.WriteLine(String.Format("[ {0} ] {1}", DateTime.Now.ToString("HH:mm:ss"), strError))
            writer.Close()
            fs.Close()
        Catch ex As Exception

        Finally
        End Try
    End Sub

    Public Function closeOpenedWord(fPath As String) As Boolean
        Dim retVal As Boolean = False
        Dim fs As FileStream = Nothing
        Try
            If File.Exists(fPath) Then
                fs = New FileStream(fPath, FileMode.Open, FileAccess.Read)
            End If
        Catch ex As Exception
            retVal = True
            Dim fileName As String = Path.GetFileNameWithoutExtension(fPath).ToUpper().Trim()
            Dim oProcess As Process() = Process.GetProcessesByName("WINWORD")
            For Each item As Process In oProcess
                If item.MainWindowTitle.ToUpper().Contains(fileName) Then
                    item.Kill()
                End If
                'WriteLog("closeOpenWord()-->" + ex.Message);
            Next
        Finally
            If fs IsNot Nothing Then
                fs.Close()
            End If
        End Try
        Return retVal
    End Function
    Public Function closeOpenedExcel(fPath As String) As Boolean
        Dim retVal As Boolean = False
        Dim fs As FileStream = Nothing
        Try
            fs = New FileStream(fPath, FileMode.Open, FileAccess.Read)
        Catch ex As Exception
            retVal = True
            Dim fileName As String = Path.GetFileNameWithoutExtension(fPath).ToUpper().Trim()
            Dim oProcess As Process() = Process.GetProcessesByName("EXCEL")
            For Each item As Process In oProcess
                If item.MainWindowTitle.ToUpper().Contains(fileName) Then
                    item.Kill()
                End If
                'WriteLog("closeOpenedExcel()-->" + ex.Message);
            Next
        Finally
            If fs IsNot Nothing Then
                fs.Close()
            End If
        End Try
        Return retVal
    End Function


End Module
