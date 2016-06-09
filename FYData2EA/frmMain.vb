Imports System.IO
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Xml

Public Class frmMain

    Private apppath As String = ""
    Private sVer As String = ""
    Private nErr As Integer = 0
    Dim rs As New Resizer

    Structure Matrix
        Public year As String
        Public systemName As String
        Public strategicAlignment As String
        Public regulations As String
        Public span As String
        Public efficiency As String
        Public benefit As String
        Public sysUtil As String
        Public sharedServices As String
        Public futureState As String
        Public operPerf As String
        Public costVariance As String
        Public custService As String
        Public changeReq As String
        Public operAnalysis As String
        Public risks As String
        Public userSatisfaction As String
        Public sysEnhancement As String
        Public techRelevance As String
    End Structure
    Enum id
        strategicAlignment
        regulations
        span
        efficiency
        benefit
        sysUtil
        sharedServices
        futureState
        operPerf
        costVariance
        custService
        changeReq
        operAnalysis
        risks
        userSatisfaction
        sysEnhancement
        techRelevance
    End Enum
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        rs.FindAllControls(Me)
        '*** open a File Dialog Box to allow the user to select a folder
        Try
            Dim m_xmld = New XmlDocument()
            m_xmld.Load(Application.ExecutablePath & ".manifest")
            sVer = "v" & m_xmld.ChildNodes.Item(1).ChildNodes.Item(0).Attributes.GetNamedItem("version").Value
        Catch ex As Exception
        Finally
        End Try
        If apppath = "" Then
            Dim dialog As New FolderBrowserDialog()
            dialog.RootFolder = Environment.SpecialFolder.Desktop
            dialog.SelectedPath = "C:\"
            dialog.ShowNewFolderButton = False
            dialog.Description = "Select the directory for FY data (.xls files) - " & sVer
            If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                apppath = dialog.SelectedPath
            Else
                Application.Exit()
            End If

        End If
        Me.Text = Me.Text & " " & sVer
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object,
            ByVal e As System.EventArgs) Handles Me.Resize

        rs.ResizeAllControls(Me)

    End Sub
    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

        Try

            '*** Disable the button to prevent double-click. ***
            Me.btnImport.Enabled = False

            '*** Output message to UI. ***
            Me.txtOutput.Text = "Processing... please wait" & Environment.NewLine

            '*** Retrieve a list of OSIM Excel files from a given directory. ***
            Dim files As List(Of String) = GetOSIMFiles(apppath)

            '*** Proceed only if there is one or more files. ***
            If Not files Is Nothing AndAlso files.Count > 0 Then
                '*** Configure the progress bar. ***
                With Me.pbarProgress
                    .Style = Windows.Forms.ProgressBarStyle.Continuous
                    .Maximum = 100
                    .Minimum = 0
                    .Value = 0
                    .Step = CInt(100 / files.Count)
                    .Visible = True
                End With

                '*** Do processing on a separate thread. Pass in a list of file names. ***
                '*** This line will trigger the BackgroundWorker1_DoWork() event. ***
                Me.BackgroundWorker1.RunWorkerAsync(files)

            Else

                System.Windows.Forms.MessageBox.Show("There is no file (.xls) to process in the specified directory.",
                "Action Terminated", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)

                Me.txtOutput.Text = "Action terminated."

            End If

        Catch ex As Exception

            Me.txtOutput.Text = "Action terminated."
            System.Windows.Forms.MessageBox.Show(ex.Message, "Application Error",
                                                 System.Windows.Forms.MessageBoxButtons.OK,
                                                 System.Windows.Forms.MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim oApp As Excel.Application = Nothing
        Dim oWorkbooks As Excel.Workbooks = Nothing
        Dim oWorkbook As Excel.Workbook = Nothing
        Dim oSheets As Excel.Sheets = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim elementNo As Integer = 0
        Dim counter As Integer = 0
        Dim bridgeAdded As Integer = 0
        Dim bridgeElementAdded As Integer = 0
        Dim k As Integer = 0


        Try

            '*** Instantiate a new Excel application object. ***
            oApp = New Excel.Application
            oApp.Visible = False
            oApp.ScreenUpdating = False
            oApp.DisplayAlerts = False

            '*** Get a list of files from the Argument object. ***
            Dim files As List(Of String) = CType(e.Argument, List(Of String))

            '*** Iterate through all the Excel files. ***
            For Each f As String In files

                '*** Report progress back to UI thread. ***
                Me.BackgroundWorker1.ReportProgress(counter, String.Format("Processing {0}", f))

                k += 1
                Using w As StreamWriter = File.AppendText(apppath & "\log.txt")
                    Log("Reading " & f, w)
                End Using
                Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                    LogText(String.Format("[{0}] - {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString()), w)
                End Using

                '*** Convert the counter into a percent. ***
                counter = CInt((k * 100) / files.Count)

                '*** Open each Excel file. ***
                oWorkbooks = oApp.Workbooks
                oWorkbook = oWorkbooks.Open(f)
                System.Threading.Thread.Sleep(3000)
                oWorkbook.KeepChangeHistory = False
                oSheets = oWorkbook.Worksheets

                Dim sheetName As String = ""
                Dim tmp As Integer = -1
                Dim observations(19) As String
                Dim observations2(19) As String
                Dim observations3(19) As String
                '*** Iterate through all the Excel worksheets. ***
                For i As Integer = 1 To oSheets.Count

                    oSheet = CType(oSheets(i), Excel.Worksheet)
                    sheetName = oSheet.Name.Trim

                    Array.Clear(observations, 0, observations.Length)
                    Array.Clear(observations2, 0, observations.Length)
                    Array.Clear(observations3, 0, observations.Length)

                    If sheetName = "EA Business Sub-Functions" Then
                        ReadSubFunctionsSheet(oSheet, observations)
                        If Not observations(0) = String.Empty Then
                            Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                                LogText(String.Format("[Business SubFunctions] - Budget distribution towards sub-functions: {0}", observations(0)), w)
                            End Using
                        End If
                    End If
                    If sheetName = "EA Business Functions" Then
                        ReadFunctionsSheet(oSheet, observations)
                        If Not observations(0) = String.Empty Then
                            Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                                LogText(String.Format("[Business Functions] - Alignment towards Business Functions: {0}", observations(0)), w)
                            End Using
                        End If
                    End If

                    If sheetName = "EA-Surveillance Systems" Then
                        ReadSurveillanceSheet(oSheet, observations2)
                        If Not observations2(0) = String.Empty Or Not observations2(1) = String.Empty Then
                            Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                                If Not observations2(0) = String.Empty Then
                                    LogText(String.Format("Surveillance Portfolio: {0}", observations2(0)), w)
                                End If
                            End Using
                        End If

                    End If

                    If sheetName = "IT System Assessment Details" Then

                        '*** Proceed here only if the worksheet is the AssessmentData. ***
                        ReadData(oSheet, observations, observations2, observations3)

                        For d As Integer = 0 To 16
                            Dim cnt As Integer = 0
                            Dim cnt2 As Integer = 0
                            Dim cnt3 As Integer = 0
                            Dim msg As String = ""
                            Dim msg2 As String = ""
                            Dim msg3 As String = ""

                            If Not IsNothing(observations(d)) Then
                                cnt = GetArrayLength(observations(d), "|")
                            End If
                            If Not IsNothing(observations2(d)) Then
                                cnt2 = GetArrayLength(observations2(d), "|")
                            End If
                            If Not IsNothing(observations3(d)) Then
                                cnt3 = GetArrayLength(observations3(d), "|")
                            End If
                            If cnt > 0 Then
                                Dim str As String = ""
                                Select Case d
                                    Case id.strategicAlignment
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio have been reported as providing minimal or no contribution to the CDC and/or CIO's strategic objectives.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} system(s) in the portfolio have been reported as providing minimal or no contribution to the CDC and/or CIO's strategic objectives: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was reported as providing minimal or no contribution to the CDC and/or CIO's strategic objectives.", observations(d))
                                        End If
                                        msg = msg & " Recommendation: Assess the actual business need during the next annual operational analysis and/or begin planning the retirement of these systems."

                                    Case id.efficiency
                                        Exit Select

                                    Case id.changeReq
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio have been reported with operational and/or technical deficiencies which prevent meeting the business needs.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} system(s) in the portfolio have been reported with operational and/or technical deficiencies which prevent meeting the business needs: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was reported with operational and/or technical deficiencies which prevent meeting the business needs.", observations(d))
                                        End If
                                        msg = msg & " Recommendation:  An analysis of alternatives (AoA) would help determine a suitable solution replacement and/or technology refresh."

                                    Case id.benefit
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio are not meeting the business benefits and/or business needs.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} system(s) in the portfolio are not meeting the business benefits and/or business needs: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} is not meeting the business benefits and/or business needs.", observations(d))
                                        End If
                                        msg = msg & " Recommendation:  Begin planning retirement for these systems or alternatively consider performing a gap analysis during the next OA to identify suitable replacement."
                                    Case id.sysEnhancement
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio were reported as rigid and un-feasible to undergo enhancements when new business requirements emerge.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} system(s) in the portfolio were reported as rigid and un-feasible to enhance/modify when new business requirements emerge: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was reported as rigid and un-feasible to enhance/modify when new business requirements emerge.", observations(d))
                                        End If
                                        msg = msg & " Recommendation:  Systems that cannot adapt to emerging business requirements represent a risk to the organization and should be considered for a technology refresh or solution replacement."


                                    Case id.futureState
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio have been flagged as candidate for retirement during the latest FY IT assessment.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} system(s) in the portfolio were flagged as candidate for retirement during the latest FY IT assessment: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was flagged for retirement during the latest FY IT assessment.", observations(d))
                                        End If
                                        msg = msg & " Recommendation:  Revisit the latest operational analysis for these systems and plan accordingly."

                                    Case id.techRelevance
                                        If cnt > 9 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio are reported at risk for operational interruption given their dependencies on outdated or obsolete technologies.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg = String.Format("The following {0} systems in the portfolio are reported at risk for operational interruption given their dependencies on outdated and obsolete technologies: {1}.", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was reprted at risk for operational interruption given its reliance on outdated and obsolete technologies.", observations(d))
                                        End If
                                        msg = msg & " Recommendation:  Begin planning a technology refresh to ensure proper continuity of services with the exception of those systems targeted for retirement."

                                End Select
                            End If
                            If cnt2 > 0 Then
                                Dim str As String = ""
                                Select Case d
                                    Case id.sharedServices
                                        If cnt2 > 9 Then
                                            msg2 = String.Format("Exactly {0} systems in the portfolio have been reported as having capabilities that would be valuable to other organizations but are not currently shared.", cnt2)
                                        ElseIf cnt2 > 1 Then
                                            str = ReplaceLastOccurrence(observations2(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg2 = String.Format("The following {0} systems in the portfolio have been reported as having capabilities that would be valuable to other organizations but are not currently shared: {1}.", cnt2, str)
                                        Else
                                            msg2 = String.Format("The [system] {0} is reported as having capabilities that would be valuable to other organizations but these capabilities are not currently shared.", observations2(d))
                                        End If
                                        msg2 = msg2 & " Recommendation: Shared Services can save taxpayer money while helping to standardize process and increase efficiencies across the organization. Contact EITPO's Enterprise Architects to learn how to position systems as shared services."

                                    Case id.futureState
                                        If cnt2 > 9 Then
                                            msg2 = String.Format("Exactly {0} systems in the portfolio were flagged for potential consolidation or replacement based on existing and improved alternatives.", cnt2)
                                        ElseIf cnt2 > 1 Then
                                            str = ReplaceLastOccurrence(observations2(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg2 = String.Format("The following {0} systems in the portfolio were flagged for potential consolidation or replacement based on existing and improved alternatives: {1}.", cnt2, str)
                                        Else
                                            msg2 = String.Format("The [system] {0} was flagged for potential consolidation or replacement based on existing and improved alternatives.", observations2(d))
                                        End If
                                        msg2 = msg2 & " Recommendation: Begin planning a consolidation or replacement strategy for these systems to increase efficiencies across the organization. Leverage existing shared services and/or Cloud services as appropriate."
                                End Select
                            End If
                            If cnt3 > 0 Then
                                Dim str As String = ""
                                Select Case d
                                    Case id.sharedServices
                                        If cnt3 > 9 Then
                                            msg3 = String.Format("Exactly {0} systems in the portfolio currently offer capabilities as services to other organizations.", cnt3)
                                        ElseIf cnt3 > 1 Then
                                            str = ReplaceLastOccurrence(observations3(d), "|", " and")
                                            str = str.Replace("|", ", ")
                                            msg3 = String.Format("The following {0} systems in the portfolio currently offer capabilities as services to other organizations: {1}.", cnt3, str)
                                        Else
                                            msg3 = String.Format("The [system] {0} currently offers capabilities as services to other organizations.", observations3(d))
                                        End If
                                        msg3 = msg3 & "Comment: Shared Services are fundamental to the HHS and CDC IT Strategy."
                                End Select
                            End If

                            If Not msg = String.Empty Or Not msg2 = String.Empty Then
                                Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                                    If Not msg = String.Empty Then
                                        LogText(String.Format("[{0}] - {1}", [Enum].GetName(GetType(id), d) & ".1", msg), w)
                                    End If
                                    If Not msg2 = String.Empty Then
                                        LogText(String.Format("[{0}] - {1}", [Enum].GetName(GetType(id), d) & ".2", msg2), w)
                                    End If
                                    'If Not msg3 = String.Empty Then
                                    '    LogText(String.Format("[{0}] - {1}", [Enum].GetName(GetType(id), d) & ".3", msg3), w)
                                    'End If
                                End Using
                            End If
                        Next
                    End If

                Next '*** Get next worksheet. ***
                oWorkbook.Close()
                bridgeAdded += 1
            Next '*** Get next file. ***

            If nErr > 0 Then
                Me.BackgroundWorker1.ReportProgress(100, "Completed with ERROR(s). See \errlog.txt")
            Else
                Me.BackgroundWorker1.ReportProgress(100, "Completed.")
            End If
            Me.BackgroundWorker1.ReportProgress(100, "Completed.")

        Catch ex As Exception
            Using w As StreamWriter = File.AppendText(apppath & "\errlog.txt")
                Log(String.Format("EXCEPTION: {0}", ex.Message), w)
            End Using
            If Not ex.Message.StartsWith("Excel cannot open") Then
                Throw
            Else
                Exit Try
            End If

        Finally

            '*** Clean up COM objects. ***
            If Not oWorkbooks Is Nothing Then Marshal.FinalReleaseComObject(oWorkbooks)
            If Not oWorkbook Is Nothing Then Marshal.FinalReleaseComObject(oWorkbook)
            If Not oSheets Is Nothing Then Marshal.FinalReleaseComObject(oSheets)
            If Not oSheet Is Nothing Then Marshal.FinalReleaseComObject(oSheet)

            oApp.Quit()
            Marshal.FinalReleaseComObject(oApp)

        End Try

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        '*** Retrieve data from the worker thread and display them on the UI. *** 

        Me.txtOutput.AppendText(e.UserState.ToString & Environment.NewLine)
        Me.lblStatus.Text = String.Format("Percent Complete... {0}%", e.ProgressPercentage)
        Me.pbarProgress.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        '*** This event is automatically triggered when the worker is done. *** 

        MessageBox.Show("Done", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Me.pbarProgress.Visible = False
        Me.lblStatus.Text = "Ready"
        Me.btnImport.Enabled = False

    End Sub

    Private Function GetOSIMFiles(ByVal parentDirectory As String) As List(Of String)

        GetOSIMFiles = Nothing

        If String.IsNullOrEmpty(parentDirectory) = False Then

            Dim oDirInfo As New DirectoryInfo(parentDirectory)
            Dim ret As List(Of String) = Nothing

            '*** Get all the Excel files within the specified directory. ***
            For Each oFileInfo As FileInfo In oDirInfo.GetFiles("*.xls", SearchOption.AllDirectories)

                If ret Is Nothing Then ret = New List(Of String)

                '*** Add full file path to a running list. ***
                ret.Add(oFileInfo.FullName)

            Next

            Return ret

        End If

    End Function

    Private Function GetCellAsString(ByRef oSheet As Excel.Worksheet, ByVal cell As String) As String

        GetCellAsString = Nothing

        If Not oSheet Is Nothing AndAlso String.IsNullOrEmpty(cell) = False Then

            Dim oRange As Excel.Range = oSheet.Range(cell)

            If Not oRange Is Nothing AndAlso Not oRange.Text Is Nothing Then

                If String.IsNullOrEmpty(oRange.Text.ToString.Trim) = False Then
                    GetCellAsString = oRange.Text.ToString.Trim
                End If

            End If

            If Not oRange Is Nothing Then Marshal.ReleaseComObject(oRange)

        End If

    End Function

    Private Function GetCellAsDouble(ByRef oSheet As Excel.Worksheet, ByVal cell As String) As Double?

        Return GetCellAsDouble(oSheet, cell, Nothing, Nothing)

    End Function

    Private Function GetCellAsDouble(ByRef oSheet As Excel.Worksheet, ByVal cell As String, ByVal match As String, ByVal def As Double) As Double?

        GetCellAsDouble = Nothing

        If Not oSheet Is Nothing AndAlso String.IsNullOrEmpty(cell) = False Then

            Dim oRange As Excel.Range = oSheet.Range(cell)

            If Not oRange Is Nothing AndAlso Not oRange.Text Is Nothing Then

                '*** If a number is found, return the number. ***
                '*** If text is found, return nothing unless the text matches an input string. ***
                '*** In that case, return the user-specified default value. ***
                If IsNumeric(oRange.Text) Then
                    GetCellAsDouble = CDbl(oRange.Text)
                Else
                    If String.IsNullOrEmpty(match) = False AndAlso oRange.Text.ToString.Trim = match Then
                        GetCellAsDouble = def
                    End If
                End If

            End If

            If Not oRange Is Nothing Then Marshal.ReleaseComObject(oRange)

        End If

    End Function

    Private Function GetCellAsDate(ByRef oSheet As Excel.Worksheet, ByVal cell As String) As Date?

        GetCellAsDate = Nothing

        If Not oSheet Is Nothing AndAlso String.IsNullOrEmpty(cell) = False Then

            Dim oRange As Excel.Range = oSheet.Range(cell)

            If Not oRange Is Nothing AndAlso Not oRange.Text Is Nothing Then

                If String.IsNullOrEmpty(oRange.Text.ToString.Trim) = False Then

                    Dim val As String = oRange.Text.ToString.Trim

                    '*** The cell value could be anything. ***
                    '*** e.g., "2013", "Unknown", "21-Sep-12" ***

                    If val Like "####" Then

                        '*** If the cell value has a 4-digit number, treat it as a year. ***
                        '*** Return a January 1st date for that year. ***
                        GetCellAsDate = New Date(CInt(val), 1, 1)

                    Else

                        '*** If the cell value matches this format "24-Apr-12", then parse it. ***
                        If val Like "*-???-##" Then
                            GetCellAsDate = Date.ParseExact(val, "d-MMM-yy", New System.Globalization.CultureInfo("en-ca"))
                        End If

                    End If

                End If

            End If

            If Not oRange Is Nothing Then Marshal.ReleaseComObject(oRange)

        End If

    End Function

    Private Function FindLabelRowNumber(ByRef oSheet As Excel.Worksheet, ByVal findString As String, ByVal fromRange As String, ByVal toRange As String) As Integer

        FindLabelRowNumber = -1

        Dim oSearch As Excel.Range
        Dim oFind As Excel.Range

        '*** Look up a string within a user-specified search range. ***
        oSearch = oSheet.Range(fromRange, toRange)
        oFind = oSearch.Find(findString, ,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows,
                Excel.XlSearchDirection.xlNext, False)

        '*** If found, return the row number. ***
        If Not oFind Is Nothing Then FindLabelRowNumber = oFind.Row

        If Not oSearch Is Nothing Then Marshal.ReleaseComObject(oSearch)
        If Not oFind Is Nothing Then Marshal.ReleaseComObject(oFind)

    End Function

    Private Sub ReadData(ByRef oSheet As Excel.Worksheet, ByRef obs() As String, ByRef obs2() As String, ByRef obs3() As String)
        ' GetCell = mExcelApp.Sheets(SheetName).cells(Column, Row).value
        Try
            Dim rowCount As Integer = oSheet.UsedRange.Rows.Count
            Dim columnCount As Integer = oSheet.UsedRange.Columns.Count
            Dim rows As Excel.Range = oSheet.UsedRange.Rows
            Dim counter As Int64
            Dim charsToTrim() As Char = {"."c, " "c}
            'Dim LineRow As Matrix


            For rowNo As Integer = 15 To rowCount
                Dim value_range As Excel.Range = oSheet.Range("A" & rowNo, "R" & rowNo)
                Dim array As Object = value_range.Value2
                '*** Convert the counter into a percent. ***
                counter = CInt(((rowNo - 14) * 100) / (rowCount - 14))
                '*** Telling user that a system was processed. ***
                If Not IsNumeric(array(1, 1)) Then
                    Me.BackgroundWorker1.ReportProgress(counter, String.Format("... analyzing '{0}'", array(1, 1)))
                End If

                For i = 2 To 18
                    If IsNumeric(array(1, i)) Then
                        If (Val(array(1, i)) = 1) Then
                            If Not obs(i - 2) = String.Empty Then obs(i - 2) += "| "
                            obs(i - 2) += Regex.Replace(array(1, 1), "[\d-]", String.Empty).TrimStart().TrimEnd(charsToTrim)
                        End If
                        If (Val(array(1, i)) = 2) Then
                            If Not obs2(i - 2) = String.Empty Then obs2(i - 2) += "| "
                            obs2(i - 2) += Regex.Replace(array(1, 1), "[\d-]", String.Empty).TrimStart().TrimEnd(charsToTrim)
                        End If
                        If (Val(array(1, i)) = 3) Then
                            If Not obs3(i - 2) = String.Empty Then obs3(i - 2) += "| "
                            obs3(i - 2) += Regex.Replace(array(1, 1), "[\d-]", String.Empty).TrimStart().TrimEnd(charsToTrim)
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            nErr = nErr + 1
            Me.BackgroundWorker1.ReportProgress(100, "ERROR encountered processing EXCEL sheet '" & oSheet.Name & ". Processing aborted. See \errlog.txt")
            Using w As StreamWriter = File.AppendText(apppath & "\errlog.txt")
                Log(String.Format("EXCEPTION: Sheet={0} - {1}", oSheet.Name, ex.Message), w)
            End Using
        End Try
    End Sub

    Private Sub ReadSubFunctionsSheet(ByRef oSheet As Excel.Worksheet, ByRef obs() As String)
        Try
            Dim rowCount As Integer = oSheet.UsedRange.Rows.Count
            Dim columnCount As Integer = oSheet.UsedRange.Columns.Count
            Dim rows As Excel.Range = oSheet.UsedRange.Rows
            Dim oFunctions As New List(Of cFunction)
            Dim oSubFunctions As New List(Of cFunction)
            Dim tot As New cFunction
            ' **************************************
            ' EA Business Sub-Functions (tab)
            '***************************************
            Dim vr As Excel.Range = oSheet.Range("C6", "C6")
            Dim words() As String = vr.Value2.Split()
            Dim FYLabel As String = words(2)
            Dim Total As Double = 0


            For rowNo As Integer = 7 To rowCount + 10
                Dim value_range As Excel.Range = oSheet.Range("B" & rowNo, "C" & rowNo)
                Dim array As Object = value_range.Value2
                Dim func As New cFunction

                If Not array(1, 1) = Nothing Then ' Function or Sub-Function Name
                    func.Name = array(1, 1)
                    func.fy = array(1, 2)
                    Total += func.fy
                    oSubFunctions.Add(func)
                End If
            Next

            Dim pct, pct1, pct2 As Double
            Dim str As String
            'Sort by Sub-function's total FY amount descendent
            oSubFunctions = oSubFunctions.OrderByDescending(Function(x) x.fy).ToList

            'pct = oFunctions.Item(0).fy / tot.fy
            'pct2 = oFunctions.Item(1).fy / tot.fy
            'str = "The majority of the budget allocation for " & FYLabel & " aligns to " & oFunctions.Item(0).Name & " (" & FormatPercent(pct) & ") and " & oFunctions.Item(1).Name & " (" & FormatPercent(pct2) & ") "
            'str &= "business functions."
            'obs(0) = str
            pct = oSubFunctions.Item(0).fy / Total
            pct1 = oSubFunctions.Item(1).fy / Total
            pct2 = oSubFunctions.Item(2).fy / Total
            str = "The top three sub-functions for " & FYLabel & " are " & oSubFunctions.Item(0).Name & " (" & FormatPercent(pct) & ")," & oSubFunctions.Item(1).Name & " (" & FormatPercent(pct1) & ") and " & oSubFunctions.Item(2).Name & " (" & FormatPercent(pct2) & ")"
            str &= "."
            obs(0) = str

        Catch ex As Exception
            nErr = nErr + 1
            Me.BackgroundWorker1.ReportProgress(100, "ERROR encountered processing EXCEL sheet '" & oSheet.Name & ". Processing aborted. See \errlog.txt")
            Using w As StreamWriter = File.AppendText(apppath & "\errlog.txt")
                Log(String.Format("EXCEPTION: Sheet={0} - {1}", oSheet.Name, ex.Message), w)
            End Using
        End Try

    End Sub
    Private Sub ReadFunctionsSheet(ByRef oSheet As Excel.Worksheet, ByRef obs() As String)
        Try
            Dim rows As Excel.Range = oSheet.UsedRange.Rows
            Dim rowCount As Integer = oSheet.UsedRange.Rows.Count
            Dim columnCount As Integer = oSheet.UsedRange.Columns.Count
            Dim oFunctions As New List(Of cFunction)
            ' **************************************
            ' EA Business Functions (tab)
            '***************************************
            Dim vr As Excel.Range = oSheet.Range("B5", "B5")
            Dim words() As String = vr.Value2.Split()
            Dim FYLabel As String = words(2)
            Dim Total As Double = 0


            For rowNo As Integer = 7 To rowCount + 10
                Dim value_range As Excel.Range = oSheet.Range("A" & rowNo, "C" & rowNo)
                Dim array As Object = value_range.Value2
                Dim func As New cFunction

                If array(1, 1) = "Grand Total" Then ' This a grand total
                    Exit For
                End If

                If Not array(1, 1) = Nothing Then ' Function or Sub-Function Name
                    func.Name = array(1, 1)
                    func.fy = array(1, 2)
                    Total += func.fy
                    oFunctions.Add(func)
                End If
            Next

            Dim pct, pct1, pct2 As Double
            Dim str As String
            'Sort by Sub-function's total FY amount descendent
            oFunctions = oFunctions.OrderByDescending(Function(x) x.fy).ToList
            pct = oFunctions.Item(0).fy / Total
            pct1 = oFunctions.Item(1).fy / Total
            pct2 = oFunctions.Item(2).fy / Total
            str = "The majority of the budget allocation for " & FYLabel & " aligns to " & oFunctions.Item(0).Name & " (" & FormatPercent(pct) & ") And " & oFunctions.Item(1).Name & " (" & FormatPercent(pct2) & ") "
            str &= "business functions."
            obs(0) = str

        Catch ex As Exception
            nErr = nErr + 1
            Me.BackgroundWorker1.ReportProgress(100, "ERROR encountered processing EXCEL sheet '" & oSheet.Name & ". Processing aborted. See \errlog.txt")
            Using w As StreamWriter = File.AppendText(apppath & "\errlog.txt")
                Log(String.Format("EXCEPTION: Sheet={0} - {1}", oSheet.Name, ex.Message), w)
            End Using
        End Try

    End Sub
    Private Sub ReadSurveillanceSheet(ByRef oSheet As Excel.Worksheet, ByRef obs() As String)
        Try
            Dim rowCount As Integer = oSheet.UsedRange.Rows.Count
            Dim columnCount As Integer = oSheet.UsedRange.Columns.Count
            Dim rows As Excel.Range = oSheet.UsedRange.Rows
            Dim n As Integer = 0
            Dim total As Double
            Dim names As String = ""
            Dim charsToTrim() As Char = {"."c, " "c}

            Dim vrF As Excel.Range = oSheet.Range("G7", "G7")
            Dim words() As String = vrF.Value2.Split()
            Dim FYLabel As String = words(2)

            Dim vr As Excel.Range = oSheet.Range("A8", "A8")
            Dim Org As String = vr.Value2.ToString()


            For rowNo As Integer = 9 To rowCount + 10
                Dim value_range As Excel.Range = oSheet.Range("A" & rowNo, "G" & rowNo)
                Dim array As Object = value_range.Value2

                If array(1, 1) = "Grand Total" Then ' This a grand total
                    total = CDbl(array(1, 7))
                    Exit For
                End If

                If Not array(1, 1).ToString().In("--", "Grand Total") Then ' Function or Sub-Function Name
                    If Not names = String.Empty Then names += "| "
                    names += Regex.Replace(array(1, 1), "[\d-]", String.Empty).TrimStart().TrimEnd(charsToTrim)
                    n = n + 1
                End If
            Next
            Dim msg As String
            If n > 5 Then
                msg = String.Format("There are {0} systems in the portfolio supporting Public Health Surveillance activities for {1} with a total {2} budget of {3}", n, Org, FYLabel, FormatNumber(total))
            ElseIf n > 1 Then
                names = ReplaceLastOccurrence(names, "|", " And")
                names = names.Replace("|", ", ")
                msg = String.Format("The following {0} systems in the portfolio supports Public Health Surveillance activities for {1} with a total {2} budget of {3}: {4}", n, Org, FYLabel, FormatNumber(total), names)
            Else
                msg = String.Format("The [system] {0} supports Public Health Surveillance activities for {1} with an {2} budget of {3}.", names, Org, FYLabel, FormatNumber(total))
            End If
            obs(0) = msg

        Catch ex As Exception
            nErr = nErr + 1
            Me.BackgroundWorker1.ReportProgress(100, "ERROR encountered processing EXCEL sheet '" & oSheet.Name & ". Processing aborted. See \errlog.txt")
            Using w As StreamWriter = File.AppendText(apppath & "\errlog.txt")
                Log(String.Format("EXCEPTION: Sheet={0} - {1}", oSheet.Name, ex.Message), w)
            End Using
        End Try

    End Sub
    Private Sub trace(ByVal str As String)
        Me.txtOutput.AppendText(str & Environment.NewLine)
    End Sub
    Private Function ReplaceLastOccurrence(Source As String, Find As String, Replace As String) As String
        Dim i As Integer = Source.LastIndexOf(Find)
        Return Source.Remove(i, Find.Length()).Insert(i, Replace)
    End Function


    Private Function GetCheckBoxValue(ByRef oSheet As Excel.Worksheet, ByVal cell As String) As Boolean

        GetCheckBoxValue = False

        If Not oSheet Is Nothing AndAlso String.IsNullOrEmpty(cell) = False Then

            Dim oShapes As Excel.Shapes = oSheet.Shapes

            If Not oShapes Is Nothing Then

                For Each oShape As Excel.Shape In oShapes

                    Dim ctr As Object = oShape.OLEFormat.Object

                    If Not ctr Is Nothing AndAlso TypeName(ctr) = "Rectangle" Then
                        If oShape.TopLeftCell.Address.Replace("$", "") = cell Then
                            If String.IsNullOrEmpty(ctr.Text.trim) = False Then GetCheckBoxValue = True
                            Exit For
                        End If
                    End If

                Next

            End If

        End If

    End Function

    Public Shared Sub Log(logMessage As String, w As TextWriter)
        w.WriteLine("{0} {1}: {2}", DateTime.Now.ToLongTimeString(),
            DateTime.Now.ToLongDateString(), logMessage)
    End Sub
    Public Shared Sub LogText(logMessage As String, w As TextWriter)
        w.WriteLine("{0}", logMessage)
        w.WriteLine("-------------------------------")
    End Sub
    Public Function GetFileName(ByVal filepath As String) As String

        'This Function Gets the name of a file without the path or extension.

        'Input:
        '   filepath - Full path/filename of file.
        'Return:
        '   GetFileName - Name of file without the extension.

        'Get indices of characters directly before and after filename
        Dim slashindex As Integer = filepath.LastIndexOf("\")
        Dim dotindex As Integer = filepath.LastIndexOf(".")

        GetFileName = filepath.Substring(slashindex + 1, dotindex - slashindex - 1)
    End Function
    Public Function GetArrayLength(ByRef s As String, ByRef delimiter As Char()) As Integer
        Dim items As String() = s.Split(delimiter)
        Return items.Length
    End Function
    Function FormatNumber(ByVal num As Double) As String
        If (num >= 100000000) Then Return (num / 1000000D).ToString("0.#M")
        If (num >= 1000000) Then Return (num / 1000000D).ToString("0.##M")
        If (num >= 100000) Then Return (num / 1000D).ToString("0.#k")
        If (num >= 10000) Then Return (num / 1000D).ToString("0.##k")

        Return num.ToString("#,0")

    End Function

End Class