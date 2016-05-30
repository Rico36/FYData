Imports System.IO
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Class frmMain

    Private apppath As String = ""
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
        '*** Define a list of Excel worksheet names to exlude from processing. ***
        If apppath = "" Then
            Dim dialog As New FolderBrowserDialog()
            dialog.RootFolder = Environment.SpecialFolder.Desktop
            dialog.SelectedPath = "C:\"
            dialog.ShowNewFolderButton = False
            dialog.Description = "Select the directory with the FY data (Excel files)"
            If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                apppath = dialog.SelectedPath
            Else
                Application.Exit()
            End If

        End If

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
            Dim files As List(Of String) = GetOSIMFiles("C:\Temp")

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

                System.Windows.Forms.MessageBox.Show("There is no file to process in the specified directory.",
                "Action Terminated", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)

                Me.btnImport.Enabled = True
                Me.txtOutput.Text = "Action terminated."

            End If

        Catch ex As Exception

            System.Windows.Forms.MessageBox.Show(ex.ToString, "Application Error",
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
                '*** Iterate through all the Excel worksheets. ***
                For i As Integer = 1 To oSheets.Count

                    oSheet = CType(oSheets(i), Excel.Worksheet)
                    sheetName = oSheet.Name.Trim

                    ' Get observations
                    If sheetName = "IT System Assessment Details" Then

                        '*** Proceed here only if the worksheet is the AssessmentData. ***
                        ReadData(oSheet, observations, observations2)

                        For d As Integer = 0 To 16
                            Dim cnt As Integer = 0
                            Dim cnt2 As Integer = 0
                            Dim msg As String = ""
                            Dim msg2 As String = ""

                            If Not IsNothing(observations(d)) Then
                                cnt = GetArrayLength(observations(d), "|")
                            End If
                            If Not IsNothing(observations2(d)) Then
                                cnt2 = GetArrayLength(observations2(d), "|")
                            End If
                            If cnt > 0 Then
                                Dim str As String = ""
                                Select Case d
                                    Case id.strategicAlignment
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio provides minimal or no contribution to the CDC and/or CIO's strategic objectives.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} system(s) in the portfolio provides minimal or no contribution to the CDC and/or CIO's strategic objectives:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was assessed to provide minimal contribution to the CDC and/or CIO's strategic objectives.", observations(d))
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation: Assess the actual business need during the next annual operational analysis and/or begin planning the retirement of these systems ."

                                    Case id.efficiency
                                        Exit Select

                                    Case id.changeReq
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio have been flagged with operational and/or technical deficiencies which prevent meeting the business needs.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} system(s) in the portfolio have been flagged with operational and/or technical deficiencies which prevent meeting the business needs:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was flagged with operational and/or technical deficiencies which prevent meeting the business needs.", observations(d))
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation:  An analisys of alternatives (AoA) would help determine a suitable solution replacement and/or needed technology refresh."

                                    Case id.benefit
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio are not meeting the business benefits and/or business needs.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} system(s) in the portfolio are not meeting the business benefits and/or business needs:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} is not meeting the business benefits and/or business needs.", observations(d))
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation:  Begin planning retirement for these systems or alternatively consider performing a gap analysis during the next OA to identify suitable replacement."
                                    Case id.sysEnhancement
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio were deemed rigid and un-feasible to undergo enhancements when new business requirements emerge.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} system(s) in the portfolio were deemed rigid and un-feasible to enhance/modify when new business requirements emerge:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was deemed rigid and un-feasible to enhance/modify when new business requirements emerge.", observations(d))
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation:  Systems that cannot adapt to emerging business requirements represent a risk for the organization and should be considered for a technology refresh or solution replacement."


                                    Case id.futureState
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio have been flagged as candidate for retirement during the latest FY IT assessment.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} system(s) in the portfolio were flagged as candidate for retirement during the latest FY IT assessment:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} was flagged for retirement during the latest FY IT assessment.", observations(d))
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation:  Revisit the latest operational analysis for these systems to confirm obsolence and begin a plan for their retirement."

                                    Case id.techRelevance
                                        If cnt > 3 Then
                                            msg = String.Format("Exactly {0} systems in the portfolio are running on outdated or obsolete technologies.", cnt)
                                        ElseIf cnt > 1 Then
                                            str = ReplaceLastOccurrence(observations(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg = String.Format("The following {0} systems in the portfolio are at risk for operational interruptions given their dependencies on outdated and obsolete technologies:  {1}", cnt, str)
                                        Else
                                            msg = String.Format("The [system] {0} is at risk for operational interruptions given its reliance on outdated and obsolete technologies.", str)
                                        End If
                                        msg = msg & Environment.NewLine & "Recommendation:  Begin planning a technology refresh or solution replacement to ensure proper continuity of services."

                                End Select
                            End If
                            If cnt2 > 0 Then
                                Dim str As String = ""
                                Select Case d
                                    Case id.sharedServices
                                        If cnt2 > 3 Then
                                            msg2 = String.Format("Exactly {0} systems in the portfolio have functionality that would potentially be valuable to other organizations but are not currently shared.", cnt2)
                                        ElseIf cnt2 > 1 Then
                                            str = ReplaceLastOccurrence(observations2(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg2 = String.Format("The following {0} systems in the portfolio have capabilities that would be valuable to other organizations but are not currently shared.:  {1}", cnt2, str)
                                        Else
                                            msg2 = String.Format("The [system] {0} has functionality potentially valuable to other organizations but these are not currently shared.", observations2(d))
                                        End If
                                        msg2 = msg2 & Environment.NewLine & "Recommendation: Shared Services can save money to the taxpayers while helping to standardize process and increase effciencies across the organization. Contact EITPO's Enterprise Architects to learn how to position some of these systems as shared services."
                                    Case id.futureState
                                        If cnt2 > 3 Then
                                            msg2 = String.Format("Exactly {0} systems in the portfolio were flaged for potential consolidation or replacement based on existing and improved alternatives.", cnt2)
                                        ElseIf cnt2 > 1 Then
                                            str = ReplaceLastOccurrence(observations2(d), "|", " and")
                                            str = str.Replace("|", ",")
                                            msg2 = String.Format("The following {0} systems in the portfolio were flaged for potential consolidation or replacement based on existing and improved alternatives:  {1}", cnt2, str)
                                        Else
                                            msg2 = String.Format("The [system] {0} was flaged for potential consolidation or replacement based on existing and improved alternatives.", observations2(d))
                                        End If
                                        msg2 = msg2 & Environment.NewLine & "Recommendation: Begin planning a consolidation or replacement strategy for these systems to increase effciencies across the organization."
                                End Select
                            End If

                            If Not msg = String.Empty Or Not msg2 = String.Empty Then
                                Using w As StreamWriter = File.AppendText(apppath & "\" & GetFileName(f) & ".txt")
                                    If Not msg = String.Empty Then
                                        LogText(String.Format("[{0}] - {1}", [Enum].GetName(GetType(id), d), msg), w)
                                    End If
                                    If Not msg2 = String.Empty Then
                                        LogText(String.Format("[{0}] - {1}", [Enum].GetName(GetType(id), d) & ".2", msg2), w)
                                    End If

                                End Using
                            End If
                        Next
                    End If
                    Array.Clear(observations, 0, observations.Length)
                    Array.Clear(observations2, 0, observations.Length)
                Next '*** Get next worksheet. ***

                oWorkbook.Close()

                bridgeAdded += 1

            Next '*** Get next file. ***


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

    Private Sub ReadData(ByRef oSheet As Excel.Worksheet, ByRef obs() As String, ByRef obs2() As String)
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
                    Me.BackgroundWorker1.ReportProgress(counter, String.Format("... Processing '{0}'", array(1, 1)))
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
                    End If
                Next
            Next
        Catch ex As Exception
            trace(ex.Message)
        End Try
        Me.BackgroundWorker1.ReportProgress(100, "Completed.")
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
End Class