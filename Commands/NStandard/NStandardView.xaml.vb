Imports System.Windows
Imports Microsoft.Win32
Imports System.IO
Imports System.Data
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Windows.Controls
Imports System.Reflection
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Xml
Imports System.Windows.Shapes

Namespace Autoform.Commands.NStandard

    Public Class FilterValueItem
        Implements INotifyPropertyChanged

        Private _isSelected As Boolean
        Public Property IsSelected As Boolean
            Get
                Return _isSelected
            End Get
            Set(value As Boolean)
                _isSelected = value
                OnPropertyChanged(NameOf(IsSelected))
            End Set
        End Property

        Public Property Value As String

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
    End Class

    Public Class ParameterMapping
        Public Property Key As String
        Public Property Value As String
    End Class

    Public Class ProgressReport
        Public Property Percentage As Integer
        Public Property Message As String
        Public Property Current As Integer
        Public Property Total As Integer
    End Class

    Public Partial Class NStandardView
        Inherits Window

        Private ReadOnly _commandData As ExternalCommandData
        Private _originalDataTable As DataTable
        Private _isUpdatingFilterCheckboxes As Boolean = False
        Private _uiApp As UIApplication
        Private _doc As Document
        Private ReadOnly _settingsFilePath As String

        Public Sub New(commandData As ExternalCommandData)
            InitializeComponent()
            _commandData = commandData
            AddHandler SelectFileButton.Click, AddressOf SelectFileButton_Click
            AddHandler SyncFabricationButton.Click, AddressOf SyncFabricationButton_Click
            AddHandler ExportExcelButton.Click, AddressOf ExportExcelButton_Click
            AddHandler GenerateSheetsButton.Click, AddressOf GenerateSheetsButton_Click
            AddHandler FilterColumnComboBox.SelectionChanged, AddressOf FilterColumnComboBox_SelectionChanged
            AddHandler ApplyFilterButton.Click, AddressOf ApplyFilterButton_Click
            AddHandler ClearFilterButton.Click, AddressOf ClearFilterButton_Click
            AddHandler SaveSettingsButton.Click, AddressOf SaveSettingsButton_Click

            _uiApp = commandData.Application
            _doc = commandData.Application.ActiveUIDocument.Document

            Dim appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim settingsDir = IO.Path.Combine(appDataPath, "Autoform")
            IO.Directory.CreateDirectory(settingsDir)
            _settingsFilePath = IO.Path.Combine(settingsDir, "ParameterSettings.xml")
            LoadSettings()
        End Sub

        Private Sub SelectFileButton_Click(sender As Object, e As RoutedEventArgs)
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
            If openFileDialog.ShowDialog() = True Then
                Dim filePath As String = openFileDialog.FileName
                ' Store filepath in a resource dictionary to be accessible later
                If Resources.Contains("filePath") Then
                    Resources("filePath") = filePath
                Else
                    Resources.Add("filePath", filePath)
                End If

                Dim sheetNames = get_excel_sheet_names_xl(filePath)
                SheetsPanel.Children.Clear()
                For Each sheetName As String In sheetNames
                    Dim btn As New Button()
                    btn.Content = sheetName
                    btn.Tag = filePath ' Store filePath in the button's Tag
                    btn.Style = CType(FindResource("SheetButtonStyle"), Style)
                    AddHandler btn.Click, AddressOf SheetButton_Click
                    SheetsPanel.Children.Add(btn)
                Next
            End If
        End Sub

        Private Async Sub SheetButton_Click(sender As Object, e As RoutedEventArgs)
            Dim button = TryCast(sender, Button)
            If button Is Nothing OrElse button.Tag Is Nothing Then Return

            Dim sheetName = button.Content.ToString()
            Dim filePath = button.Tag.ToString()

            Await LoadSheetData(filePath, sheetName)
        End Sub

        Private Async Function LoadSheetData(filePath As String, sheetName As String) As Task
            Dispatcher.Invoke(Sub()
                              MainProgressBar.Visibility = System.Windows.Visibility.Visible
                              ProgressTextBlock.Visibility = System.Windows.Visibility.Visible
                              MainProgressBar.Value = 0
                          End Sub)

            _originalDataTable = New DataTable()

            Try
                Await Task.Run(Sub()
                                   _originalDataTable = get_excel_sheet_data_xl(filePath, sheetName)
                               End Sub)


                ' Load data row-by-row into the display DataTable for progressive loading effect
                Dim totalRows = _originalDataTable.Rows.Count
                Dim displayDt As New DataTable()
                For Each col As DataColumn In _originalDataTable.Columns
                    displayDt.Columns.Add(col.ColumnName, col.DataType)
                Next

                Dispatcher.Invoke(Sub()
                                      MainDataGrid.ItemsSource = displayDt.DefaultView
                                  End Sub)


                For i = 0 To totalRows - 1
                    displayDt.ImportRow(_originalDataTable.Rows(i))
                    ' Simulate work
                    Await Task.Delay(1) ' Small delay to make progress visible

                    ' Update progress bar on UI thread
                    Dim percentage = CInt((i + 1) * 100 / totalRows)
                    Dispatcher.Invoke(Sub()
                                          MainProgressBar.Value = percentage
                                          ProgressTextBlock.Text = $"{percentage}%"
                                      End Sub)
                Next

                Dispatcher.Invoke(Sub()
                                      ' Add Select column if it doesn't exist
                                      If Not _originalDataTable.Columns.Contains("Select") Then
                                          _originalDataTable.Columns.Add("Select", GetType(Boolean)).SetOrdinal(0)
                                      End If

                                      ' This is tricky, we need to merge the display table back or re-assign
                                      MainDataGrid.ItemsSource = _originalDataTable.DefaultView

                                      PopulateFilterControls()
                                      SyncFabricationButton.IsEnabled = True
                                      GenerateSheetsButton.IsEnabled = False

                                      ' Hide progress bar
                                      MainProgressBar.Visibility = System.Windows.Visibility.Collapsed
                                      ProgressTextBlock.Visibility = System.Windows.Visibility.Collapsed
                                  End Sub)
            Catch ex As Exception
                MessageBox.Show($"Error loading data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Dispatcher.Invoke(Sub()
                                      ' Hide progress bar in case of error
                                      MainProgressBar.Visibility = System.Windows.Visibility.Collapsed
                                      ProgressTextBlock.Visibility = System.Windows.Visibility.Collapsed
                                  End Sub)
            End Try
        End Function

        Private Sub FilterValueItem_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            If _isUpdatingFilterCheckboxes OrElse e.PropertyName <> "IsSelected" Then Return

            _isUpdatingFilterCheckboxes = True

            Dim changedItem = CType(sender, FilterValueItem)
            Dim items = TryCast(FilterValueComboBox.ItemsSource, List(Of FilterValueItem))
            If items Is Nothing Then
                _isUpdatingFilterCheckboxes = False
                Return
            End If

            If changedItem.Value = "ALL" Then
                ' Check/uncheck all other items based on "ALL" state
                For Each item In items
                    item.IsSelected = changedItem.IsSelected
                Next
            Else
                ' If an individual item is unchecked, uncheck "ALL"
                If Not changedItem.IsSelected Then
                    items.First(Function(i) i.Value = "ALL").IsSelected = False
                Else
                    ' If an individual item is checked, check "ALL" if all other items are also checked
                    If items.Where(Function(i) i.Value <> "ALL").All(Function(i) i.IsSelected) Then
                        items.First(Function(i) i.Value = "ALL").IsSelected = True
                    End If
                End If
            End If

            _isUpdatingFilterCheckboxes = False
        End Sub

        Private Sub PopulateFilterControls()
            FilterColumnComboBox.ItemsSource = Nothing
            FilterValueComboBox.ItemsSource = Nothing
            If _originalDataTable Is Nothing Then Return

            Dim columnNames = New List(Of String)()
            columnNames.Add("ALL")
            For Each col As DataColumn In _originalDataTable.Columns
                columnNames.Add(col.ColumnName)
            Next

            FilterColumnComboBox.ItemsSource = columnNames
            FilterColumnComboBox.SelectedIndex = 0
        End Sub

        Private Sub FilterColumnComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If FilterColumnComboBox.SelectedItem Is Nothing Then Return

            Dim selectedColumn = FilterColumnComboBox.SelectedItem.ToString()
            Dim filterValues = New List(Of FilterValueItem)()

            If selectedColumn = "ALL" Then
                FilterValueComboBox.ItemsSource = Nothing
                Return
            End If

            ' Add "ALL" option
            filterValues.Add(New FilterValueItem With {.IsSelected = True, .Value = "ALL"})

            ' Get unique values from the selected column
            Dim uniqueValues = _originalDataTable.AsEnumerable().Select(Function(row) row(selectedColumn).ToString()).Distinct().OrderBy(Function(x) x).ToList()
            For Each value In uniqueValues
                If value IsNot Nothing Then
                    filterValues.Add(New FilterValueItem With {.IsSelected = True, .Value = value})
                End If
            Next

            For Each item In filterValues
                AddHandler item.PropertyChanged, AddressOf FilterValueItem_PropertyChanged
            Next

            FilterValueComboBox.ItemsSource = filterValues
        End Sub

        Private Sub ApplyFilterButton_Click(sender As Object, e As RoutedEventArgs)
            If _originalDataTable Is Nothing Then Return

            ' If no column is selected or "ALL" is selected, clear the filter.
            If FilterColumnComboBox.SelectedItem Is Nothing OrElse FilterColumnComboBox.SelectedItem.ToString() = "ALL" Then
                _originalDataTable.DefaultView.RowFilter = String.Empty
                Return
            End If

            Dim selectedColumn = FilterColumnComboBox.SelectedItem.ToString()
            Dim itemsSource = TryCast(FilterValueComboBox.ItemsSource, List(Of FilterValueItem))

            If itemsSource Is Nothing Then
                _originalDataTable.DefaultView.RowFilter = String.Empty
                Return
            End If

            ' Get all checked values, excluding "ALL" itself.
            Dim selectedValues = itemsSource _
                .Where(Function(item) item.IsSelected AndAlso item.Value <> "ALL") _
                .Select(Function(item) item.Value).ToList()
        
            ' If no specific values are selected, clear the filter.
            If selectedValues.Count = 0 Then
                _originalDataTable.DefaultView.RowFilter = String.Empty
                Return
            End If
            
            ' Apply the filter to the DefaultView of the original table.
            _originalDataTable.DefaultView.RowFilter = $"[{selectedColumn}] IN ({String.Join(",", selectedValues.Select(Function(v) $"'{v.Replace("'", "''")}'"))})"
        End Sub

        Private Function get_excel_sheet_names_xl(filePath As String) As List(Of String)
            Dim excelApp As Excel.Application = Nothing
            Dim workbook As Excel.Workbook = Nothing
            Dim sheetNames As New List(Of String)()
            Try
                excelApp = New Excel.Application()
                workbook = excelApp.Workbooks.Open(filePath)
                For Each sheet As Excel.Worksheet In workbook.Sheets
                    sheetNames.Add(sheet.Name)
                Next
            Catch ex As Exception
                MessageBox.Show($"Error reading Excel file: {ex.Message}", "Excel Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Finally
                If workbook IsNot Nothing Then
                    workbook.Close(False)
                    Marshal.ReleaseComObject(workbook)
                End If
                If excelApp IsNot Nothing Then
                    excelApp.Quit()
                    Marshal.ReleaseComObject(excelApp)
                End If
            End Try
            Return sheetNames
        End Function

        Private Function get_excel_sheet_data_xl(filePath As String, sheetName As String) As DataTable
            Dim dt As New DataTable()
            Dim excelApp As Excel.Application = Nothing
            Dim workbook As Excel.Workbook = Nothing
            Dim worksheet As Excel.Worksheet = Nothing
            Dim range As Excel.Range = Nothing

            Try
                excelApp = New Excel.Application()
                workbook = excelApp.Workbooks.Open(filePath)
                worksheet = CType(workbook.Sheets(sheetName), Excel.Worksheet)
                range = worksheet.UsedRange

                Dim rowCount As Integer = range.Rows.Count
                Dim colCount As Integer = range.Columns.Count

                ' Get header
                For c As Integer = 1 To colCount
                    Dim cellValue = CType(range.Cells(1, c), Excel.Range).Text
                    dt.Columns.Add(If(Not String.IsNullOrEmpty(cellValue), cellValue, $"Column_{c}"))
                Next

                ' Get rows
                For r As Integer = 2 To rowCount
                    Dim row As DataRow = dt.NewRow()
                    For c As Integer = 1 To colCount
                        row(c - 1) = CType(range.Cells(r, c), Excel.Range).Text
                    Next
                    dt.Rows.Add(row)
                Next
            Catch ex As Exception
                MessageBox.Show($"Error reading Excel sheet data: {ex.Message}", "Excel Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Finally
                ' Release COM objects
                If range IsNot Nothing Then Marshal.ReleaseComObject(range)
                If worksheet IsNot Nothing Then Marshal.ReleaseComObject(worksheet)
                If workbook IsNot Nothing Then
                    workbook.Close(False)
                    Marshal.ReleaseComObject(workbook)
                End If
                If excelApp IsNot Nothing Then
                    excelApp.Quit()
                    Marshal.ReleaseComObject(excelApp)
                End If
            End Try
            Return dt
        End Function

        Private Function get_csv_data(filePath As String) As DataTable
            Dim dt As New DataTable()
            Using reader As New IO.StreamReader(filePath)
                ' Read header
                If Not reader.EndOfStream Then
                    Dim headers As String() = reader.ReadLine().Split(","c)
                    For Each header As String In headers
                        dt.Columns.Add(header)
                    Next
                End If

                ' Read rows
                While Not reader.EndOfStream
                    Dim rows As String() = reader.ReadLine().Split(","c)
                    dt.Rows.Add(rows)
                End While
            End Using
            Return dt
        End Function

        Private Sub SyncFabricationButton_Click(sender As Object, e As RoutedEventArgs)
            If _originalDataTable Is Nothing Then
                MessageBox.Show("Please load data from an Excel sheet first.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim doc As Document = _commandData.Application.ActiveUIDocument.Document

            ' Ensure required columns exist.
            If Not _originalDataTable.Columns.Contains("Sync Status") Then
                _originalDataTable.Columns.Add("Sync Status", GetType(String))
            End If
            If Not _originalDataTable.Columns.Contains("sample drawing name") Then
                _originalDataTable.Columns.Add("sample drawing name", GetType(String))
            End If
            If Not _originalDataTable.Columns.Contains("Select") Then
                _originalDataTable.Columns.Add("Select", GetType(Boolean)).SetOrdinal(0)
            End If

            ' Get sheets from Revit
            Dim filteredSheets = New FilteredElementCollector(doc).OfClass(GetType(ViewSheet)).Cast(Of ViewSheet)().Where(Function(s)
                                                                                                                                Dim param = s.LookupParameter("Sub-Category")
                                                                                                                                Return param IsNot Nothing AndAlso param.AsString() = "Sample Drawings"
                                                                                                                            End Function).ToList()

            ' Parse sheet names into a more usable format
            Dim revitSheetData = New List(Of (Code As String, MinRange As Double, MaxRange As Double, FullName As String))()
            For Each sheet In filteredSheets
                Dim parts = sheet.Name.Split("-"c)
                If parts.Length = 3 Then
                    Dim code = parts(0)
                    Dim minRange As Double
                    Dim maxRange As Double
                    If Double.TryParse(parts(1), minRange) AndAlso Double.TryParse(parts(2), maxRange) Then
                        revitSheetData.Add((code, minRange, maxRange, sheet.Name))
                    End If
                End If
            Next

            ' Create a copy to modify. This helps in forcing a UI update.
            Dim updatedDataTable = _originalDataTable.Copy()

            ' Perform the comparison on the rows of the copied table
            For Each row As DataRow In updatedDataTable.Rows
                Dim codeValue = row("Code")?.ToString()
                If String.IsNullOrEmpty(codeValue) Then
                    row("Sync Status") = "NoMatch"
                    row("Select") = False
                    Continue For
                End If

                Dim length As Double
                If updatedDataTable.Columns.Contains("Length") AndAlso Double.TryParse(row("Length")?.ToString(), length) Then
                    ' Use Length if available and valid
                ElseIf updatedDataTable.Columns.Contains("Arm A") AndAlso Double.TryParse(row("Arm A")?.ToString(), length) Then
                    ' Otherwise use Arm A if available and valid
                Else
                    ' No valid length found for this row, mark as no match and continue
                    row("Sync Status") = "NoMatch"
                    row("Select") = False
                    Continue For
                End If

                Dim isMatch As Boolean = False
                For Each revitSheet In revitSheetData
                    If codeValue.Equals(revitSheet.Code, StringComparison.OrdinalIgnoreCase) AndAlso length >= revitSheet.MinRange AndAlso length <= revitSheet.MaxRange Then
                        row("Sync Status") = "Match"
                        row("sample drawing name") = revitSheet.FullName
                        row("Select") = True
                        isMatch = True
                        Exit For ' Found a match, no need to check other sheets
                    End If
                Next

                If Not isMatch Then
                    row("Sync Status") = "NoMatch"
                    row("Select") = False
                    row("sample drawing name") = String.Empty
                End If
            Next

            ' Replace the old table with the updated one and re-bind
            _originalDataTable = updatedDataTable
            MainDataGrid.ItemsSource = _originalDataTable.DefaultView

            GenerateSheetsButton.IsEnabled = True
            MessageBox.Show("Synchronization complete.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub

        Private Class ParameterSettingsData
            Public Property DrawnBy As String
            Public Property CheckedBy As String
            Public Property DesignedBy As String
            Public Property ApprovedBy As String
            Public Property UnitColumn As String
            Public Property LotNoColumn As String
            Public Property DrawingNumbersColumn As String
        End Class

        Private Class GenericParameterInstruction
            Public Property PossibleNames As List(Of String)
            Public Property ValueInFeet As Double
        End Class

        Private Class SheetGenerationData
            Public Property SampleSheetName As String
            Public Property NewSheetName As String
            Public Property SheetCode As String
            Public Property SheetParameters As Dictionary(Of String, String)
            Public Property GeometryDetailParameters As Dictionary(Of String, Double) ' For Arm A, Height etc.
            Public Property GenericDetailParameters As List(Of GenericParameterInstruction) ' For P1, P2, etc.
        End Class

        Private Async Function CreateGenerationData(selectedRows As List(Of DataRowView), settings As ParameterSettingsData) As Task(Of List(Of SheetGenerationData))
            Dim generationTasks = New List(Of SheetGenerationData)()
            Await Task.Run(Sub()
                               For Each rowView In selectedRows
                                   ' Pre-process and validate all data for this row
                                   Dim sheetData = New SheetGenerationData With {
                                .SampleSheetName = rowView("sample drawing name")?.ToString(),
                                .NewSheetName = rowView("Title")?.ToString(),
                                .SheetCode = rowView("Code")?.ToString(),
                                .SheetParameters = New Dictionary(Of String, String)(),
                                .GeometryDetailParameters = New Dictionary(Of String, Double)(),
                                .GenericDetailParameters = New List(Of GenericParameterInstruction)()
                            }

                                   ' --- Static Sheet Parameters ---
                                   sheetData.SheetParameters("Drawn By") = settings.DrawnBy
                                   sheetData.SheetParameters("Checked By") = settings.CheckedBy
                                   sheetData.SheetParameters("Designed By") = settings.DesignedBy
                                   sheetData.SheetParameters("Approved By") = settings.ApprovedBy

                                   ' --- Dynamic Sheet Parameters from Columns ---
                                   Dim addSheetParam = Sub(paramName As String, colName As String)
                                                           If Not String.IsNullOrWhiteSpace(colName) AndAlso rowView.Row.Table.Columns.Contains(colName) Then
                                                               Dim value = rowView.Row(colName)?.ToString()
                                                               If Not String.IsNullOrEmpty(value) Then
                                                                   sheetData.SheetParameters(paramName) = value
                                                               End If
                                                           End If
                                                       End Sub
                                   addSheetParam("Sub-Category", settings.UnitColumn)
                                   addSheetParam("Unit", settings.UnitColumn)
                                   addSheetParam("Lot No", settings.LotNoColumn)
                                   addSheetParam("Drawing Numbers", settings.DrawingNumbersColumn)
                                   sheetData.SheetParameters("Drawing Type") = "Gen_Fabrication"


                                   ' --- Detail Item Parameters (Pre-validated) ---
                                   Dim addDetailParam = Sub(paramName As String, colName As String)
                                                            If Not String.IsNullOrWhiteSpace(colName) AndAlso rowView.Row.Table.Columns.Contains(colName) Then
                                                                Dim valueStr = rowView.Row(colName)?.ToString()
                                                                Dim numericValue As Double
                                                                If Double.TryParse(valueStr, numericValue) Then
                                                                    sheetData.GeometryDetailParameters(paramName) = numericValue / 304.8 ' Convert mm to feet
                                                                End If
                                                            End If
                                                        End Sub

                                   Dim geometryType = rowView("Geometry Type")?.ToString()
                                   Select Case geometryType
                                       Case "Corner-CH"
                                           addDetailParam("Arm A", "Arm A")
                                           addDetailParam("Arm B", "Arm B")
                                           addDetailParam("Height", "Height")
                                       Case "Corner-Sec"
                                           addDetailParam("Arm A", "Arm A")
                                           addDetailParam("Arm B", "Arm B")
                                           addDetailParam("Height", "Height")
                                           addDetailParam("Width", "Width")
                                       Case "Rect."
                                           addDetailParam("Height", "Height")
                                           addDetailParam("Length", "Length")
                                       Case "Sec-2H"
                                           addDetailParam("Height", "Height")
                                           addDetailParam("Length", "Length")
                                           addDetailParam("Width", "Width")
                                   End Select

                                   ' --- NEW: Generic P-column parameter processing ---
                                   For i = 1 To 50 ' Check for up to 50 P-columns
                                       Dim pColName = $"P{i}"
                                       If rowView.Row.Table.Columns.Contains(pColName) Then
                                           Dim cellValue = rowView.Row(pColName)?.ToString()
                                           If String.IsNullOrWhiteSpace(cellValue) OrElse Not cellValue.Contains("-") Then Continue For

                                           Dim parts = cellValue.Split("-"c)
                                           If parts.Length <> 2 Then Continue For

                                           Dim prefix = parts(0).Trim().ToUpper()
                                           Dim valueStr = parts(1).Trim()
                                           Dim numericValue As Double
                                           If Not Double.TryParse(valueStr, numericValue) Then Continue For

                                           Dim instruction = New GenericParameterInstruction With {
                                                .ValueInFeet = numericValue / 304.8,
                                                .PossibleNames = New List(Of String)()
                                            }

                                           Select Case prefix
                                               Case "B" : instruction.PossibleNames.Add($"BH#{i}")
                                               Case "T" : instruction.PossibleNames.Add($"TH#{i}")
                                               Case "S"
                                                   instruction.PossibleNames.Add($"ST#{i}")
                                                   instruction.PossibleNames.Add($"Stiffner_{i}")
                                               Case "H"
                                                   instruction.PossibleNames.Add($"SL#{i}")
                                                   instruction.PossibleNames.Add($"SH#{i}")
                                               Case "A1B" : instruction.PossibleNames.Add($"A1_BH#{i}")
                                               Case "A2B" : instruction.PossibleNames.Add($"A2_BH#{i}")
                                               Case "A1T" : instruction.PossibleNames.Add($"A1_TH#{i}")
                                               Case "A2T" : instruction.PossibleNames.Add($"A2_TH#{i}")
                                               Case "P" : instruction.PossibleNames.Add($"P#{i}")
                                           End Select

                                           If instruction.PossibleNames.Any() Then
                                               sheetData.GenericDetailParameters.Add(instruction)
                                           End If
                                       End If
                                   Next


                                   generationTasks.Add(sheetData)
                               Next
                           End Sub)
            Return generationTasks
        End Function

        Private Async Sub GenerateSheetsButton_Click(sender As Object, e As RoutedEventArgs)
            Dim stopwatch As New System.Diagnostics.Stopwatch()
            Dim progress As New Progress(Of ProgressReport)(
                Sub(report)
                    MainProgressBar.Value = report.Percentage
                    Dim remainingStr = ""
                    If stopwatch.IsRunning AndAlso report.Percentage > 0 Then
                        Dim elapsed = stopwatch.Elapsed
                        Dim totalEstimatedTime = TimeSpan.FromMilliseconds(elapsed.TotalMilliseconds / report.Percentage * 100)
                        Dim remainingTime = totalEstimatedTime - elapsed
                        remainingStr = $" | Estimated Time {remainingTime:mm\:ss} "
                    End If
                    ProgressTextBlock.Text = $"[ Done {report.Percentage}% {report.Current}/{report.Total}{remainingStr}]"
                End Sub)

            MainProgressBar.Visibility = System.Windows.Visibility.Visible
            ProgressTextBlock.Visibility = System.Windows.Visibility.Visible

            Dim selectedRows = CType(MainDataGrid.ItemsSource, DataView).Table.AsEnumerable().Where(Function(r) r.Field(Of Boolean)("Select")).ToList()
            Dim selectedData As DataTable = If(selectedRows.Any(), selectedRows.CopyToDataTable(), New DataTable())

            Dim settings = New Dictionary(Of String, String) From {
                {"DrawnBy", DrawnByTextBox.Text},
                {"CheckedBy", CheckedByTextBox.Text},
                {"DesignedBy", DesignedByTextBox.Text},
                {"ApprovedBy", ApprovedByTextBox.Text},
                {"UnitColumn", UnitColumnTextBox.Text},
                {"LotNoColumn", LotNoColumnTextBox.Text},
                {"DrawingNumbersColumn", DrawingNumbersColumnTextBox.Text}
            }

            ' Pass the UI's dispatcher to the command
            Dim command As New NStandardCommand With {
                .UIDispatcher = Me.Dispatcher
            }
            stopwatch.Start()
            Dim generatedCount = Await command.ExecuteSheetGeneration(_commandData, selectedData, progress, settings)
            stopwatch.Stop()

            MainProgressBar.Visibility = System.Windows.Visibility.Collapsed
            ProgressTextBlock.Visibility = System.Windows.Visibility.Collapsed
            Dim elapsedTime = stopwatch.Elapsed.ToString("g")
            MessageBox.Show($"Sheet generation complete!{Environment.NewLine}{generatedCount} sheets created in {elapsedTime}", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub

        Private Sub ExportExcelButton_Click(sender As Object, e As RoutedEventArgs)
            If MainDataGrid.ItemsSource Is Nothing Then
                MessageBox.Show("There is no data to export.", "No Data", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            Dim docPath = _doc.PathName
            If String.IsNullOrEmpty(docPath) Then
                MessageBox.Show("Please save the Revit document first to specify a location for the export.", "Save Document", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim exportPath = IO.Path.Combine(IO.Path.GetDirectoryName(docPath), "ExportedData.xlsx")
            Dim excelApp As Excel.Application = Nothing
            Dim workbook As Excel.Workbook = Nothing
            Dim worksheet As Excel.Worksheet = Nothing

            Try
                If IO.File.Exists(exportPath) Then
                    IO.File.Delete(exportPath)
                End If
            Catch ex As IO.IOException
                MessageBox.Show($"Could not overwrite the existing export file. Please ensure 'ExportedData.xlsx' is closed and try again.{vbCrLf}{vbCrLf}Details: {ex.Message}", "Export File In Use", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End Try

            Try
                excelApp = New Excel.Application()
                excelApp.Visible = False
                workbook = excelApp.Workbooks.Add()
                worksheet = CType(workbook.ActiveSheet, Excel.Worksheet)

                For j As Integer = 0 To MainDataGrid.Columns.Count - 1
                    worksheet.Cells(1, j + 1) = MainDataGrid.Columns(j).Header.ToString()
                Next

                For i As Integer = 0 To MainDataGrid.Items.Count - 1
                    Dim rowView = CType(MainDataGrid.Items(i), DataRowView)
                    For j As Integer = 0 To MainDataGrid.Columns.Count - 1
                        worksheet.Cells(i + 2, j + 1) = rowView.Row(j).ToString()
                    Next

                    If rowView.Row.Table.Columns.Contains("Sync Status") Then
                        Dim status = rowView.Row("Sync Status").ToString()
                        Dim range = CType(worksheet.Rows(i + 2), Excel.Range)
                        If status = "Match" Then
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        ElseIf status = "NoMatch" Then
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral)
                        End If
                    End If
                Next
                
                worksheet.Columns.AutoFit()
                workbook.SaveAs(exportPath)
                MessageBox.Show($"Data exported successfully to:{vbCrLf}{exportPath}", "Export Successful", MessageBoxButton.OK, MessageBoxImage.Information)

            Catch ex As Exception
                MessageBox.Show($"An error occurred during export: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Finally
                If workbook IsNot Nothing Then
                    workbook.Close(False)
                    Marshal.ReleaseComObject(workbook)
                End If
                If excelApp IsNot Nothing Then
                    excelApp.Quit()
                    Marshal.ReleaseComObject(excelApp)
                End If
            End Try
        End Sub

        Private Sub ClearFilterButton_Click(sender As Object, e As RoutedEventArgs)
            If _originalDataTable IsNot Nothing Then
                _originalDataTable.DefaultView.RowFilter = String.Empty
                FilterColumnComboBox.SelectedIndex = -1
                FilterValueComboBox.ItemsSource = Nothing
            End If
        End Sub

        Private Sub SaveSettingsButton_Click(sender As Object, e As RoutedEventArgs)
            SaveSettings()
        End Sub

        Private Sub LoadSettings()
            If Not IO.File.Exists(_settingsFilePath) Then Return

            Try
                Dim doc As New XmlDocument()
                doc.Load(_settingsFilePath)
                DrawnByTextBox.Text = doc.SelectSingleNode("//DrawnBy")?.InnerText
                CheckedByTextBox.Text = doc.SelectSingleNode("//CheckedBy")?.InnerText
                DesignedByTextBox.Text = doc.SelectSingleNode("//DesignedBy")?.InnerText
                ApprovedByTextBox.Text = doc.SelectSingleNode("//ApprovedBy")?.InnerText
                UnitColumnTextBox.Text = doc.SelectSingleNode("//UnitColumn")?.InnerText
                LotNoColumnTextBox.Text = doc.SelectSingleNode("//LotNoColumn")?.InnerText
                DrawingNumbersColumnTextBox.Text = doc.SelectSingleNode("//DrawingNumbersColumn")?.InnerText
            Catch ex As Exception
                MessageBox.Show($"Error loading settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            End Try
        End Sub

        Private Sub SaveSettings()
            Try
                Dim doc As New XmlDocument()
                Dim root = doc.CreateElement("Settings")
                doc.AppendChild(root)

                Dim addElement = Sub(name As String, value As String)
                                     Dim elem = doc.CreateElement(name)
                                     elem.InnerText = value
                                     root.AppendChild(elem)
                                 End Sub

                addElement("DrawnBy", DrawnByTextBox.Text)
                addElement("CheckedBy", CheckedByTextBox.Text)
                addElement("DesignedBy", DesignedByTextBox.Text)
                addElement("ApprovedBy", ApprovedByTextBox.Text)
                addElement("UnitColumn", UnitColumnTextBox.Text)
                addElement("LotNoColumn", LotNoColumnTextBox.Text)
                addElement("DrawingNumbersColumn", DrawingNumbersColumnTextBox.Text)

                doc.Save(_settingsFilePath)
                MessageBox.Show("Settings saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub
    End Class
End Namespace 