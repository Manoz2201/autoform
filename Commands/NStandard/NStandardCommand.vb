#Region "Imported Namespaces"
Imports System
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports System.Data
Imports System.Windows.Threading
Imports System.Threading.Tasks
#End Region

Namespace Autoform.Commands.NStandard
    <Transaction(TransactionMode.Manual)>
    Public Class NStandardCommand
        Implements IExternalCommand

        Public Property UIDispatcher As Dispatcher

        Public Function Execute(
          ByVal commandData As ExternalCommandData,
          ByRef message As String,
          ByVal elements As ElementSet) _
        As Result Implements IExternalCommand.Execute

            Try
                Dim view As New NStandardView(commandData)
                view.ShowDialog()
                Return Result.Succeeded
            Catch ex As Exception
                message = ex.Message
                Return Result.Failed
            End Try
        End Function

        Public Async Function ExecuteSheetGeneration(commandData As ExternalCommandData, selectedData As DataTable, progress As IProgress(Of ProgressReport), settings As Dictionary(Of String, String)) As Task(Of Integer)
            Dim doc As Document = commandData.Application.ActiveUIDocument.Document
            If selectedData Is Nothing OrElse selectedData.Rows.Count = 0 Then
                Return 0
            End If

            Dim totalRows = selectedData.Rows.Count
            Dim generatedCount = 0

            Using tg As New TransactionGroup(doc, "Generate Fabrication Sheets")
                tg.Start()

                For i = 0 To totalRows - 1
                    Dim row = selectedData.Rows(i)
                    Using t As New Transaction(doc, "Generate Sheet")
                        t.Start()
                        Try
                            Dim success = GenerateSingleSheet(doc, row, settings)

                            If success Then
                                t.Commit()
                                generatedCount += 1
                            Else
                                t.RollBack()
                            End If
                        Catch ex As Exception
                            t.RollBack()
                        End Try
                    End Using

                    Dim percentage = CInt(((i + 1) * 100) / totalRows)
                    progress.Report(New ProgressReport With {
                        .Percentage = percentage,
                        .Message = $"Processing row {i + 1} of {totalRows}",
                        .Current = i + 1,
                        .Total = totalRows
                    })
                    Await Task.Delay(1)
                Next

                tg.Assimilate()
            End Using
            Return generatedCount
        End Function

        Private Function GenerateSingleSheet(doc As Document, row As DataRow, settings As Dictionary(Of String, String)) As Boolean
            Try
                Dim sampleSheetName = row("sample drawing name")?.ToString()
                Dim newSheetTitle = row("Title")?.ToString()
                Dim baseSheetCode = row("Code")?.ToString()

                If String.IsNullOrEmpty(sampleSheetName) OrElse String.IsNullOrEmpty(newSheetTitle) OrElse String.IsNullOrEmpty(baseSheetCode) Then
                    Return False
                End If

                Dim existingSheetNumbers = New FilteredElementCollector(doc).OfClass(GetType(ViewSheet)).Cast(Of ViewSheet)().Select(Function(s) s.SheetNumber).ToHashSet(StringComparer.OrdinalIgnoreCase)
                Dim existingViewNames = New FilteredElementCollector(doc).OfClass(GetType(View)).Cast(Of View)().Select(Function(v) v.Name).ToHashSet(StringComparer.OrdinalIgnoreCase)
                Dim codeCounters = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                Dim sampleSheet As ViewSheet = New FilteredElementCollector(doc).OfClass(GetType(ViewSheet)).Cast(Of ViewSheet)().FirstOrDefault(Function(s) s.Name.Equals(sampleSheetName, StringComparison.OrdinalIgnoreCase))

                Dim titleBlockTypeId As ElementId = ElementId.InvalidElementId
                If sampleSheet IsNot Nothing Then
                    Dim titleBlockCollector = New FilteredElementCollector(doc, sampleSheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(GetType(FamilyInstance))
                    If titleBlockCollector.Any() Then
                        titleBlockTypeId = titleBlockCollector.FirstElement().GetTypeId()
                    End If
                End If

                If sampleSheet Is Nothing OrElse titleBlockTypeId = ElementId.InvalidElementId Then
                    Return False
                End If

                Dim nextNum = If(codeCounters.ContainsKey(baseSheetCode), codeCounters(baseSheetCode), 1)
                Dim newSheetNumber As String = String.Empty
                While True
                    newSheetNumber = $"{baseSheetCode}{nextNum:D3}"
                    If Not existingSheetNumbers.Contains(newSheetNumber) Then
                        Exit While
                    End If
                    nextNum += 1
                End While
                codeCounters(baseSheetCode) = nextNum + 1
                existingSheetNumbers.Add(newSheetNumber)

                Dim newSheet As ViewSheet = ViewSheet.Create(doc, titleBlockTypeId)
                newSheet.Name = GetUniqueName(existingViewNames, newSheetTitle)
                newSheet.SheetNumber = newSheetNumber

                SetSheetParameters(newSheet, row, settings("DrawnBy"), settings("CheckedBy"), settings("DesignedBy"), settings("ApprovedBy"), settings("UnitColumn"), settings("LotNoColumn"), settings("DrawingNumbersColumn"))

                Dim noTitleViewportTypeId As ElementId = ElementId.InvalidElementId
                Dim noTitleViewportType = New FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports).OfClass(GetType(ElementType)).Cast(Of ElementType)().FirstOrDefault(Function(et) et.Name.Equals("No Title", StringComparison.OrdinalIgnoreCase))
                If noTitleViewportType IsNot Nothing Then
                    noTitleViewportTypeId = noTitleViewportType.Id
                End If

                For Each viewportId As ElementId In sampleSheet.GetAllViewports()
                    Dim viewport = TryCast(doc.GetElement(viewportId), Viewport)
                    If viewport Is Nothing Then Continue For

                    Dim viewToDuplicate = TryCast(doc.GetElement(viewport.ViewId), View)
                    If viewToDuplicate Is Nothing Then Continue For

                    Dim newViewId = viewToDuplicate.Duplicate(ViewDuplicateOption.WithDetailing)
                    Dim newView = TryCast(doc.GetElement(newViewId), View)

                    If newView IsNot Nothing Then
                        newView.Name = GetUniqueName(existingViewNames, $"{newSheet.Name} - {viewToDuplicate.Name}")
                        Dim detailItems = New FilteredElementCollector(doc, newView.Id).OfCategory(BuiltInCategory.OST_DetailComponents).OfClass(GetType(FamilyInstance)).ToElements().Cast(Of FamilyInstance)()
                        For Each detailItem In detailItems
                            SetGeometryParameters(detailItem, row)
                            SetGenericParameters(detailItem, row)
                        Next
                    End If

                    Dim newViewport = Viewport.Create(doc, newSheet.Id, newViewId, viewport.GetBoxCenter())
                    If noTitleViewportTypeId <> ElementId.InvalidElementId Then
                        newViewport.ChangeTypeId(noTitleViewportTypeId)
                    End If
                Next
                Return True
            Catch ex As Exception
                ' Log exception
                Return False
            End Try
        End Function

        Private Sub SetSheetParameters(sheet As ViewSheet, dataRow As DataRow, drawnBy As String, checkedBy As String, designedBy As String, approvedBy As String, unitCol As String, lotCol As String, dwgNumCol As String)
            SetSheetParameter(sheet, "Drawn By", drawnBy)
            SetSheetParameter(sheet, "Checked By", checkedBy)
            SetSheetParameter(sheet, "Designed By", designedBy)
            SetSheetParameter(sheet, "Approved By", approvedBy)

            SetSheetParameterFromColumn(sheet, "Unit", dataRow, unitCol)
            SetSheetParameterFromColumn(sheet, "Sub-Category", dataRow, unitCol)
            SetSheetParameterFromColumn(sheet, "Lot No.", dataRow, lotCol)
            SetSheetParameterFromColumn(sheet, "Drawing Numbers", dataRow, dwgNumCol)

            SetSheetParameter(sheet, "Drawing Type", "Gen_Fabrication")
        End Sub

        Private Sub SetSheetParameter(sheet As ViewSheet, paramName As String, value As String)
            If String.IsNullOrWhiteSpace(value) Then Return
            Dim param = sheet.LookupParameter(paramName)
            If param IsNot Nothing AndAlso Not param.IsReadOnly Then
                param.Set(value)
            End If
        End Sub

        Private Sub SetSheetParameterFromColumn(sheet As ViewSheet, paramName As String, row As DataRow, colName As String)
            If String.IsNullOrWhiteSpace(colName) OrElse Not row.Table.Columns.Contains(colName) Then Return

            Dim valueObj = row(colName)
            If valueObj IsNot DBNull.Value AndAlso valueObj IsNot Nothing Then
                Dim param = sheet.LookupParameter(paramName)
                If param IsNot Nothing AndAlso Not param.IsReadOnly Then
                    param.Set(valueObj.ToString())
                End If
            End If
        End Sub

        Private Function GetUniqueName(existingNames As HashSet(Of String), baseName As String) As String
            Dim name = baseName
            Dim counter = 1
            While existingNames.Contains(name)
                name = $"{baseName} ({counter})"
                counter += 1
            End While
            existingNames.Add(name)
            Return name
        End Function

        Private Sub SetGeometryParameters(detailItem As FamilyInstance, dataRow As DataRow)
            Dim geometryType = dataRow("Geometry Type")?.ToString()
            If String.IsNullOrEmpty(geometryType) Then Return

            Select Case geometryType
                Case "Corner-CH"
                    SetNumericParameterFromColumn(detailItem, "Arm A", dataRow, "Arm A")
                    SetNumericParameterFromColumn(detailItem, "Arm B", dataRow, "Arm B")
                    SetNumericParameterFromColumn(detailItem, "Height", dataRow, "Height")
                Case "Corner-Sec"
                    SetNumericParameterFromColumn(detailItem, "Arm A", dataRow, "Arm A")
                    SetNumericParameterFromColumn(detailItem, "Arm B", dataRow, "Arm B")
                    SetNumericParameterFromColumn(detailItem, "Height", dataRow, "Height")
                    SetNumericParameterFromColumn(detailItem, "Width", dataRow, "Width")
                Case "Rect."
                    SetNumericParameterFromColumn(detailItem, "Height", dataRow, "Height")
                    SetNumericParameterFromColumn(detailItem, "Length", dataRow, "Length")
                Case "Sec-2H"
                    SetNumericParameterFromColumn(detailItem, "Height", dataRow, "Height")
                    SetNumericParameterFromColumn(detailItem, "Length", dataRow, "Length")
                    SetNumericParameterFromColumn(detailItem, "Width", dataRow, "Width")
            End Select
        End Sub

        Private Sub SetNumericParameterFromColumn(item As FamilyInstance, paramName As String, row As DataRow, colName As String)
            If Not row.Table.Columns.Contains(colName) Then Return
            Dim valueStr = row(colName)?.ToString()
            Dim numericValue As Double
            If Double.TryParse(valueStr, numericValue) Then
                Dim param = item.LookupParameter(paramName)
                If param IsNot Nothing AndAlso Not param.IsReadOnly Then
                    param.Set(numericValue / 304.8) ' Convert mm to feet
                End If
            End If
        End Sub

        Private Sub SetGenericParameters(detailItem As FamilyInstance, dataRow As DataRow)
            For i = 1 To 50 ' Check for up to 50 P-columns
                Dim pColName = $"P{i}"
                If Not dataRow.Table.Columns.Contains(pColName) Then Continue For

                Dim cellValue = dataRow(pColName)?.ToString()
                If String.IsNullOrWhiteSpace(cellValue) OrElse Not cellValue.Contains("-") Then Continue For

                Dim parts = cellValue.Split("-"c)
                If parts.Length <> 2 Then Continue For

                Dim prefixAndIndex = parts(0).Trim().ToUpper()
                Dim valueStr = parts(1).Trim()
                Dim numericValue As Double
                If Not Double.TryParse(valueStr, numericValue) Then Continue For

                Dim firstDigitIndex = -1
                For charIndex = 0 To prefixAndIndex.Length - 1
                    If Char.IsDigit(prefixAndIndex(charIndex)) Then
                        firstDigitIndex = charIndex
                        Exit For
                    End If
                Next
                
                If firstDigitIndex <= 0 Then Continue For

                Dim letterPart = prefixAndIndex.Substring(0, firstDigitIndex)
                Dim indexPartStr = prefixAndIndex.Substring(firstDigitIndex)

                If letterPart = "P#" Then letterPart = "P"

                Dim possibleNames = New List(Of String)()
                Select Case letterPart
                    Case "B" : possibleNames.Add($"BH#{indexPartStr}")
                    Case "T" : possibleNames.Add($"TH#{indexPartStr}")
                    Case "S"
                        possibleNames.Add($"ST#{indexPartStr}")
                        possibleNames.Add($"Stiffner_{indexPartStr}")
                    Case "H"
                        possibleNames.Add($"SL#{indexPartStr}")
                        possibleNames.Add($"SH#{indexPartStr}")
                    Case "A1B" : possibleNames.Add($"A1_BH#{indexPartStr}")
                    Case "A2B" : possibleNames.Add($"A2_BH#{indexPartStr}")
                    Case "A1T" : possibleNames.Add($"A1_TH#{indexPartStr}")
                    Case "A2T" : possibleNames.Add($"A2_TH#{indexPartStr}")
                    Case "P" : possibleNames.Add($"P#{indexPartStr}")
                End Select

                If possibleNames.Any() Then
                    Dim valueInFeet = numericValue / 304.8 ' Convert mm to feet
                    For Each paramName In possibleNames
                        Dim param = detailItem.LookupParameter(paramName)
                        If param IsNot Nothing AndAlso Not param.IsReadOnly Then
                            param.Set(valueInFeet)
                            Exit For
                        End If
                    Next
                End If
            Next
        End Sub
    End Class
End Namespace 