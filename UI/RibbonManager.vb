Imports Autodesk.Revit.UI
Imports System.Reflection
Imports System.IO
Imports System.Windows.Media.Imaging

Namespace Autoform.UI
    Public Class RibbonManager
        Public Sub Initialize(ByVal app As UIControlledApplication)
            Dim tabName As String = "Autoform"
            Try
                app.CreateRibbonTab(tabName)
            Catch ex As Exception
                ' Tab already exists
            End Try


            Dim panelName As String = "Auto Fabrications"
            Dim panel As RibbonPanel = app.CreateRibbonPanel(tabName, panelName)

            Dim assemblyPath As String = Assembly.GetExecutingAssembly().Location
            Dim commandPath As String = "Autoform.Commands.NStandard.NStandardCommand"

            Dim pushButtonData As New PushButtonData("NSTD_Button", "N Standard", assemblyPath, commandPath)

            Dim iconPath As String = Path.GetDirectoryName(assemblyPath)
            pushButtonData.LargeImage = New BitmapImage(New Uri(Path.Combine(iconPath, "Resources", "Icons", "N_STD_32.bmp")))
            pushButtonData.Image = New BitmapImage(New Uri(Path.Combine(iconPath, "Resources", "Icons", "N_STD_16.bmp")))

            panel.AddItem(pushButtonData)
        End Sub
    End Class
End Namespace 