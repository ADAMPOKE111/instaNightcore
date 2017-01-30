Imports MahApps.Metro.Controls
Imports System.Net
Imports Ionic.Zip

Class MainWindow
    Inherits MetroWindow

    ' Declare variables
    Dim droppedItems() As String
    Dim fileLoaded As Boolean
    Dim multiplier As Decimal
    Dim rate As Decimal
    Dim web As WebClient = New WebClient()
    Dim appdata As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

    Private Sub form_Load(sender As Object, e As RoutedEventArgs)
        If My.Settings.firstRun = True Then
            My.Settings.firstRun = False
            My.Settings.Save()
        End If
        If My.Settings.ffmpegDir = "" Then
            'My.Computer.FileSystem.CreateDirectory()
            btSubmit.IsEnabled = False

            If My.Computer.FileSystem.DirectoryExists(appdata & "\instaNightcore") = False Then
                My.Computer.FileSystem.CreateDirectory(appdata & "\instaNightcore")
            End If
            If My.Computer.FileSystem.FileExists(appdata & "\instaNightcore\ffmpeg.exe") = False Then
                AddHandler web.DownloadProgressChanged, AddressOf web_ProgressChanged
                AddHandler web.DownloadFileCompleted, AddressOf web_DownloadCompleted
                web.DownloadFileAsync(New Uri("https://ffmpeg.zeranoe.com/builds/win32/static/ffmpeg-3.2.2-win32-static.zip"), appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static.zip")
            End If

        End If
        If My.Computer.FileSystem.FileExists(appdata & "\instaNightcore\ffmpeg.exe") = True Then
            lbStatus.Content = "Ready to nightcore-ify!"
            slideSpeed.Value = My.Settings.ncScale
            btSubmit.IsEnabled = True
        End If
    End Sub

    Private Sub web_ProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
        Dim bytesIn As Double = Double.Parse(e.BytesReceived.ToString())
        Dim totalBytes As Double = Double.Parse(e.TotalBytesToReceive.ToString())
        Dim percentage As Double = bytesIn / totalBytes * 100
        progProgress.Value = Int32.Parse(Math.Truncate(percentage).ToString())
        lbStatus.Content = "Downloading progress: ffmpeg (" & Math.Truncate(percentage).ToString() & "%)"
    End Sub

    Private Sub web_DownloadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        lbStatus.Content = "Download finished: ffmpeg"
        'ZipFile.ExtractToDirectory(appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static.zip", appdata & "\instaNightcore")
        Using zip1 As ZipFile = ZipFile.Read(appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static.zip")
            AddHandler zip1.ExtractProgress, AddressOf MyExtractProgress
            Dim z As ZipEntry
            For Each z In zip1
                z.Extract(appdata & "\instaNightcore", ExtractExistingFileAction.OverwriteSilently)
            Next
        End Using
        My.Computer.FileSystem.MoveFile(appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static\bin\ffmpeg.exe", appdata & "\instaNightcore\ffmpeg.exe", True)
        My.Settings.ffmpegDir = appdata & "\instaNightcore"
        My.Settings.Save()
        lbStatus.Content = "Removing files: ffmpeg download"
        My.Computer.FileSystem.DeleteDirectory(appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static", FileIO.DeleteDirectoryOption.DeleteAllContents)
        My.Computer.FileSystem.DeleteFile(appdata & "\instaNightcore\ffmpeg-3.2.2-win32-static.zip")
        lbStatus.Content = "Ready to nightcore-ify!"
        btSubmit.IsEnabled = True
    End Sub

    Private Sub MyExtractProgress(sender As Object, e As ExtractProgressEventArgs)
        Return
    End Sub

    Private Sub btSettings_Open(sender As Object, e As EventArgs) Handles btSettings.Click
    End Sub

    Private Sub imgTitle_Hover(sender As Object, e As DragEventArgs) Handles imgTitle.DragEnter
        imgTitle.Source = New BitmapImage(New Uri("fire2.png", UriKind.RelativeOrAbsolute))
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effects = DragDropEffects.All
        Else
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub imgTitle_HoverLeave(sender As Object, e As DragEventArgs) Handles imgTitle.DragLeave
        imgTitle.Source = New BitmapImage(New Uri("fire.png", UriKind.RelativeOrAbsolute))
    End Sub

    Private Sub imgTitle_HoverDrop(sender As Object, e As DragEventArgs) Handles imgTitle.Drop
        droppedItems = e.Data.GetData(DataFormats.FileDrop, False)
        fileLoaded = True
        tbFile.Text = Convert.ToString(droppedItems(0))
        lbStatus.Content = "Speed multiplier: " & Convert.ToString(multiplier) & "x normal speed"
    End Sub

    Private Sub btSubmit_Click(sender As Object, e As RoutedEventArgs) Handles btSubmit.Click
        If fileLoaded = False Then
            MsgBox("Drag-and-drop a file first")
            Return
        End If

        My.Settings.ncScale = slideSpeed.Value
        My.Settings.Save()

        lbStatus.Content = "Processing..."
        Process.Start(My.Settings.ffmpegDir & "\ffmpeg.exe", " -i """ & droppedItems(0) & """ -filter:a ""asetrate=r=" & Math.Round(multiplier * 44.1).ToString & "K,atempo=" & multiplier & """ """ & Microsoft.VisualBasic.Left(droppedItems(0), droppedItems(0).Length - 3) & """.wav")
        lbStatus.Content = "Done! File saved to: " & Microsoft.VisualBasic.Left(droppedItems(0), droppedItems(0).Length - 3) & ".wav"
        imgTitle.Source = New BitmapImage(New Uri("fire.png", UriKind.RelativeOrAbsolute))
    End Sub

    Private Sub slideSpeed_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles slideSpeed.ValueChanged
        multiplier = Math.Round((100 + slideSpeed.Value) / 100, 1)
        If fileLoaded = True Then
            lbStatus.Content = "Speed multiplier: " & Convert.ToString(multiplier) & "x normal speed"
        End If
    End Sub
End Class
