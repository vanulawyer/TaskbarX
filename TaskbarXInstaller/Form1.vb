Imports System.Environment
Imports System.IO
Imports Microsoft.Win32.TaskScheduler

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        install()
    End Sub

    Public Sub install()
        Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim programfiles = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim path As String = programfiles & "\TaskbarX"
        Button1.Enabled = False

        ProgressBar1.Value = 0
        Label3.Text = "Stopping TaskbarX..."
        Shell("taskkill /F /IM TaskbarX.exe")
        Shell("taskkill /F /IM " & Chr(34) & "TaskbarX Configurator.exe" & Chr(34))
        ProgressBar1.Value = ProgressBar1.Value + 10
        System.Threading.Thread.Sleep(20) : Application.DoEvents()

        Try
            If Directory.Exists(path) Then
                System.IO.Directory.Delete(path, True)
            End If
        Catch
        End Try

        Directory.CreateDirectory(path)

        For Each Filex In Directory.GetFiles(Application.StartupPath)
            Dim filename = Filex.Replace(Application.StartupPath, "")
            Try
                File.Copy(Filex, path & "\" & filename)
            Catch
            End Try
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label3.Text = "Installing " & filename & "..."
            System.Threading.Thread.Sleep(10) : Application.DoEvents()
        Next

        Try
            CreateShortCut(path & "\TaskbarX.exe", FileIO.SpecialDirectories.Desktop, "TaskbarX")
            CreateShortCut(path & "\TaskbarX Configurator.exe", FileIO.SpecialDirectories.Desktop, "TaskbarX Configurator")
        Catch
        End Try

        Try
            If Directory.Exists(appData & "\Microsoft\Windows\Start Menu\Programs\Chris Andriessen") Then
                System.IO.Directory.Delete(appData & "\Microsoft\Windows\Start Menu\Programs\Chris Andriessen", True)
            End If
        Catch
        End Try

        Try
            Directory.CreateDirectory(appData & "\Microsoft\Windows\Start Menu\Programs\TaskbarX")
            CreateShortCut(path & "\TaskbarX.exe", appData & "\Microsoft\Windows\Start Menu\Programs\TaskbarX", "TaskbarX")
            CreateShortCut(path & "\TaskbarX Configurator.exe", appData & "\Microsoft\Windows\Start Menu\Programs\TaskbarX", "TaskbarX Configurator")
        Catch
        End Try
        '  Console.WriteLine(appData)

        Dim FileVer As String = FileVersionInfo.GetVersionInfo(path & "\TaskbarX.exe").FileVersion

        Console.WriteLine(FileVer)

        'Setting My Values
        Dim ApplicationName As String = "TaskbarX"
        Dim ApplicationVersion As String = FileVer
        Dim ApplicationIconPath As String = "%APPDATA%\TaskbarX\TaskbarX-Install-100.ico"
        Dim ApplicationPublisher As String = "Chris Andriessen"
        Dim ApplicationUnInstallPath As String = "%APPDATA%\TaskbarX\TaskbarXInstaller.exe"
        Dim ApplicationInstallDirectory As String = "%APPDATA%\TaskbarX\"

        ProgressBar1.Value = ProgressBar1.Value + 10
        Label3.Text = "Creating uninstall entry..."
        System.Threading.Thread.Sleep(10) : Application.DoEvents()

        Try
            'Openeing the Uninstall RegistryKey (don't forget to set the writable flag to true)
            With My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall", True)

                'Creating my AppRegistryKey
                Dim AppKey As Microsoft.Win32.RegistryKey = .CreateSubKey(ApplicationName)

                'Adding my values to my AppRegistryKey
                AppKey.SetValue("DisplayName", ApplicationName, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("DisplayVersion", ApplicationVersion, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("DisplayIcon", ApplicationIconPath, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("Publisher", ApplicationPublisher, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("UninstallString", ApplicationUnInstallPath, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("UninstallPath", ApplicationUnInstallPath, Microsoft.Win32.RegistryValueKind.String)
                AppKey.SetValue("InstallLocation", ApplicationInstallDirectory, Microsoft.Win32.RegistryValueKind.String)

            End With
        Catch
        End Try

        ProgressBar1.Value = 100
        Label3.Text = "Done!"

        Process.Start(path & "\TaskbarX.exe")
        Process.Start(path & "\TaskbarX Configurator.exe")

        Button1.Enabled = True
        Button2.Enabled = True
        Button1.Text = "Update"
    End Sub

    Private Function CreateShortCut(ByVal TargetName As String, ByVal ShortCutPath As String, ByVal ShortCutName As String) As Boolean
        Dim oShell As Object
        Dim oLink As Object
        'you don’t need to import anything in the project reference to create the Shell Object

        Try

            oShell = CreateObject("WScript.Shell")
            oLink = oShell.CreateShortcut(ShortCutPath & "\" & ShortCutName & ".lnk")

            oLink.TargetPath = TargetName
            oLink.WindowStyle = 1
            oLink.Save()
        Catch ex As Exception

        End Try

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ProgressBar1.Value = 0

        Label3.Text = "Stopping TaskbarX..."
        System.Threading.Thread.Sleep(20) : Application.DoEvents()

        Dim programfiles = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim path As String = programfiles & "\TaskbarX"

        Process.Start(path & "\TaskbarX.exe", "-stop")

        System.Threading.Thread.Sleep(1000) : Application.DoEvents()
        Shell("taskkill /F /IM TaskbarX.exe")
        Shell("taskkill /F /IM " & Chr(34) & "TaskbarX Configurator.exe" & Chr(34))
        ProgressBar1.Value = ProgressBar1.Value + 20
        System.Threading.Thread.Sleep(20) : Application.DoEvents()
        System.Threading.Thread.Sleep(1000) : Application.DoEvents()

        Label3.Text = "Removing Files..."
        Try
            If Directory.Exists(path) Then
                System.IO.Directory.Delete(path, True)
            End If
        Catch
        End Try

        Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        ProgressBar1.Value = ProgressBar1.Value + 20
        Label3.Text = "Removing Shortcuts..."
        System.Threading.Thread.Sleep(20) : Application.DoEvents()

        Try
            If Directory.Exists(appData & "\Microsoft\Windows\Start Menu\Programs\Chris Andriessen") Then
                System.IO.Directory.Delete(appData & "\Microsoft\Windows\Start Menu\Programs\Chris Andriessen", True)
            End If
            If Directory.Exists(appData & "\Microsoft\Windows\Start Menu\Programs\TaskbarX") Then
                System.IO.Directory.Delete(appData & "\Microsoft\Windows\Start Menu\Programs\TaskbarX", True)
            End If
        Catch
        End Try

        ProgressBar1.Value = ProgressBar1.Value + 20
        Label3.Text = "Removing Taskschedule..."
        System.Threading.Thread.Sleep(20) : Application.DoEvents()
        Try
            Using ts As TaskService = New TaskService()
                ts.RootFolder.DeleteTask("TaskbarX")
                ts.RootFolder.DeleteTask("FalconX")
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        ProgressBar1.Value = ProgressBar1.Value + 20
        ' Label3.Text = "Removing Taskschedule..."
        System.Threading.Thread.Sleep(20) : Application.DoEvents()

        Try
            With My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall", True)
                .DeleteSubKey("TaskbarX")
            End With
        Catch
        End Try

        ProgressBar1.Value = 100
        System.Threading.Thread.Sleep(20) : Application.DoEvents()
        Label3.Text = "Uninstalled!"

        Button1.Enabled = True
        Button2.Enabled = False
        Button1.Text = "Install"

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim newFileVer As String = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\TaskbarX.exe").FileVersion
        Label5.Text = "Version: " & newFileVer

        Dim programfiles = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim path As String = programfiles & "\TaskbarX\TaskbarX.exe"

        Try
            If File.Exists(path) Then
                Dim curFileVer As String = FileVersionInfo.GetVersionInfo(path).FileVersion
                Button1.Text = "Update"
                Label3.Text = "TaskbarX " & curFileVer & " is currently installed."
            Else
                Button1.Enabled = True
                Button2.Enabled = False
                Label3.Text = "TaskbarX is currently not installed."
            End If
        Catch
        End Try

    End Sub

End Class