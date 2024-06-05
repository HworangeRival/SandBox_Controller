Public Class Form1
    ReadOnly rnd As New Random, SleepMessage As Integer = 700
    Dim ThrExec As Threading.Thread = Nothing
    Private Sub BlueButtonB2_Click(sender As Object, e As EventArgs) Handles BlueButtonB2.Click
        Environment.Exit(0)
    End Sub
    Private Sub BlueButtonC1_Click(sender As Object, e As EventArgs) Handles BlueButtonC1.Click
        If ThrExec IsNot Nothing Then ThrExec.Abort() : ThrExec = Nothing
        ThrExec = New Threading.Thread(AddressOf LoadSandBox)
        ThrExec.TrySetApartmentState(Threading.ApartmentState.STA) : ThrExec.Start()
    End Sub
    Private Function GenerateSandBoxConfigFile() As String
        Dim Src_Config As String = SandBoxConfig.Text '
        Src_Config = Src_Config.Replace("§AUDIO§", CheckBox8.Checked)
        Src_Config = Src_Config.Replace("§VIDEO§", CheckBox9.Checked)
        Src_Config = Src_Config.Replace("§VGPU§", CheckBox1.Checked)
        Src_Config = Src_Config.Replace("§NETWORK§", CheckBox3.Checked)
        Src_Config = Src_Config.Replace("§CLIENT§", CheckBox5.Checked)
        Src_Config = Src_Config.Replace("§PRINTER§", CheckBox6.Checked)
        Src_Config = Src_Config.Replace("§CLIPBOARD§", CheckBox7.Checked)
        Src_Config = Src_Config.Replace("§MEMORY§", NumericUpDown1.Value)
        Src_Config = Src_Config.Replace("§HOSTFOLDERPATH§", TextBox1.Text)
        Src_Config = Src_Config.Replace("§SANDBOXFOLDERPATH§", "C:\Users\WDAGUtilityAccount\" & ComboBox1.Text)
        Src_Config = Src_Config.Replace("§READONLYSBFOLDER§", CheckBox4.Checked)
        Src_Config = Src_Config.Replace("§COMMANDTOEXECUTE§", TextBox2.Text)
        GenerateSandBoxConfigFile = Src_Config
    End Function

    Private Sub LoadSandBox()
        BlueLabel5.Text = "Status: Checking system.."
        Dim Sb As String = "%windir%\System32\WindowsSandbox.exe"
        If IO.File.Exists(Sb) Then
            Temppath = Temppath.Replace("Roaming", "Local") & "\SandBoxConfigurator"
            If IO.Directory.Exists(Temppath) Then IO.Directory.Delete(Temppath, True)
            IO.Directory.CreateDirectory(Temppath)
            Temppath = Temppath & "\Sand_" & rnd.Next(111, 999) & ".wsb"
            IO.File.WriteAllText(Temppath, GenerateSandBoxConfigFile)
            BlueLabel5.Text = "Status: WindowsSandbox Loading..."
            Dim SandBox As New ProcessStartInfo With {.FileName = Sb, .Arguments = Temppath}
            Dim NewProcess As New Process With {.StartInfo = SandBox}
            NewProcess.Start()
            NewProcess.WaitForExit()
            SetStatusBottomLabel("WindowsSandbox Closed!")
            IO.Directory.Delete(Temppath, True)
        Else
            SetStatusBottomLabel("Load Failed - WindowsSandbox is not installed!")
            TryToInstallSB()
        End If
    End Sub
    Dim Temppath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

    Private Sub TryToInstallSB()
        Try
            Dim W_Id = Security.Principal.WindowsIdentity.GetCurrent()
            Dim WP = New Security.Principal.WindowsPrincipal(W_Id)
            Dim isAdmin As Boolean = WP.IsInRole(Security.Principal.WindowsBuiltInRole.Administrator)
            If isAdmin Then
                Temppath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).Replace("Roaming", "Local") & "\SandBoxConfigurator"
                If IO.Directory.Exists(Temppath) = False Then IO.Directory.CreateDirectory(Temppath)
                Dim Installer_SB = Temppath & "\InstallerSandBox" & ".ps1"
                IO.File.WriteAllText(Installer_SB, Installer_SandBox_ps1.Text)
                SetStatusBottomLabel("Trying install WindowsSandbox...")
                Dim psi As New ProcessStartInfo()
                psi.Verb = "runas" : psi.FileName = "cmd.exe"
                psi.Arguments = "/c " & "powershell.exe -ExecutionPolicy ByPass -File '" & Installer_SB & "'" ' <- pass arguments for the command you want to run
                Process.Start(psi).WaitForExit()
                '                Process.Start("cmd.exe", "powershell.exe -ExecutionPolicy ByPass -File '" & Installer_SB & "'").WaitForExit()
                IO.File.Delete(Installer_SB)
                SetStatusBottomLabel("Work completed!")
            Else
                MsgBox("Run the software with Admin Right for download/install SandBox!", MsgBoxStyle.Information, Me.Text)
                SetStatusBottomLabel("You not have the right for installation!")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Control.CheckForIllegalCrossThreadCalls = False
        TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        ComboBox1.SelectedIndex += 1
        BlueLabel5.Text = "Status: -Idle"
    End Sub
    Private Sub SaveConfigFile()
        BlueLabel5.Text = "Status: Saving configuration file (.wsb).."
        Dim sfd As New SaveFileDialog With {.Title = "Save your configuration File!", .FileName = "Config_" & rnd.Next(11, 99), .Filter = "Windows Sandbox File|*.wsb",
             .RestoreDirectory = True, .SupportMultiDottedExtensions = False}
        If sfd.ShowDialog = DialogResult.OK Then
            IO.File.WriteAllText(sfd.FileName, GenerateSandBoxConfigFile)
            SetStatusBottomLabel("Configuration file saved! (.wsb)")
        Else
            SetStatusBottomLabel("-Idle")
        End If
    End Sub
    Private Sub BlueButtonB1_Click(sender As Object, e As EventArgs) Handles BlueButtonB1.Click
        If ThrExec IsNot Nothing Then ThrExec.Abort() : ThrExec = Nothing
        ThrExec = New Threading.Thread(AddressOf SaveConfigFile)
        ThrExec.TrySetApartmentState(Threading.ApartmentState.STA) : ThrExec.Start()
    End Sub

    Private Sub BlueButtonB3_Click(sender As Object, e As EventArgs) Handles BlueButtonB3.Click
        Dim fbd As New FolderBrowserDialog With {.ShowNewFolderButton = False, .Description = "Select Local Sahred Folder!"}
        If fbd.ShowDialog = DialogResult.OK Then TextBox1.Text = fbd.SelectedPath
    End Sub

    Private Sub SetStatusBottomLabel(text As String)
        BlueLabel5.Text = "Status: " & text : Threading.Thread.Sleep(SleepMessage) : BlueLabel5.Text = "Status: -Idle"
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        MsgBox("Powered by ..:: Hworange Rival ::..", MsgBoxStyle.Information, Me.Text)
    End Sub
End Class
