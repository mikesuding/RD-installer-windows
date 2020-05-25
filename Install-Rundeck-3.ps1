# Use TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Pre defined variables
$Script:JavaDownloadURL = 'https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html'
$Script:RundeckDownloadURL = 'https://rundeckpro.bintray.com/war/rundeckpro-enterprise-3.2.4-20200318.war'
$Script:NssmDownloadURL = 'http://nssm.cc/release/nssm-2.24.zip'
$Script:SqlDjbcDll = 'http://wowie.us/sqljdbc_auth.dll'   # Mike Suding website for convenient download
$Script:InstallTimeOut = 240
$Script:DefaultRundeckDirectory = "C:\Rundeck"
$Script:UsedRundeckDirectory = ""

#Get new URL of war file on ENTERPRISE download page
$AllURLs = "https://download.rundeck.com/eval/"
$Response = Invoke-WebRequest -Uri $AllURLs -UseBasicParsing
if ($Response.Content -match "https://.*rundeckpro-ent.*war")
{
    $Script:RundeckDownloadURL = $Matches[0]
}
else
{
    Write-Warning "WARNING: Not able to find URL of Rundeck"
}

# Add forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore,PresentationFramework

Function Show-MainForm
{
    $Form = New-Object System.Windows.Forms.Form
    $MenuStrip = New-Object System.Windows.Forms.MenuStrip
    $StripMenuItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
    $StripMenuItem2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $StripMenuItem3 = New-Object System.Windows.Forms.ToolStripMenuItem
    $StripMenuItem4 = New-Object System.Windows.Forms.ToolStripMenuItem
    $StripMenuItem5 = New-Object System.Windows.Forms.ToolStripMenuItem
    $GroupBox1 = New-Object System.Windows.Forms.GroupBox
    $GroupBox2 = New-Object System.Windows.Forms.GroupBox
    $GroupBox3 = New-Object System.Windows.Forms.GroupBox
    $GroupBox4 = New-Object System.Windows.Forms.GroupBox
    $GroupBox5 = New-Object System.Windows.Forms.GroupBox
    $Button1 = New-Object System.Windows.Forms.Button
    $Button2 = New-Object System.Windows.Forms.Button
    $Button3 = New-Object System.Windows.Forms.Button
    $Button4 = New-Object System.Windows.Forms.Button
    $Button5 = New-Object System.Windows.Forms.Button
    $Button6 = New-Object System.Windows.Forms.Button
    $Button7 = New-Object System.Windows.Forms.Button
    $Button8 = New-Object System.Windows.Forms.Button
    $Button9 = New-Object System.Windows.Forms.Button
    $Button10 = New-Object System.Windows.Forms.Button
    $Button11 = New-Object System.Windows.Forms.Button
    $Button12 = New-Object System.Windows.Forms.Button
    $Button13 = New-Object System.Windows.Forms.Button
    $LinkLabel1 = New-Object System.Windows.Forms.LinkLabel
    $Label1 = New-Object System.Windows.Forms.Label
    $Label2 = New-Object System.Windows.Forms.Label
    $Label3 = New-Object System.Windows.Forms.Label
    $Label4 = New-Object System.Windows.Forms.Label
    $Label5 = New-Object System.Windows.Forms.Label
    $Label6 = New-Object System.Windows.Forms.Label
    $Label7 = New-Object System.Windows.Forms.Label
    $Label8 = New-Object System.Windows.Forms.Label
    $Label9 = New-Object System.Windows.Forms.Label
    $Label10 = New-Object System.Windows.Forms.Label
    $Label11 = New-Object System.Windows.Forms.Label
    $Label12 = New-Object System.Windows.Forms.Label
    $Label13 = New-Object System.Windows.Forms.Label
    $Label14 = New-Object System.Windows.Forms.Label
    $Label15 = New-Object System.Windows.Forms.Label
    $Label16 = New-Object System.Windows.Forms.Label
    $ComboBox1 = New-Object System.Windows.Forms.ComboBox
    $ComboBox2 = New-Object System.Windows.Forms.ComboBox
    $TextBox1 = New-Object System.Windows.Forms.TextBox
    $TextBox2 = New-Object System.Windows.Forms.TextBox
    $TextBox3 = New-Object System.Windows.Forms.TextBox
    $TextBox4 = New-Object System.Windows.Forms.TextBox
    $TextBox5 = New-Object System.Windows.Forms.TextBox
    $TextBox6 = New-Object System.Windows.Forms.TextBox
    $TextBox7 = New-Object System.Windows.Forms.TextBox
    $TextBox8 = New-Object System.Windows.Forms.TextBox
    $TextBox9 = New-Object System.Windows.Forms.TextBox
    $TextBox10 = New-Object System.Windows.Forms.TextBox
    $TextBox11 = New-Object System.Windows.Forms.TextBox
    $TextBox12 = New-Object System.Windows.Forms.TextBox
    $TextBox13 = New-Object System.Windows.Forms.TextBox
    $TextBox14 = New-Object System.Windows.Forms.TextBox
    $TextBox15 = New-Object System.Windows.Forms.TextBox
    $TextBox16 = New-Object System.Windows.Forms.TextBox

    $StripMenuItem1.Text = "1. Java"
    $StripMenuItem1.add_Click({
        $GroupBox1.Visible = $true
        $GroupBox2.Visible = $false
        $GroupBox3.Visible = $false
        $GroupBox4.Visible = $false
        $GroupBox5.Visible = $false
    })
    $MenuStrip.Items.Add($StripMenuItem1)

    $StripMenuItem2.Text = "2. Rundeck"
    $StripMenuItem2.add_Click({
        $GroupBox1.Visible = $false
        $GroupBox2.Visible = $true
        $GroupBox3.Visible = $false
        $GroupBox4.Visible = $false
        $GroupBox5.Visible = $false
    })
    $MenuStrip.Items.Add($StripMenuItem2)

    $StripMenuItem3.Text = "3. Database"
    $StripMenuItem3.add_Click({
        $GroupBox1.Visible = $false
        $GroupBox2.Visible = $false
        $GroupBox3.Visible = $true
        $GroupBox4.Visible = $false
        $GroupBox5.Visible = $false
    })
    $MenuStrip.Items.Add($StripMenuItem3)

    $StripMenuItem4.Text = "4. Authentication"
    $StripMenuItem4.add_Click({
        $GroupBox1.Visible = $false
        $GroupBox2.Visible = $false
        $GroupBox3.Visible = $false
        $GroupBox4.Visible = $true
        $GroupBox5.Visible = $false
    })
    $MenuStrip.Items.Add($StripMenuItem4)

    $StripMenuItem5.Text = "5. WinRM"
    $StripMenuItem5.add_Click({
        $GroupBox1.Visible = $false
        $GroupBox2.Visible = $false
        $GroupBox3.Visible = $false
        $GroupBox4.Visible = $false
        $GroupBox5.Visible = $true
    })
    $MenuStrip.Items.Add($StripMenuItem5)

    $GroupBox1.Size = "740,300"
    $GroupBox1.Location = "25,80"
    $GroupBox1.Text = "Java Installation"
    $GroupBox1.Visible = $false

    $GroupBox2.Size = "740,300"
    $GroupBox2.Location = "25,80"
    $GroupBox2.Text = "Rundeck Installation"
    $GroupBox2.Visible = $false

    $GroupBox3.Size = "740,300"
    $GroupBox3.Location = "25,80"
    $GroupBox3.Text = "Database configuration"
    $GroupBox3.Visible = $false

    $GroupBox4.Size = "740,300"
    $GroupBox4.Location = "25,80"
    $GroupBox4.Text = "Authentication Configuration"
    $GroupBox4.Visible = $false

    $GroupBox5.Size = "740,300"
    $GroupBox5.Location = "25,80"
    $GroupBox5.Text = "WinRM"
    $GroupBox5.Visible = $false

    $Button1.Text = "Download Java"
    $Button1.Location = "15,40"
    $Button1.Size = "100,25"
    $Button1.add_Click({
        # Just Open URL
        Start-Process $Script:JavaDownloadURL
        $Label2.Text = "Download page opened."
    })
    $GroupBox1.Controls.Add($Button1)

    $Button2.Text = "Install Java"
    $Button2.Location = "15,120"
    $Button2.Size = "100,25"
    $Button2.add_Click({
        # Open file dialog to select installator of JDK
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.filter = "Exe Files (.exe)|*.exe"

        if ($OpenFileDialog.ShowDialog() -eq "OK")
        {
            #Open File Dialog Ok
            $Context = @{}
            $Context.JavaInstallerPath = $OpenFileDialog.FileName
            $Context.TimeOutSeconds = $Script:InstallTimeOut

            # Create Async script block
            $BackgroundScript = [Powershell]::Create().AddScript({
                Param ($Context)
                $ExitValue = @{}

                $JavaLogPath = $env:TEMP + "\JavaInstallationLog_" + $(Get-Date -Format "ddMMyyyy_HHmmss" ) +".log"

                & $Context.JavaInstallerPath /s /L $JavaLogPath

                $IsJavaInstalled = $false
                $IsJavaPathAdded = $false
                $TimeDiffSeconds = 0

                # Silent Java instalation started, read log file until timeout
                # If instalation sucesful, leave the loop, start checking PATH enviroment variable
                while ($TimeDiffSeconds -lt $Context.TimeOutSeconds)
                {
                    Start-Sleep -Seconds 1
                    if ((Test-Path -Path $JavaLogPath) -and (Select-String -Path $JavaLogPath -Pattern "Windows Installer installed the product."))
                    {
                        $IsJavaInstalled = $true
                        break
                    }
                    $TimeDiffSeconds = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                    $EndDate = Get-Date
                }
                if ($IsJavaInstalled)
                {
                    $StartDate = Get-Date
                    $EndDate = Get-Date
                    $TimeDiffSeconds = 0

                    #Start searching for PATH update, check javapath, if discovered - update session PATH var
                    while ($TimeDiffSeconds -lt $Context.TimeOutSeconds)
                    {
                        Start-Sleep -Seconds 1
                        if ([System.Environment]::GetEnvironmentVariable("Path","Machine").Contains("javapath"))
                        {
                            $IsJavaPathAdded = $true
                            break
                        }
                        $TimeDiffSeconds = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                        $EndDate = Get-Date
                    }
                    if ($IsJavaPathAdded)
                    {
                        $ExitValue.Text = "Java Installed, PATH set."
                        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
                    }
                    else {
                        $ExitValue.Text = "Java Installed, PATH not set."
                    }
                }
                else
                {
                    $ExitValue.Text = "Java not installed, check log file."
                }
                return $ExitValue
            })
            
            # Run async script block
            $Runspace = [runspacefactory]::CreateRunspace()
            $Runspace.ApartmentState = "STA"
            $Runspace.ThreadOptions = "ReuseThread"
            $Runspace.Open()

            $BackgroundScript.AddArgument($Context)
            $BackgroundScript.Runspace = $Runspace
            $BG_Job = $BackgroundScript.BeginInvoke()

            $StartDate = Get-Date
            $EndDate = Get-Date
            $Label1.ForeColor = "Red"
            $Label2.Text = "Installing Java"

            while (-not $BG_Job.IsCompleted)
            {
                Start-Sleep -Milliseconds 500
                $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                $Label1.Visible = -not $Label1.Visible
                $Label1.Text = "Busy: $ExecutionTime s"
                $EndDate = Get-Date
            }

            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.ForeColor = "Black"
            $Label1.Visible = $true
            $Label1.Text = "Done in $ExecutionTime s"

            $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
            $BackgroundScript.Dispose()
            $Runspace.Close()

            $Label2.Text = $BgJobResult.Text
        }
        else
        {
            #Open File Dialog canceled, async script block not executed
            $Label2.Text = "Open file dialog cancelled."
        }
    })
    $GroupBox1.Controls.Add($Button2)

    $Button3.Text = "Test Java"
    $Button3.Location = "15, 200"
    $Button3.Size = "100,25"
    $Button3.add_Click({

        $BackgroundScript = [Powershell]::Create().AddScript({

            $ExitValue = @{}

            # Run java version check
            $JavaVersionOut = & cmd /c "java -version 2>&1"
            if (Select-String -InputObject $JavaVersionOut[0] -Pattern "version")
            {
                $ExitValue.Result = "OK: " + $JavaVersionOut[0]
            }
            else
            {
                $ExitValue.Result = "Err: " + $JavaVersionOut[0]
            }

            $ExitValue.Text = "Java check complete."

            return $ExitValue
        })
        
        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Checking Java"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()

        $Label2.Text = $BgJobResult.Text
        $TextBox1.Text = $BgJobResult.Result
    })
    $GroupBox1.Controls.Add($Button3)

    $Button4.Text = "Download, Install and start the Rundeck service"
    $Button4.Location = "35,160"
    $Button4.Size = "255,25"
    $Button4.add_Click({
        # Remove old folder, and create new
        if (Test-Path $Script:DefaultRundeckDirectory)
        {
            Remove-Item -Path $Script:DefaultRundeckDirectory -Recurse
        }
        New-Item -Path $Script:DefaultRundeckDirectory -ItemType Directory

        # Select folder to install rundeck
        $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $SelectFolderDialog.Description = "Select Rundeck Folder"
        $SelectFolderDialog.SelectedPath = $Script:DefaultRundeckDirectory

        if ($SelectFolderDialog.ShowDialog() -eq "OK")
        {
            # Update script variable
            $Script:UsedRundeckDirectory = $SelectFolderDialog.SelectedPath

            $Context = @{}
            $Context.RundeckFolder = $SelectFolderDialog.SelectedPath
            $Context.RundeckURL = $Script:RundeckDownloadURL
            $Context.NssmURL = $Script:NssmDownloadURL
            $Context.TimeOutSeconds = $Script:InstallTimeOut

            # Use Rundeck RunAs account
            $Context.RundeckRunAs = $false
            if ($TextBox7.Text -and $TextBox8.Text)
            {
                $Context.RundeckRunAs = $true
                $Context.RundeckRunAsUser = $TextBox7.Text
                $Context.RundeckRunAsPass = $TextBox8.Text
            }

            # Test local Admin rights
            if ($Context.RundeckRunAsUser)
            {
                $FullName = $Context.RundeckRunAsUser.Replace(".\","$($env:computername)\")
                $Admins = (Get-LocalGroupMember -Group "Administrators").Name
                if (-not ($FullName -in $Admins))
                {
                    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                    [System.Windows.MessageBox]::Show("User $FullName not in Administartors group",$MessageIcon)
                }
            }
            $BackgroundScript = [Powershell]::Create().AddScript({
                Param ($Context)
                $ExitValue = @{}

                # Start Download of Rundeck War
                Set-Location $Context.RundeckFolder
                Start-BitsTransfer -Source $Context.RundeckURL
                $RundeckFleName = $Context.RundeckURL.Split("/")[-1]

                # First run - perform installation
                $JavaScriptBlock = {
                    Set-Location $args[0]
                    set RDECK_BASE=$args[0]
                    & java -jar $args[1] >console.out 2>console.err
                }
                $JavaJob = Start-Job -ScriptBlock $JavaScriptBlock -ArgumentList ($Context.RundeckFolder, $RundeckFleName)

                $StartDate = Get-Date
                $EndDate = Get-Date
                $TimeDiffSeconds = 0
                $IsRundeckInstalled = $false

                # Java Executed - searching in rundeck log
                while ($TimeDiffSeconds -lt $Context.TimeOutSeconds)
                {
                    Start-Sleep -Seconds 1
                    if ((Test-Path "console.out") -and (Select-String -Path console.out -Pattern "Rundeck startup finished in"))
                    {
                        Stop-Job $JavaJob
                        $IsRundeckInstalled = $true
                        break
                    }
                    $TimeDiffSeconds = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                    $EndDate = Get-Date
                }

                # Rundeck installed, Downloading nssm
                if ($IsRundeckInstalled)
                {
                    $RundeckStartFile = 'set CURDIR=%~dp0
call %CURDIR%etc\profile.bat
java %RDECK_CLI_OPTS% %RDECK_SSL_OPTS% -jar ' + $RundeckFleName + ' --skipinstall -d  >> %CURDIR%\var\logs\service.log  2>&1'
                    
                    $RundeckStartFile | Out-File "$($Context.RundeckFolder)\start_rundeck.bat" -Encoding ascii
                    Start-BitsTransfer -Source $Context.NssmURL -Destination "$($Context.RundeckFolder)\nssm.zip"
                    Expand-Archive -Path "$($Context.RundeckFolder)\nssm.zip" -DestinationPath "$($Context.RundeckFolder)\nssm\"
                    Copy-Item -Path "$($Context.RundeckFolder)\nssm\*\win64\nssm.exe" -Destination "$($Context.RundeckFolder)\nssm.exe"
                    & "$($Context.RundeckFolder)\nssm.exe" install RUNDECK "$($Context.RundeckFolder)\start_rundeck.bat"

                    if ($Context.RundeckRunAs)
                    {
                        Start-Sleep -Seconds 5
                        Start-Process "$($Context.RundeckFolder)\nssm.exe" -ArgumentList "set RUNDECK ObjectName $($Context.RundeckRunAsUser) $($Context.RundeckRunAsPass)"
                    }

                    $CliOptions = Get-Content -Path "$($Context.RundeckFolder)\etc\profile.bat"
                    $CliModifiedOptions = $CliOptions -replace "-Xms64m -Xmx128m", "-Xms1024m -Xmx2048m"
                    $CliModifiedOptions | Set-Content -Path "$($Context.RundeckFolder)\etc\profile.bat"

                    # Statr new Service
                    Start-Service RUNDECK

                    $ExitValue.Text = "Please wait up to 60 seconds for Rundeck to start listening on port."
                }
                else {
                    $ExitValue.Text = "Rundeck installation failed"
                }

                return $ExitValue
            })

            # Run async script block
            $Runspace = [runspacefactory]::CreateRunspace()
            $Runspace.ApartmentState = "STA"
            $Runspace.ThreadOptions = "ReuseThread"
            $Runspace.Open()

            $BackgroundScript.AddArgument($Context)
            $BackgroundScript.Runspace = $Runspace
            $BG_Job = $BackgroundScript.BeginInvoke()

            $StartDate = Get-Date
            $EndDate = Get-Date
            $Label1.ForeColor = "Red"
            $Label2.Text = "Installing Rundeck"

            while (-not $BG_Job.IsCompleted)
            {
                Start-Sleep -Milliseconds 500
                $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                $Label1.Visible = -not $Label1.Visible
                $Label1.Text = "Busy: $ExecutionTime s"
                $EndDate = Get-Date
            }

            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.ForeColor = "Black"
            $Label1.Visible = $true
            $Label1.Text = "Done in $ExecutionTime s"

            $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
            $BackgroundScript.Dispose()
            $Runspace.Close()

            $Label2.Text = $BgJobResult.Text
        }
    })
    $GroupBox2.Controls.Add($Button4)

    $Button5.Text = "Test Rundeck"
    $Button5.Location = "35,240"
    $Button5.Size = "100,25"
    $Button5.add_Click({

        $BackgroundScript = [Powershell]::Create().AddScript({
            $ExitValue = @{}
            $targetHost = hostname
            $URL = "http://" + $targetHost + ":4440"
            try {
                $response = Invoke-WebRequest -Uri $URL
                $ExitValue.Result = "OK - response code $($response.StatusCode) on $URL"
                $ExitValue.RundeckURL = $URL
            }
            catch {
                $errorVar = $_.Exception.Message
                $ExitValue.Result = "Error - " + $errorVar
            }

            $ExitValue.Text = "Rundeck tested"

            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Testing Rundeck"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()

        if ($BgJobResult.RundeckURL) {
            $LinkLabel1.Text = $BgJobResult.RundeckURL
        }
        $Label2.Text = $BgJobResult.Text
        $TextBox2.Text = $BgJobResult.Result
    })
    $GroupBox2.Controls.Add($Button5)

    $Button6.Text = "Create a 'Rundeck' database with permissions"
    $Button6.Location = "15,130"
    $Button6.Size = "250,25"
    $Button6.Visible = $false
    $Button6.add_Click({

        $Context = @{}
        $Context.SQLServer = $TextBox3.Text
        $Context.SelectedAuth = $ComboBox1.SelectedItem
        $Context.Username = $TextBox4.Text
        $Context.Password = $TextBox5.Text

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}
            
            # Prepare SQL statements
            $CreateDBStatement = "CREATE DATABASE [Rundeck]"
            if ($Context.SelectedAuth -eq "MS-SQL with Windows authentication")
            {
                $CreateLoginStatement = "CREATE LOGIN [$($Context.Username)] FROM WINDOWS"
            }
            elseif ($Context.SelectedAuth -eq "MS-SQL with SQL authentication")
            {
                $CreateLoginStatement = "CREATE LOGIN [$($Context.Username)] WITH PASSWORD=N'$($Context.Password)', CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF"
            }
            $CreateDBUser = "USE [Rundeck]`nGO`nCREATE USER [RundeckDBUser] FOR LOGIN [$($Context.Username)]`n ALTER ROLE [db_owner] ADD MEMBER [RundeckDBUser]"

            # Execute SQL statements
            try
            {
                Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CreateDBStatement -ErrorAction 'Stop'
                Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CreateLoginStatement -ErrorAction 'Stop'
                Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CreateDBUser -ErrorAction 'Stop'
                $ExitValue.Text = "Rundeck DB created, Login created."
            }
            catch
            {
                $ExitValue.Text = "$_.ToString()"
            }
            return $ExitValue
        })
        
        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Installing DB"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()

        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox3.Controls.Add($Button6)

    $Button7.Text = "Test database" # on 3) Database screen
    $Button7.Location = "15,210"
    $Button7.Size = "100,25"
    $Button7.Visible = $false
    $Button7.add_Click({

        $Context = @{}
        $Context.SQLServer = $TextBox3.Text
        $Context.Username = $TextBox4.Text

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}

            # Prepare test statements
            $CheckDBStatement = "SELECT COUNT(*) AS DB_COUNT FROM sys.databases WHERE [name] = 'Rundeck'"
            $CheckLoginStatement = "SELECT COUNT(*) AS LOGIN_COUNT FROM sys.server_principals WHERE [name] = '$($Context.Username)'"
            $CheckRoleStatement = "USE Rundeck`nGO`nSELECT COUNT(*) AS USER_ROLES_COUNT`nFROM sys.database_role_members rm`nINNER JOIN sys.database_principals rp on rm.role_principal_id = rp.principal_id`nINNER JOIN sys.database_principals mp on rm.member_principal_id = mp.principal_id`nWHERE rp.name = 'db_owner' AND mp.name = 'RundeckDBUser'"
        
            $ExitValue.Result = ""
            # Execute SQL statements
            try
            {
                # Invoke 3 statements; all should return 1
                $DBCheckResult = Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CheckDBStatement -ErrorAction 'Stop'
                $LoginCheckResult = Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CheckLoginStatement -ErrorAction 'Stop'
                $RoleCheckResult = Invoke-SqlCmd -ServerInstance $Context.SQLServer -Query $CheckRoleStatement -ErrorAction 'Stop'

                if ($DBCheckResult.DB_COUNT -eq 1)
                {
                    $ExitValue.Result += "DB - OK; "
                }
                else {
                    $ExitValue.Result += "DB - Error; "
                }
                if ($LoginCheckResult.LOGIN_COUNT -eq 1)
                {
                    $ExitValue.Result += "Login - OK; "
                }
                else {
                    $ExitValue.Result += "Login - Error; "
                }
                if ($RoleCheckResult.USER_ROLES_COUNT -eq 1)
                {
                    $ExitValue.Result += "User - OK;"
                }
                else {
                    $ExitValue.Result += "User - Error;"
                }
                
                $ExitValue.Text = "Rundeck DB checked."
            }
            catch
            {
                $ExitValue.Text = "$_.ToString()"
            }
        
            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Testing DB"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        
        $TextBox6.Text = $BgJobResult.Result
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox3.Controls.Add($Button7)

    $Button8.Text = "Change settings and restart Rundeck" # 3) Database screen
    $Button8.Location = "15,250"
    $Button8.Size = "250,25"
    $Button8.add_Click({

        $Context = @{}
        # For script re-run - check for Rundeck Home in this session
        if (-not $Script:UsedRundeckDirectory)
        {
            $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $SelectFolderDialog.Description = "Select Rundeck Folder"
            $SelectFolderDialog.SelectedPath = $Script:DefaultRundeckDirectory
            $SelectFolderDialog.ShowDialog()
            $Script:UsedRundeckDirectory = $SelectFolderDialog.SelectedPath
        }
        $Context.SQLServer = $TextBox3.Text
        $Context.Username = $TextBox4.Text
        $Context.Password = $TextBox5.Text
        $Context.RundeckFolder = $Script:UsedRundeckDirectory
        $Context.SelectedAuth = $ComboBox1.SelectedItem

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}

            $ConfigFile = $Context.RundeckFolder + "\server\config\rundeck-config.properties"
            # Create BackUP
            Copy-Item -Path $ConfigFile -Destination "$($ConfigFile).bac"

            #Predefined vars
            $LocalDBConfig = "dataSource.dbCreate = update",
                            "dataSource.url = jdbc:h2:file:$($Context.RundeckFolder)/server/data/grailsdb;MVCC=true"

            $ExtwithWindowsAuthConfig = "rundeck.projectsStorageType=db",
                            "dataSource.dbCreate = update",
                            "dataSource.driverClassName = com.microsoft.sqlserver.jdbc.SQLServerDriver",
                            "dataSource.url = jdbc:sqlserver://$($Context.SQLServer);DatabaseName=Rundeck;integratedSecurity=true"
            
            $ExtwithSQLAuthConfig = "rundeck.projectsStorageType=db",
                            "dataSource.dbCreate = update",
                            "dataSource.driverClassName = com.microsoft.sqlserver.jdbc.SQLServerDriver",
                            "dataSource.url = jdbc:sqlserver://$($Context.SQLServer);DatabaseName=Rundeck",
                            "dataSource.username = $($Context.Username)",
                            "dataSource.password = $($Context.Password)"

            $ExtMySQL = "rundeck.projectsStorageType=db",
                            "dataSource.dbCreate = update",
                            "dataSource.driverClassName = com.mysql.jdbc.Driver",
                            "dataSource.url = jdbc:mysql://$($Context.SQLServer)/rundeck?autoReconnect=true&useSSL=false",
                            "dataSource.username = $($Context.Username)",
                            "dataSource.password = $($Context.Password)"

            $SearchConfigProps =    "rundeck.projectsStorageType",
                            "dataSource.dbCreate",
                            "dataSource.driverClassName",
                            "dataSource.url",
                            "dataSource.username",
                            "dataSource.password"

            # Read config file, remove previous DB config, insert new, save file
            $RundeckConfiguration = Get-Content -Path $ConfigFile
            $UpdateConfig = $RundeckConfiguration | Select-String -Pattern $SearchConfigProps -NotMatch

            if ($Context.SelectedAuth -eq "Use built-in database")
            {
                $UpdateConfig += $LocalDBConfig
            }
            elseif ($Context.SelectedAuth -eq "MS-SQL with Windows authentication")
            {
                $UpdateConfig += $ExtwithWindowsAuthConfig
            }
            elseif ($Context.SelectedAuth -eq "MS-SQL with SQL authentication")
            {
                $UpdateConfig += $ExtwithSQLAuthConfig
            }
            elseif ($Context.SelectedAuth -eq "MySQL")
            {
                $UpdateConfig += $ExtMySQL
            }

            Set-Content -Path $ConfigFile -Value $UpdateConfig

            Restart-Service RUNDECK

            $ExitValue.Text = "Config applied, service restarted."
            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Changing Rundeck config"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox3.Controls.Add($Button8)

    $Button9.Text = "Download and setup DLL for windows auth"
    $Button9.Location = "15,165"
    $Button9.Size = "250,25"
    $Button9.Visible = $false
    $Button9.add_Click({

        $Context = @{}
        $Context.SqlDjbcDll = $Script:SqlDjbcDll

        $BackgroundScript = [Powershell]::Create().AddScript({

            Param($Context)
            $ExitValue = @{}

            $TargetFolder = "C:\Windows\System32\"
            Start-BitsTransfer -Destination $TargetFolder -Source $Context.SqlDjbcDll

            regsrv32 /s C:\Windows\System32\sqljdbc_auth.dll

            $ExitValue.Text = "DLL uploaded and registered"

            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Downloading and registering DLL"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox3.Controls.Add($Button9)

    $Button10.Text = "Select Base OUs"
    $Button10.Location = "15,160"
    $Button10.Size = "110,25"
    $Button10.Visible = $false
    $Button10.add_Click({


        $BackgroundScript = [Powershell]::Create().AddScript({
            $ExitValue = @{}

            $AD_DomainRoot = ([adsisearcher]“”).SearchRoot.Path.Replace("LDAP://","")
            $AD_Containers = ([adsisearcher]“objectcategory=container”).findall().Path.Replace("LDAP://","") | Where-Object {$_ -notlike "*CN=System*"}
            $AD_Ous = ([adsisearcher]“objectcategory=organizationalunit”).findall().Path.Replace("LDAP://","")

            $OUs = @($AD_DomainRoot) + @($AD_Containers) + @($AD_Ous)

            
            if ($OUs)
            {
                $ExitValue.OUs = $OUs
                $ExitValue.Text = "OUs discovered"
            }
            else {
                $ExitValue.OUs = $null
                $ExitValue.Text = "No OUs discovered"
            }
            
            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Browsing OU"

        while (-not $BG_Job.IsCompleted)
            {
                Start-Sleep -Milliseconds 500
                $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
                $Label1.Visible = -not $Label1.Visible
                $Label1.Text = "Busy: $ExecutionTime s"
                $EndDate = Get-Date
            }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()

        $Label2.Text = $BgJobResult.Text
        

        $SelectedUsersOU = $BgJobResult.OUs | Out-GridView -PassThru -Title "Select Users OU"
        if ($SelectedUsersOU)
        {
            $TextBox11.Text = $SelectedUsersOU
        }

        $SelectedGroupsOU = $BgJobResult.OUs | Out-GridView -PassThru -Title "Select Groups OU"
        if ($SelectedGroupsOU)
        {
            $TextBox12.Text = $SelectedGroupsOU
        }
    })
    $GroupBox4.Controls.Add($Button10)

    $Button11.Text = "Test AD lookup"
    $Button11.Location = "15,220"
    $Button11.Size = "110,25"
    $Button11.Visible = $false
    $Button11.add_Click({

        $Context = @{}
        $Context.Server = $TextBox14.Text
        $Context.Username = $TextBox9.Text
        $Context.Password = $TextBox10.Text
        $Context.OuSearch = $TextBox11.Text
        $Context.OuGroups = $TextBox12.Text

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}

            # Create Ad credential object
            $SecPassObj = ConvertTo-SecureString $Context.Password -AsPlainText -Force
            $AdCreds = New-Object System.Management.Automation.PSCredential ($Context.Username, $SecPassObj)

            try {
                $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher

                $Searcher.Filter = "(objectCategory=User)"

                $DomainDN = "LDAP://" + $Context.OuSearch
                $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN, $Context.Username, $Context.Password
                $Searcher.SearchRoot = $Domain

                $LdapUsers = $Searcher.FindAll()
                ###

                $Searcher.Filter = "(objectCategory=group)"

                $DomainDN = "LDAP://" + $Context.OuGroups
                $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $DomainDN, $Context.Username, $Context.Password
                $Searcher.SearchRoot = $Domain

                $LdapGroups = $Searcher.FindAll()
                
                $ExitValue.Result = "Discovered $($LdapUsers.Count) users and $($LdapGroups.Count) groups."
                $Count = $LdapUsers.Count                
            }
            catch {
                $ExitValue.Result = "Error on AD lookup"
            }

            $ExitValue.Text = "AD test finished"
                    
            return $ExitValue
       })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Testing AD"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        
        $TextBox13.Text = $BgJobResult.Result
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox4.Controls.Add($Button11)

    $Button12.Text = "Change auth config and restart Rundeck" # 3) Database screen
    $Button12.Location = "15,260"
    $Button12.Size = "220,25"
    $Button12.add_Click({

        $Context = @{}

        # For script re-run - check for Rundeck Home in this session
        if (-not $Script:UsedRundeckDirectory)
        {
            $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $SelectFolderDialog.Description = "Select Rundeck Folder"
            $SelectFolderDialog.SelectedPath = $Script:DefaultRundeckDirectory
            $SelectFolderDialog.ShowDialog()
            $Script:UsedRundeckDirectory = $SelectFolderDialog.SelectedPath
        }
        #Convert Pre-logon name to Sam AccountName Extract DN from username
        $SamAccName = $TextBox9.Text.Split("\")[-1]
        ##
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectClass=user)(sAMAccountName= $SamAccName))"
        
        $UserDN = $Searcher.FindOne().Path.Replace("LDAP://","")

        $Context.Server = $TextBox14.Text
        $Context.Username = $UserDN
        $Context.Password = $TextBox10.Text
        $Context.OuUserSearch = $TextBox11.Text
        $Context.OuGroupSearch = $TextBox12.Text
        $Context.RundeckFolder = $Script:UsedRundeckDirectory
        $Context.SelectedAuth = $ComboBox2.SelectedItem

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}

            $MainConfigFile = $Context.RundeckFolder + "\server\config\rundeck-config.properties"
            $ProfileConfigFile = $Context.RundeckFolder + "\etc\profile.bat"
            $JaasConfigFile = $Context.RundeckFolder + "\server\config\jaas-ldap.conf"

            # Create BackUP
            Copy-Item -Path $MainConfigFile -Destination "$($MainConfigFile).bac"
            Copy-Item -Path $ProfileConfigFile -Destination "$($ProfileConfigFile).bac"

            # Patterns to remove
            $MainConfigPattern = "rundeck.security.syncLdapUser"
            $ProfileConfigPattern = "-Drundeck.jaaslogin",
                                    "-Dloginmodule.conf.name",
                                    "-Dloginmodule.name"

            # Main config file without Pattern, Extract Data from profile
            $MainRundeckConfiguration = Get-Content -Path $MainConfigFile
            $MainCleanConfig = $MainRundeckConfiguration | Select-String -Pattern $MainConfigPattern -NotMatch

            $ProfileRundeckConfiguration = Get-Content -Path $ProfileConfigFile
            $ProfileCleanConfig = $ProfileRundeckConfiguration | Select-String -Pattern "RDECK_SSL_OPT" -NotMatch
            
            $ProfileLineConfigArr = ($ProfileRundeckConfiguration | Select-String -Pattern "RDECK_SSL_OPT").ToString().Replace('"','').Split(" ") | Select-String -Pattern $ProfileConfigPattern -NotMatch
            $ProfileLineConfig = $ProfileLineConfigArr -join " "

            # Save clean config or save modified config
            if ($Context.SelectedAuth -eq "Use built-in authentication")
            {
                Set-Content -Path $MainConfigFile -Value $MainCleanConfig

                $ProfileCleanConfig += $ProfileLineConfig
                Set-Content -Path $ProfileConfigFile -Value $ProfileCleanConfig
            }
            elseif ($Context.SelectedAuth -eq "Use Active Directory")
            {
                $MainCleanConfig += "rundeck.security.syncLdapUser=true"
                Set-Content -Path $MainConfigFile -Value $MainCleanConfig

                $ProfileLineConfig += " -Drundeck.jaaslogin=true -Dloginmodule.conf.name=jaas-ldap.conf -Dloginmodule.name=activedirectory"
                $ProfileCleanConfig += $ProfileLineConfig
                Set-Content -Path $ProfileConfigFile -Value $ProfileCleanConfig

                $JaasConfig = 'activedirectory {
                    com.dtolabs.rundeck.jetty.jaas.JettyCachingLdapLoginModule required
                    debug="true"
                    contextFactory="com.sun.jndi.ldap.LdapCtxFactory"
                    providerUrl="ldap://' + $Context.Server + '
                    bindDn="' + $Context.Username + '"
                    bindPassword="' + $Context.Password +'"
                    authenticationMethod="simple"
                    forceBindingLogin="true"
                    userBaseDn="' + $Context.OuUserSearch + '"
                    userRdnAttribute="sAMAccountName"
                    userIdAttribute="sAMAccountName"
                    userPasswordAttribute="unicodePwd"
                    userObjectClass="user"
                    roleBaseDn="' + $Context.OuGroupSearch + '"
                    roleNameAttribute="cn"
                    roleMemberAttribute="member"
                    roleObjectClass="group"
                    cacheDurationMillis="300000"
                    reportStatistics="true"
                    userLastNameAttribute="sn"
                    userFirstNameAttribute="givenName"
                    userEmailAttribute="mail";
                };'

                $JaasConfig | Out-File $JaasConfigFile -Encoding ascii

            }

            Restart-Service RUNDECK

            $ExitValue.Text = "Config applied, service restarted."
            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Changing Rundeck config"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox4.Controls.Add($Button12)

    $Button13.Text = "Test WinRM connectivity"
    $Button13.Location = "50,260"
    $Button13.Size = "200,25"
    $Button13.add_Click({
        $Computers = $TextBox15.Text
        $ArrComputers = $Computers.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)

        $Context = @{}
        $Context.Computers = $ArrComputers
        if ($TextBox7.Text -and $TextBox8.Text)
        {
            $Context.UseCreds = $true
            $Context.Username = $TextBox7.Text
            $Context.Password = $TextBox8.Text
        }

        $BackgroundScript = [Powershell]::Create().AddScript({
            Param($Context)
            $ExitValue = @{}

            $ExitValue.Result = @()

            $PsOption = New-PSSessionOption -OpenTimeout 1000 -OperationTimeout 1000 -CancelTimeout 1000 -MaxConnectionRetryCount 1

            if ($Context.UseCreds)
            {
                $SecPass = ConvertTo-SecureString $($Context.Password) -AsPlainText -Force
                $Creds = New-Object System.Management.Automation.PSCredential ($($Context.Username), $SecPass)
                foreach ($Comp in $Context.Computers)
                {
                    try {
                        $session = New-PSSession -Authentication Default -SessionOption $PsOption -Computer $Comp -Credential $Creds -ErrorAction Stop
                        $Result = "OK"
                    }
                    catch {
                        $Result = "Error:" + $_.FullyQualifiedErrorId
                    }
                    $CompTestResulr = [pscustomobject]@{"Comp" = $Comp; "Result" = $Result}
                    $ExitValue.Result += $CompTestResulr
                }
            }
            else 
            {
                foreach ($Comp in $Context.Computers)
                {
                    try {
                        $session = New-PSSession -Authentication Default -SessionOption $PsOption -Computer $Comp -ErrorAction Stop
                        $Result = "OK"
                    }
                    catch {
                        $Result = "Error:" + $_.FullyQualifiedErrorId
                    }
                    $CompTestResulr = [pscustomobject]@{"Comp" = $Comp; "Result" = $Result}
                    $ExitValue.Result += $CompTestResulr
                }
            }

            $ExitValue.Text = "Test Finished"

            return $ExitValue
        })

        # Run async script block
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $BackgroundScript.AddArgument($Context)
        $BackgroundScript.Runspace = $Runspace
        $BG_Job = $BackgroundScript.BeginInvoke()

        $StartDate = Get-Date
        $EndDate = Get-Date
        $Label1.ForeColor = "Red"
        $Label2.Text = "Testing WinRM"

        while (-not $BG_Job.IsCompleted)
        {
            Start-Sleep -Milliseconds 500
            $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
            $Label1.Visible = -not $Label1.Visible
            $Label1.Text = "Busy: $ExecutionTime s"
            $EndDate = Get-Date
        }

        $ExecutionTime = [int](New-TimeSpan -Start $StartDate -End $EndDate).TotalSeconds
        $Label1.ForeColor = "Black"
        $Label1.Visible = $true
        $Label1.Text = "Done in $ExecutionTime s"

        $BgJobResult = $BackgroundScript.EndInvoke($BG_Job)
        $BackgroundScript.Dispose()
        $Runspace.Close()
        

        $TextBox16.Text = ""

        $BgJobResult.Result | Foreach {$TextBox16.Text += $_.Comp + "....." + $_.Result + [Environment]::NewLine}
        $Label2.Text = $BgJobResult.Text
    })
    $GroupBox5.Controls.Add($Button13)


    $LinkLabel1.Text = ""
    $LinkLabel1.Location = "440,270"
    $LinkLabel1.Size = "200,15"
    $LinkLabel1.add_LinkClicked({
        Start-Process $LinkLabel1.Text
    })
    $GroupBox2.Controls.Add($LinkLabel1)

    $Label1.Text = "Stand By."
    $Label1.Location = "25,35"
    $Label1.Size = "200,15"
    $Form.Controls.Add($Label1)

    $Label2.Text = "Ready."
    $Label2.Location = "25,55"
    $Label2.Size = "400,15"
    $Form.Controls.Add($Label2)

    $Label3.Text = "Database Server:"  # 3) Database screen
    $Label3.Location = "15,60"
    $Label3.Size = "95,15"
    $Label3.Visible = $false
    $GroupBox3.Controls.Add($Label3)

    $Label4.Text = "User:"  # 3) Database screen
    $Label4.Location = "15,80"
    $Label4.Size = "45,15"
    $Label4.Visible = $false
    $GroupBox3.Controls.Add($Label4)

    $Label5.Text = "SQL password:" # on 3) Database screen
    $Label5.Location = "15,102"
    $Label5.Size = "85,15"
    $Label5.Visible = $false
    $GroupBox3.Controls.Add($Label5)

    $Label6.Text = "Other Steps needed ( outside this Installer):"
    $Label6.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.5,([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
    $Label6.Location = "30,30"
    $Label6.Size = "345,15"
    $GroupBox2.Controls.Add($Label6)

    $Label7.Text = "1. Create a 'service account' i.e. a domain user account that is in local Administrators group on this computer. eg rundeck_svc"
    $Label7.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.5)
    $Label7.Location = "40,50"
    $Label7.Size = "670,15"
    $GroupBox2.Controls.Add($Label7)

    $Label8.Text = "Must use DOMAIN\USER format"
    $Label8.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.5,[System.Drawing.FontStyle]::Bold)
    $Label8.Location = "40,70"
    $Label8.Size = "450,15"
    $GroupBox2.Controls.Add($Label8)


    $Label9.Text = "User:" # 2) Rundeck user
    $Label9.Location = "30,95"
    $Label9.Size = "50,15"
    $GroupBox2.Controls.Add($Label9)

    $Label10.Text = "Password:"
    $Label10.Location = "30,120"
    $Label10.Size = "60,15"
    $GroupBox2.Controls.Add($Label10)

    $Label11.Text = "User:" # on 4) Authentication screen
    $Label11.Location = "15,90"
    $Label11.Size = "65,15"
    $Label11.Visible = $false
    $GroupBox4.Controls.Add($Label11)

    $Label12.Text = "Password:" # on 4) Authentication screen
    $Label12.Location = "15,120"
    $Label12.Size = "60,15"
    $Label12.Visible = $false
    $GroupBox4.Controls.Add($Label12)

    $Label13.Text = "A Domain Controller:" # on 4) Authentication screen
    $Label13.Location = "15,65"
    $Label13.Size = "110,15"
    $Label13.Visible = $false
    $GroupBox4.Controls.Add($Label13)

    $Label14.Text = "Base OU for users:"
    $Label14.Location = "130,165"
    $Label14.Size = "100,15"
    $Label14.Visible = $false
    $GroupBox4.Controls.Add($Label14)

    $Label15.Text = "Base OU for groups:"
    $Label15.Location = "130,190"
    $Label15.Size = "110,15"
    $Label15.Visible = $false
    $GroupBox4.Controls.Add($Label15)

    $Label16.Text = "Type a few servers, one per line:"
    $Label16.Location = "50,35"
    $Label16.Size = "170,15"
    $GroupBox5.Controls.Add($Label16)

    $ComboBox1 = New-Object System.Windows.Forms.ComboBox
    $ComboBox1.Size = "240,25"
    $ComboBox1.Location = "25,25"
    $ComboBox1.Items.Add("Use built-in database")
    $ComboBox1.Items.Add("MS-SQL with Windows authentication")
    $ComboBox1.Items.Add("MS-SQL with SQL authentication")
    $ComboBox1.Items.Add("MySQL")
    $ComboBox1.SelectedItem = $ComboBox1.Items[0]
    $ComboBox1.DropDownStyle = "DropDownList"
    $ComboBox1.add_SelectedIndexChanged({
        if ($ComboBox1.SelectedItem -eq "Use built-in database")
        {
            $TextBox3.Visible = $false
            $TextBox4.Visible = $false
            $TextBox5.Visible = $false
            $TextBox6.Visible = $false
            $Label3.Visible = $false
            $Label4.Visible = $false
            $Label5.Visible = $false
            $Button6.Visible = $false
            $Button7.Visible = $false
            $Button9.Visible = $false
        }
        elseif ($ComboBox1.SelectedItem -eq "MS-SQL with Windows authentication")
        {
            $TextBox3.Visible = $true
            $TextBox4.Visible = $true
            $TextBox4.Enabled = $false
            $TextBox4.Text = $TextBox7.Text
            $TextBox5.Visible = $false
            $TextBox6.Visible = $true
            $Label3.Visible = $true
            $Label4.Visible = $true
            $Label5.Visible = $false
            $Button6.Visible = $true
            $Button7.Visible = $true
            $Button9.Visible = $true
        }
        elseif ($ComboBox1.SelectedItem -eq "MS-SQL with SQL authentication")
        {
            $TextBox3.Visible = $true
            $TextBox4.Visible = $true
            $TextBox4.Enabled = $true
            $TextBox4.Text = ""
            $TextBox5.Visible = $true
            $TextBox6.Visible = $true
            $Label3.Visible = $true
            $Label4.Visible = $true
            $Label5.Visible = $true
            $Button6.Visible = $true
            $Button7.Visible = $true
            $Button9.Visible = $false
        }
        elseif ($ComboBox1.SelectedItem -eq "MySQL")
        {
            $TextBox3.Visible = $true
            $TextBox4.Visible = $true
            $TextBox4.Enabled = $true
            $TextBox4.Text = ""
            $TextBox5.Visible = $true
            $TextBox6.Visible = $true
            $Label3.Visible = $true
            $Label4.Visible = $true
            $Label5.Visible = $true
            $Button6.Visible = $true
            $Button7.Visible = $true
            $Button9.Visible = $false
        }
    })
    $GroupBox3.Controls.Add($ComboBox1)

    $ComboBox2 = New-Object System.Windows.Forms.ComboBox
    $ComboBox2.Size = "200,25"
    $ComboBox2.Location = "25,25"
    $ComboBox2.Items.Add("Use built-in authentication")
    $ComboBox2.Items.Add("Use Active Directory")
    $ComboBox2.SelectedItem = $ComboBox2.Items[0]
    $ComboBox2.DropDownStyle = "DropDownList"
    $ComboBox2.add_SelectedIndexChanged({

        $TextBox9.Text = $TextBox7.Text
        $TextBox10.Text = $TextBox8.Text

        if ($ComboBox2.SelectedItem -eq "Use built-in authentication")
        {
            $TextBox9.Visible = $false
            $TextBox10.Visible = $false
            $TextBox11.Visible = $false
            $TextBox12.Visible = $false
            $TextBox13.Visible = $false
            $TextBox14.Visible = $false
            $Label11.Visible = $false
            $Label12.Visible = $false
            $Label13.Visible = $false
            $Label14.Visible = $false
            $Label15.Visible = $false
            $Button10.Visible = $false
            $Button11.Visible = $false
        }
        elseif ($ComboBox2.SelectedItem -eq "Use Active Directory")
        {
            $TextBox9.Visible = $true
            $TextBox10.Visible = $true
            $TextBox11.Visible = $true
            $TextBox12.Visible = $true
            $TextBox13.Visible = $true
            $TextBox14.Visible = $true
            $Label11.Visible = $true
            $Label12.Visible = $true
            $Label13.Visible = $true
            $Label14.Visible = $true
            $Label15.Visible = $true
            $Button10.Visible = $true
            $Button11.Visible = $true
        }
    })
    $GroupBox4.Controls.Add($ComboBox2)
    
    # Java Status
    $TextBox1.Location = "130,203"
    $TextBox1.Size = "200,25"
    $TextBox1.ReadOnly = $true
    $GroupBox1.Controls.Add($TextBox1)

    # Rundeck Status
    $TextBox2.Location = "140,242"
    $TextBox2.Size = "290,25"
    $TextBox2.ReadOnly = $true
    $GroupBox2.Controls.Add($TextBox2)

    # SQL server field
    $TextBox3.Location = "110,58"
    $TextBox3.Size = "120,25"
    $TextBox3.Visible = $false
    $GroupBox3.Controls.Add($TextBox3)

    # SQL user  on 3) Database screen
    $TextBox4.Location = "110,78"
    $TextBox4.Size = "120,25"
    $TextBox4.Visible = $false
    $GroupBox3.Controls.Add($TextBox4)

    # SQL password   on 3) Database screen
    $TextBox5.Location = "110,100"
    $TextBox5.Size = "120,25"
    $TextBox5.Visible = $false
    $TextBox5.PasswordChar = "*"
    $GroupBox3.Controls.Add($TextBox5)

    # Test MS-SQL result on 3) Database
    $TextBox6.Location = "125,212"
    $TextBox6.Size = "240,25"
    $TextBox6.Visible = $false
    $TextBox6.ReadOnly = $true
    $GroupBox3.Controls.Add($TextBox6)

    # Rundeck service username
    $TextBox7.Location = "90, 90"
    $TextBox7.Size = "160,25"
    $GroupBox2.Controls.Add($TextBox7)

    # Rundeck service password
    $TextBox8.Location = "90, 115"
    $TextBox8.Size = "160,25"
    $TextBox8.PasswordChar = "*"
    $GroupBox2.Controls.Add($TextBox8)

    # 4) Authentication Domain Controller
    $TextBox14.Location = "125,60"
    $TextBox14.Size = "120,25"
    $TextBox14.Visible = $false
    $GroupBox4.Controls.Add($TextBox14)

    # 4) Authentication user
    $TextBox9.Location = "125,90"
    $TextBox9.Size = "120,25"
    $TextBox9.Visible = $false
    $GroupBox4.Controls.Add($TextBox9)

    # 4) Authentication password
    $TextBox10.Location = "125,120"
    $TextBox10.Size = "120,25"
    $TextBox10.PasswordChar = "*"
    $TextBox10.Visible = $false
    $GroupBox4.Controls.Add($TextBox10)

    # Browse for users OU result
    $TextBox11.Location = "240,160"
    $TextBox11.Size = "450,25"
    $TextBox11.Visible = $false
    $GroupBox4.Controls.Add($TextBox11)

    # Browse for Groups OU result
    $TextBox12.Location = "240,190"
    $TextBox12.Size = "450,25"
    $TextBox12.Visible = $false
    $GroupBox4.Controls.Add($TextBox12)

    # Test AD Lookup result
    $TextBox13.Location = "160,225"
    $TextBox13.Size = "240,25"
    $TextBox13.Visible = $false
    $TextBox13.ReadOnly = $true
    $GroupBox4.Controls.Add($TextBox13)

    #WinRM Server list to check
    $TextBox15.Location = "50,50"
    $TextBox15.Size = "230,200"
    $TextBox15.Multiline = $true
    $GroupBox5.Controls.Add($TextBox15)

    #WinRM Server check results
    $TextBox16.Location = "290,50"
    $TextBox16.Size = "350,200"
    $TextBox16.Multiline = $true
    $TextBox16.ReadOnly = $True
    $GroupBox5.Controls.Add($TextBox16)

    $Form.Controls.Add($MenuStrip)
    $Form.Controls.Add($GroupBox1)
    $Form.Controls.Add($GroupBox2)
    $Form.Controls.Add($GroupBox3)
    $Form.Controls.Add($GroupBox4)
    $Form.Controls.Add($GroupBox5)
    
    $Form.Text = "Rundeck Installer for Windows"
    $Form.Size = "800,450"
    $Form.MaximizeBox = $false
    $Form.FormBorderStyle = "FixedDialog"
    $Form.add_Load({
        $GroupBox1.Visible = $true

        $MissingModules = ""

        if ( -not (get-module -name SqlServer -listavailable)) {
            $MissingModules += "SqlServer"
        }

        if ($MissingModules)
        {
            [System.Windows.MessageBox]::Show("Required modules missing: $MissingModules. Install and restart script", "Missing Modules", "OK", "Warning")
        }
    })
    $Form.ShowDialog() | Out-Null 
}

#Minimize parent powershell window
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
$showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)

Show-MainForm