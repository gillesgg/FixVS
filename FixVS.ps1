
function getProgramFiles32bit()
{
  $out = ${env:PROGRAMFILES(X86)}
  if ($null -eq $out) {
    $out = ${env:PROGRAMFILES}
  }

  if ($null -eq $out) {
    throw "Could not find [Program Files 32-bit]"
  }

  return $out
}


# ;Switch.System.Windows.Interop.MouseInput.OptOutOfMoveToChromedWindowFix=true;Switch.System.Windows.Interop.MouseInput.DoNotOptOutOfMoveToChromedWindowFix=true
#

function validateOptOutOfMoveToChromedWindowFix($itemssection)
{
   foreach ($itemsection in $itemssection)
   {
    if ($itemsection -eq 'Switch.System.Windows.Interop.MouseInput.OptOutOfMoveToChromedWindowFix=true')
    {
        return $true
    }
   }
   return $false
}

# Update the section
function validateDoNotOptOutOfMoveToChromedWindowFix($itemssection, [ref] $OptOutOfMoveToChromedWindowFix, [ref] $DoNotOptOutOfMoveToChromedWindowFix)
{
   foreach ($itemsection in $itemssection)
   {    
    if ($itemsection -eq 'Switch.System.Windows.Interop.MouseInput.DoNotOptOutOfMoveToChromedWindowFix=true')
    {
        return $true
    }
   }
   return $false
}

# Update the devenv.exe.config
function fixsection($itemssectionText, $OptOutOfMoveToChromedWindowFix, $DoNotOptOutOfMoveToChromedWindowFix)
{
    if ($OptOutOfMoveToChromedWindowFix -eq $false)
    {
        $itemssectionText += ';Switch.System.Windows.Interop.MouseInput.OptOutOfMoveToChromedWindowFix=true'

    }

    if ($DoNotOptOutOfMoveToChromedWindowFix -eq $false)
    {
        $itemssectionText += ';Switch.System.Windows.Interop.MouseInput.DoNotOptOutOfMoveToChromedWindowFix=true'
    }

    #Write-Host "new value=" $itemssectionText -ForegroundColor Yellow
    return $itemssectionText
}

# Backup the devenv.exe.config file
function backupconfig($appConfig, $devenvconfig)
{
    Try 
    {
        $backupconfig = $devenvconfig + "_backup.config"
        $FileExists = Test-Path $backupconfig
        if ($FileExists-eq $false)
        {
            Copy-Item $devenvconfig -Destination $backupconfig -Force
        }
    }
    Catch
    {
        #Write-Host "Message: [$($_.Exception.Message)"] -ForegroundColor Red
        Write-Host "Unable to backup the file:" $backupconfig -ForegroundColor Red
        return $false
    }
    return $true
}

# save the devenv.exe.config file
function saveconfig($appConfig, $devenvconfig)
{
    Try 
    {
        $appConfig.Save($devenvconfig)
    }
    Catch
    {
        #Write-Host "Message: [$($_.Exception.Message)"] -ForegroundColor Red
        Write-Host "Unable to save the file:" $backupconfig -ForegroundColor Red
    }
}


# test if the script start as Admin
function testadmin()
{
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) 
    {
        Write-Host "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again." -ForegroundColor Red
        return $false
    }
    else 
    {
        Write-Host "Code is running as administrator — go on executing the script..." -ForegroundColor Yellow
        return $true
    }
}

function rollbackissue($ProgramPath)
{
    foreach ($instance in $ProgramPath)
    {
        $devenvconfigbackup = $instance + '.config_backup.config'

        $devenvconfig = $instance + '.config'

        if (Test-Path $devenvconfigbackup)
        {
            Copy-Item $devenvconfigbackup -Destination $devenvconfig -Force
            Write-Host "Rollback " $devenvconfig " from : " $devenvconfigbackup -ForegroundColor Green
            Remove-Item $devenvconfigbackup -Force
        }
        else
        {
            Write-Host "The backup file :" $devenvconfigbackup " do not exist" -ForegroundColor Red

        }
   }
}

function fixwpfissue($ProgramPath)
{
    foreach ($instance in $ProgramPath)
    {
        $devenvconfig = $instance + '.config'
        if (Test-Path $devenvconfig)
        {
            $appConfig = [xml](get-content $devenvconfig)
            
            $root = $appConfig.get_DocumentElement(); 
            $AppContextSwitchOverrides = $root.runtime.AppContextSwitchOverrides
            if ($AppContextSwitchOverrides)
            {
                $OptOutOfMoveToChromedWindowFix = $false
                $DoNotOptOutOfMoveToChromedWindowFix = $false
                $itemssection = $AppContextSwitchOverrides.value.Split(';')
                $itemssectionText = $AppContextSwitchOverrides.value
                $OptOutOfMoveToChromedWindowFix = validateOptOutOfMoveToChromedWindowFix $itemssection
                $DoNotOptOutOfMoveToChromedWindowFix = validateDoNotOptOutOfMoveToChromedWindowFix $itemssection                  
                if ($OptOutOfMoveToChromedWindowFix -eq $false -or $OptOutOfMoveToChromedWindowFix -eq $false)
                {
                    Write-Host "Patch Visual Studio config file:" $devenvconfig  -ForegroundColor Green
                    $itemssectionText = fixsection $itemssectionText $OptOutOfMoveToChromedWindowFix $DoNotOptOutOfMoveToChromedWindowFix
                    $root.runtime.AppContextSwitchOverrides.value = $itemssectionText
                    $ret = backupconfig $appConfig $devenvconfig
                    if ($ret -eq $true)
                    {
                        saveconfig $appConfig $devenvconfig
                        fixvsixinstance $instance                        
                    }
                 }
                 else
                 {
                    Write-Host "Visual Studio patch for config file:" $devenvconfig  " already set" -ForegroundColor Yellow
                 }
             }
             else
             {
                Write-Host "the section configuration.runtime.AppContextSwitchOverrides for path " $devenvconfig  " do not exist" -ForegroundColor Red               
             }
        }
    }
}



#Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\11.0 - 2012


function getLatestVisualStudioWithDesktopWorkloadPathlegacy($ProgramPath)
{
    $vs2012key = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\11.0'
    $vs2015key = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0'
    $valuekey = 'InstallDir'
     
    if (Test-Path $vs2012key) 
    {
        $vs2012 = Get-ItemProperty $vs2012key $valuekey  -ErrorAction SilentlyContinue
        if ($vs2012)
        {
            $vs2012path = Join-Path $vs2012.InstallDir "devenv.exe"
            $ProgramPath.Add($vs2012path)  > $null
        }
    }

    if (Test-Path $vs2015key)
    {
        $vs2015 = Get-ItemProperty $vs2015key $valuekey  -ErrorAction SilentlyContinue
        if ($vs2015)
        {
            $vs2015path = Join-Path $vs2015.InstallDir "devenv.exe"
            $ProgramPath.Add($vs2015path)  > $null
        }
    }
}

# Get the list of devenv.exe using vswhere.exe
# $ProgramPath - list of devenv.exe
function getLatestVisualStudioWithDesktopWorkloadPath($ProgramPath)
{ 
 
  $programFiles = getProgramFiles32bit
  $vswhereExe = "$programFiles\Microsoft Visual Studio\Installer\vswhere.exe"
  if (Test-Path $vswhereExe)
  {
    $output = & $vswhereExe -format xml
    [xml]$asXml = $output
    foreach ($instance in $asXml.instances.instance)
    {
        

        $ProgramPath.Add($instance.productPath)  > $null
    }

    if ($ProgramPath.Count -eq 0)
    {
      Write-Host "Could not locate any installation of Visual Studio" -ForegroundColor Red
      return $null;
    }    
  }
  else 
  {
    Write-Host "Could not locate vswhere at $vswhereExe" -ForegroundColor Red
  }
}

# Run devenv /updateConfiguration to update the configuration
# $ProgramPath - list of devenv.exe 
function fixvsixinstance($ProgramPath)
{
    if (Test-Path $ProgramPath)
    {
         Start-Process $ProgramPath -ArgumentList "/updateConfiguration"
    }
}


$param1=$args[0]




if ($param1 -eq '-help')
{
      Write-Host "Fix Visual Studio 2019 and Visual Studio 2017 crash issue with KB4598301" -ForegroundColor Green
      Write-Host "Usage " -ForegroundColor Green
      Write-Host "FixVS              : Fix the issue" -ForegroundColor Green
      Write-Host "FixVS -rollback    : undo the fix" -ForegroundColor Green
      return;
}



if (testadmin -eq $true)
{
     $ProgramPath = New-Object System.Collections.ArrayList($null)
     getLatestVisualStudioWithDesktopWorkloadPath $ProgramPath
     getLatestVisualStudioWithDesktopWorkloadPathlegacy $ProgramPath
     if ($param1 -eq '-rollback')
     {
        rollbackissue $ProgramPath
     }
     else
     {
        if ($args.Count -eq 0)
        {
            fixwpfissue $ProgramPath
        }
        else
        {
          Write-Host "Fix Visual Studio 2019 and Visual Studio 2017 crash issue with KB4598301" -ForegroundColor Green
          Write-Host "Usage " -ForegroundColor Green
          Write-Host "FixVS              : Fix the issue" -ForegroundColor Green
          Write-Host "FixVS -rollback    : undo the fix" -ForegroundColor Green
          return;
        }
     }
}