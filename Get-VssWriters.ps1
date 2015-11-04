function Get-VssWriters {
<# 
 .Synopsis
  Function to get information about VSS Writers.

 .Description
  Function will parse information from VSSAdmin tool and return object containing
  WriterName, StateID, StateDescription, and LastError
  Function will display a progress bar while it retrives information from different
  computers.

 .Parameter ComputerName
  This is the name (not IP address) of the computer. 
  If absent, localhost is assumed.

 .Example
  Get-VssWriters
  This example will return a list of VSS Writers on localhost

 .Example
  # Get VSS Writers on localhost, sort list by WriterName
  $VssWriters = Get-VssWriters | Sort "WriterName" 
  $VssWriters | FT -AutoSize # Displays it on screen
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Get VSS Writers on the list of $Computers, sort list by ComputerName
  $Computers = "xHost11","notThere","xHost12"
  $VssWriters = Get-VssWriters -ComputerName $Computers -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Reports any errors on VSS Writers on the computers listed in MyComputerList.txt, sorts list by ComputerName
  $Computers = Get-Content ".\MyComputerList.txt"
  $VssWriters = Get-VssWriters $Computers -Verbose | 
    Where { $_.StateDesc -ne 'Stable' } | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV 
 
 .Example
  # Get VSS Writers on all computers in current AD domain, sort list by ComputerName
  $Computers = (Get-ADComputer -Filter *).Name
  $VssWriters = Get-VssWriters $Computers -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Get VSS Writers on all Hyper-V hosts in current AD domain, sort list by ComputerName
  $FilteredComputerList = $null
  $Computers = (Get-ADComputer -Filter *).Name 
  Foreach ($Computer in $Computers) {
      if (Get-WindowsFeature -ComputerName $Computer -ErrorAction SilentlyContinue | 
        where { $_.Name -eq "Hyper-V" -and $_.InstallState -eq "Installed"}) {
          $FilteredComputerList += $Computer
      }
  }
  $VssWriters = Get-VssWriters $FilteredComputerList -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .OUTPUT
  Scripts returns a PS Object with the following properties:
    ComputerName                                                                                                                                                                           
    WriterName  
    StateID                                                                                                                                                                                
    StateDesc                                                                                                                                                                              
    LastError                                                                                                                                                                              

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 09/17/2014

#>

    [CmdletBinding(SupportsShouldProcess=$false,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$false,
               ValueFromPipeLine=$true,
               ValueFromPipeLineByPropertyName=$true,
               Position=0)]
        [ValidateNotNullorEmpty()]
        [String[]]$ComputerName = $env:COMPUTERNAME
    )

    Write-Debug "Capturing Vss Writers"
    $Writers = @()
    $RawWriters = @()
    $progressRate = 0
    foreach($computer in $ComputerName)
    {
        $percentCompleteRate = "{0:N0}" -f ($progressRate*100/$ComputerName.Count)
        Write-Progress -Activity "Processing $computer." -Status "$progressRate% Complete:" -PercentComplete $progressRate
        <#
        $progress = "{0:N0}" -f ($progressRate*100/$ComputerName.Count)
        Write-Progress -Activity "Processing computer $computer ... $progressRate out of $($ComputerName.Count) computers" `
            -PercentComplete $progress -Status "Please wait" -CurrentOperation "$progress% complete"
        #>
        if($computer -ne $env:COMPUTERNAME)
        {
            try
            {
                $RawWriters += Invoke-Command -ComputerName $computer -ScriptBlock { VssAdmin List Writers } -ErrorAction Stop | Select-Object -Skip 3 | Select-Object -SkipLast 1
            }
            catch
            {
                Write-Warning "Could not access $computer."
            }
        }
        else
        {
            $RawWriters += VssAdmin List Writers | Select-Object -Skip 3 | Select-Object -SkipLast 1
        }

        Write-Debug "Building results"
        for ($i=0; $i -lt $RawWriters.Count/6; $i++) {
            $WriterIdX = $RawWriters[($i*6)+1].Trim().IndexOf("{")
            $WriterIdY = $RawWriters[($i*6)+1].Trim().IndexOf("}") + 1 - $WriterIdX
            $InstanceIdX = $RawWriters[($i*6)+2].Trim().IndexOf("{")
            $InstanceIdY = $RawWriters[($i*6)+2].Trim().IndexOf("}") + 1 - $InstanceIdX
            $StateIdX = $RawWriters[($i*6)+3].Trim().IndexOf("[")
            $StateIdY = $RawWriters[($i*6)+3].Trim().IndexOf("]") + 1 - $StateIdX
            $StateDescX = $RawWriters[($i*6)+3].Trim().IndexOf("]") + 1
            $StateDescY = $RawWriters[($i*6)+3].Trim().Length - $StateDescX
            $LastErrorX = $RawWriters[($i*6)+4].Trim().IndexOf(":") + 1
            $LastErrorY = $RawWriters[($i*6)+4].Trim().Length - $LastErrorX

            $Writer = New-Object -TypeName PSObject
            $Writer | Add-Member -MemberType NoteProperty -Name "WriterName" -Value $RawWriters[($i*6)].Split("'")[1]
            $Writer | Add-Member -MemberType NoteProperty -Name "WriterId" -Value $RawWriters[($i*6)+1].Trim().SubString($WriterIdX, $WriterIdY)
            $Writer | Add-Member -MemberType NoteProperty -Name "WriterInstanceId" -Value $RawWriters[($i*6)+2].Trim().SubString($InstanceIdX,$InstanceIdY)
            $Writer | Add-Member -MemberType NoteProperty -Name "StateID" -Value $RawWriters[($i*6)+3].Trim().SubString($StateIdX,$StateIdY)
            $Writer | Add-Member -MemberType NoteProperty -Name "StateDescription" -Value $RawWriters[($i*6)+3].Trim().SubString($StateDescX,$StateDescY).Trim()
            $Writer | Add-Member -MemberType NoteProperty -Name "LastError" -Value $RawWriters[($i*6)+4].Trim().SubString($LastErrorX,$LastErrorY).Trim()
            $Writers += $Writer
        }
        $progressRate++
    }

    Write-Debug "Done"
    Write-Output $Writers
}