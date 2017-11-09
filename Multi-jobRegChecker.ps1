Set-StrictMode -Version 2.0

$objectsarray = @()
$counter = 0
$timecode = [Math]::Round((Get-Date).ToFileTime()/10000)
$MaxThreads = 5

$pcsearchbase = "OU=COMPUTERS,DC=SOMEDC,DC=org"
$outputcsvpath = "C:\Users\$([Environment]::UserName)\Desktop\RegVals$($timecode).csv"
$RegSubKey = "SOFTWARE\SomeCoolProgram\Lovely"
$RegGetValueOf = "SerialNumber"

Write-Host "Fetching a list of PC's in $pcsearchbase... Please wait."
$pcnames = Get-ADComputer -Filter "*" -SearchBase $pcsearchbase -Properties * | select-object -expandproperty name


foreach($name in $pcnames){
    Write-Progress -Activity "Checking $($RegSubKey)\$($RegGetValueOf) on PC $($counter) out of $($pcnames.Count)" -status "Working on $($name)"  -percentComplete (($counter / $pcnames.Count)*100)        
    

    While (@(Get-Job | Where { $_.State -eq "Running" }).Count -ge $MaxThreads){  
        Write-Host "Waiting for an open thread...($MaxThreads Maximum)"
        Start-Sleep -Seconds 3
    }
 
    $Scriptblock = {
        Param (
            [string]$name,
            [string]$rsk,
            [string]$rgvo
        )
        $TestObject = New-Object PSObject
        $errors = ""
        $RegVal = ""        
        
        if(Test-Connection -ComputerName $name -Quiet){
            $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $name)
            $RegKey= $Reg.OpenSubKey("$($rsk)")
            try{
                $RegVal = $RegKey.GetValue("$($rgvo)")
            }catch{
                $errors = "No registry key found"
            }       
        }else{
            $errors = "$name could not be contacted to test."
        }

        $TestObject | Add-Member -Type NoteProperty -Name ComputerName -Value $name
        $TestObject | Add-Member -Type NoteProperty -Name RegVal -Value $RegVal
        $TestObject | Add-Member -Type NoteProperty -Name Errors -Value $errors  
        $TestObject
    }

    Start-Job -Name "Job: $name" -ScriptBlock $Scriptblock -ArgumentList $name, $RegSubkey, $RegGetValueOf
    $counter++
}

While (@(Get-Job | Where { $_.State -eq "Running" }).Count -ne 0)
{  Write-Host "Waiting for background jobs..."
   Get-Job
   Start-Sleep -Seconds 3
}
 
Get-Job

$Data = ForEach ($Job in (Get-Job)) {
   $objectsarray += Receive-Job $Job
   Remove-Job $Job
}

foreach($xobject in $objectsarray){
    $xobject | Export-Csv -Append -Path $outputcsvpath -NoTypeInformation
}