$ErrorActionPreference = "Stop"

function Get-DateAndTime {

	#region .NET
	[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	$TimeForm = New-Object -TypeName System.Windows.Forms.Form
	[System.Windows.Forms.DateTimePicker]$StartTimePicker = $null
	[System.Windows.Forms.Label]$Start = $null
	[System.Windows.Forms.Label]$24Hour = $null
	[System.Windows.Forms.Label]$Occur = $null
	[System.Windows.Forms.Button]$TimeOk = $null
	[System.Windows.Forms.Button]$timecancel = $null
	[System.Windows.Forms.Button]$button1 = $null
	$StartTimePicker = New-Object -TypeName System.Windows.Forms.DateTimePicker
	$Start = New-Object -TypeName System.Windows.Forms.Label
	$24Hour = New-Object -TypeName System.Windows.Forms.Label
	$Occur = New-Object -TypeName System.Windows.Forms.Label
	$TimeOk = New-Object -TypeName System.Windows.Forms.Button
	$timecancel = New-Object -TypeName System.Windows.Forms.Button
	$TimeForm.SuspendLayout()
	#
	#StartTimePicker
	#
	$StartTimePicker.CustomFormat = 'MMMMdd, yyyy  |  HH:mm'
	$StartTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
	$StartTimePicker.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(150,58)
	$StartTimePicker.Name = 'StartTimePicker'
	$StartTimePicker.ShowUpDown = $true
	$StartTimePicker.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(280,20)
	$StartTimePicker.TabIndex = 0
	$StartTimePicker.Value = (get-date).AddHours(-4)
	$StartTimePicker.MaxDate = (get-date).AddHours(-1)
	#
	#Start
	#
	$Start.AutoSize = $true
	$Start.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(12,62)
	$Start.Name = 'Start'
	$Start.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(58,13)
	$Start.TabIndex = 2
	$Start.Text = 'Start Time:'
	#
	#24Hour
	#
	$24Hour.AutoSize = $true
	$24Hour.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(150,95)
	$24Hour.Name = '24Hour'
	$24Hour.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(138,13)
	$24Hour.TabIndex = 4
	$24Hour.Text = '** UTC and 24 hour format'
	#
	#Occur
	#
	$Occur.AutoSize = $true
	$Occur.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(12,20)
	$Occur.Name = 'Occur'
	$Occur.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(127,13)
	$Occur.TabIndex = 5
	$Occur.Text = 'When did your traffic occured?'
	#
	#TimeOk
	#
	$TimeOk.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(162,131)
	$TimeOk.Name = 'TimeOk'
	$TimeOk.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(96,28)
	$TimeOk.TabIndex = 6
	$TimeOk.Text = 'Ok'
	$TimeOk.UseVisualStyleBackColor = $true
	$TimeOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
	#
	#timecancel
	#
	$timecancel.Location = New-Object -TypeName System.Drawing.Point -ArgumentList @(15,131)
	$timecancel.Name = 'timecancel'
	$timecancel.Size = New-Object -TypeName System.Drawing.Size -ArgumentList @(99,28)
	$timecancel.TabIndex = 7
	$timecancel.Text = 'Cancel'
	$timecancel.UseVisualStyleBackColor = $true
	$timecancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	#
	#TimeForm
	#
	$TimeForm.ClientSize = New-Object -TypeName System.Drawing.Size -ArgumentList @(500,200)
	$TimeForm.Controls.Add($timecancel)
	$TimeForm.Controls.Add($TimeOk)
	$TimeForm.Controls.Add($Occur)
	$TimeForm.Controls.Add($24Hour)
	$TimeForm.Controls.Add($Start)
	$TimeForm.Controls.Add($StartTimePicker)
	$TimeForm.Name = 'TimeForm'
	$TimeForm.Text = 'Input the time'
	$TimeForm.ResumeLayout($false)
	$TimeForm.PerformLayout()
	Add-Member -InputObject $TimeForm -Name base -Value $base -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name StartTimePicker -Value $StartTimePicker -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name Start -Value $Start -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name 24Hour -Value $24Hour -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name Occur -Value $Occur -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name TimeOk -Value $TimeOk -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name timecancel -Value $timecancel -MemberType NoteProperty
	Add-Member -InputObject $TimeForm -Name button1 -Value $button1 -MemberType NoteProperty
	$TimeForm.Topmost = $True
	$TimeForm.StartPosition = "CenterScreen"
	$TimeForm.MaximizeBox = $false
	$TimeForm.FormBorderStyle = 'Fixed3D'
	$timeresult = $TimeForm.ShowDialog()
	#endregion .NET
	if ($timeresult -eq [System.Windows.Forms.DialogResult]::Cancel){
        $result = $null
    }
	if ($timeresult -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = $StartTimePicker.Value
    }
    return $result
}

# https://qiita.com/stofu/items/a8dc81c958d416684f2e
function Convert-UnixTimeToDateTime($unixTime){
    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $convertedDateTime = (Get-date $origin.AddSeconds($unixTime) -f yyyy/MM/dd-HH:mm:ss)
    return $convertedDateTime
}

$enabledNsg = New-Object System.Collections.ArrayList

if (!(Get-AzContext)){
    Write-Output "Please login Azure with Login-AzAccount"
}

# select the subscription which your NSG is in
$Subscription = Get-AzSubscription | Out-GridView -PassThru -Title "Please select your subscription which NSG is in"
$Subscription | Select-AzSubscription 

# Get all NSGs and all network watchers
$networkSecurityGroups = Get-AzNetworkSecurityGroup
$networkWatchers = Get-AzNetworkWatcher

# Extravt the NSG which FLowLogs was enabled
$networkSecurityGroups | ForEach-Object {
    $nsg = $_

    if ( $networkWatchers.location -eq $nsg.Location ){
        $targetNetworkWatcher = $networkWatchers | Where-Object { $_.Location -eq $nsg.Location }
        $result = Get-AzNetworkWatcherFlowLogStatus -TargetResourceId $nsg.Id -NetworkWatcher $targetNetworkWatcher

        if ($result.Enabled -eq $true){
            $enabledNsg.add($result) | Out-Null
        }
    }
}

# select the NSG which user want to confirm
$targetNsg = $enabledNsg.TargetResourceId | Out-GridView -PassThru -Title "Select the NSG"
$targetNsg -match "/subscriptions/(.*)/resourceGroups/(.*)/providers/Microsoft.Network/networkSecurityGroups/(.*)" | Out-Null
$prefixSubId = $Matches[1].ToUpper()
$prefixResourceGroup = $Matches[2].ToUpper()
$prefixNsg = $Matches[3].ToUpper()

# Create the prefix to download the target NSG FLowLogs
$logDate = Get-DateAndTime
Write-output "You inputted $logDate"

if ( $null -eq $logDate ){
    throw "You pushed cancel. You need to input the time and push OK"
}

$blogPrefix = "resourceId=/SUBSCRIPTIONS/$prefixSubId/RESOURCEGROUPS/$prefixResourceGroup/PROVIDERS/MICROSOFT.NETWORK/NETWORKSECURITYGROUPS/$prefixNsg/"
# ToDo: Add the method to convert from UTC to JST
#$utcLogDate = Get-date $logDate.AddHours(-9)
$utcLogDate = (Get-Date $logDate -Format yyyyMMddHH)
$blogPrefix += "y=$($utcLogDate.Substring(0, 4))/m=$($utcLogDate.Substring(4, 2))/d=$($utcLogDate.Substring(6, 2))/h=$($utcLogDate.Substring(8, 2))/m=00/"

# Download the PT1H.json as tmp-PT1H.json.
$enabledNsg | ForEach-Object {
    if ($_.TargetResourceId -eq $targetNsg){
        $targetStorageAccountId = $_.StorageId

        $targetStorageAccountId -match "resourceGroups/(.*)/providers/Microsoft.Storage/storageAccounts/(.*)" | Out-Null
        $targetStorageAccountResourceGroup = $Matches[1]
        $targetStorageAccountName = $Matches[2]

        $ctx = (Get-AzStorageAccount -Name $targetStorageAccountName -ResourceGroupName $targetStorageAccountResourceGroup).Context
        $blobs = Get-AzStorageBlob -Container "insights-logs-networksecuritygroupflowevent" -Context $ctx -Prefix $blogPrefix
        $blob = $blobs | Out-GridView -PassThru -Title "Select the file which you want to look"

        Get-AzStorageBlobContent -Context $ctx -Container "insights-logs-networksecuritygroupflowevent" `
            -Blob $blob.Name -Destination "./tmp-PT1H.json" -Force
    }
}

$protpcplConverter = @{
    "T" = "TCP"
    "U" = "UDP"
}

$trafficFlowConverter = @{
    "I" = "Inound"
    "O" = "Outbound"
}

$trafficDecisionConverter = @{
    "A" = "Allowed"
    "D" = "Deny"
}

# Convert from PT1H.json to readable records.
$plainFlowlogs = New-Object System.Collections.ArrayList
$records = (Get-content "./tmp-PT1H.json" -Raw | ConvertFrom-Json).records
$records | ForEach-Object {
    $record = $_

    $record.properties.flows | ForEach-Object{
        $flow1 = $_
        $flow1.flows | ForEach-Object {
            $flow2 = $_
            $flow2.flowTuples | ForEach-Object {
                $flowTuple = $_
                $flowTupleArray = $flowTuple.Split(",")

                $plainFlowLog  = New-Object PSCustomObject
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "rule" -Value $flow1.rule
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "mac" -Value $flow2.mac
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "TimeStamp" -Value "$(Convert-UnixTimeToDateTime($flowTupleArray[0]))"
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "SourceIP" -Value $flowTupleArray[1]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "DestinationIP" -Value $flowTupleArray[2]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "SourcePort" -Value $flowTupleArray[3]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "DestinationPort" -Value $flowTupleArray[4]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "Protocol" -Value $protpcplConverter[$flowTupleArray[5]]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "TrafficFlow" -Value $trafficFlowConverter[$flowTupleArray[6]]
                $plainFlowLog | Add-Member -MemberType NoteProperty -Name "TrafficDecision" -Value $trafficDecisionConverter[$flowTupleArray[7]]
                
                $plainFlowlogs.Add($plainFlowlog) | Out-Null
            }
        }
    }
}

# Open Grid-view 
$targetNsg -match "Microsoft.Network/networkSecurityGroups/(.*)" | Out-Null
$plainFlowlogs | Out-GridView -Title "FlogLogs of $($matches[1])"