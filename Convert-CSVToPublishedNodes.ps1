﻿[CmdletBinding()]
Param(
    [Parameter()]
    [Alias("h")]
    [switch]$Help,

    [Parameter(Mandatory=$true)]
    [String]$InputFileName,

    [Parameter()]
    [string]$OutputFileName,

    [Parameter()]
    [string]$Delimiter=','
    )


if((-not $InputFileName) -or ($Help -eq $true))
{
    Write-Host "Usage: $($MyInvocation.MyCommand.Name) [OPTIONS]"
    Write-Host "OPTIONS:"
    Write-Host "   -h, Help       : (Optional)  Display this screen."
    Write-Host "   -InputFileName : (Mandatory) Path to the input file. Needs to be in CSV format."
    Write-Host "   -OutputFileName: (Optional)  Path to the output file. Default: Same name as input file, with JSON extension."
    Write-Host "   -Delimiter     : (Optional)  Column delimiter in CSV file. Default: ','"
    Write-Host ""
    Write-Host "* InputFileName includes a line for each distinct OPC Node (OpcNodes_Id or OpcNodes_ExpandedNodeId)."
    Write-Host "* Only mandatory columns required to exist in the file."
    Write-Host "* Server related properties (ones that don't start with 'OPCNodes_') repeat for each OPC Node on that server. "
    Write-Host "* If server related properties (ones that don't start with 'OPCNodes_') are null for a given line, last known properties from previous lines are used."
    Write-Host "* InputFileName column structure is as follows:"
    Write-Host "    EndpointUrl                   : ([mandatory], [string] ) URL of OPC UA Server in format ""opc.tcp://<your_opcua_server>:<your_opcua_server_port>/<your_opcua_server_path>""."
    Write-Host "    UseSecurity                   : ([optional] , [boolean]) Allows to access the endpoint with SecurityPolicy.None when set to 'false' (no signing and encryption applied to the OPC UA communication), default is true"
    Write-Host "    OpcAuthenticationMode         : ([optional] , [string] ) ""Anonymous"" or ""UsernamePassword"", default is ""Anonymous"""
    Write-Host "    OpcAuthenticationUsername     : ([optional] , [string] ) Valid only if ""OpcAuthenticationMode"": ""UsernamePassword"""
    Write-Host "    OpcAuthenticationPassword     : ([optional] , [string] ) Valid only if ""OpcAuthenticationMode"": ""UsernamePassword"""
    Write-Host "    OpcNodes_Id                   : ([mandatory], [string] ) OPC node to publish in either NodeId format (contains ""ns="", e.g. ""ns=3;i=1234"") or ExpandedNodeId format (contains ""nsu="", e.g. ""nsu=http://mycompany.com/UA/Data;i=1234"")."
    Write-Host "                                                             Only one of ""Id"" or ""ExpandedNodeID"" is mandatory."
    Write-Host "    OpcNodes_ExpandedNodeId       : ([mandatory], [string] ) OPC node to publish in either NodeId format (contains ""ns="", e.g. ""ns=3;i=1234"") or ExpandedNodeId format (contains ""nsu="", e.g. ""nsu=http://mycompany.com/UA/Data;i=1234"")."
    Write-Host "                                                             Included for backward compatibility"
    Write-Host "    OpcNodes_OpcSamplingInterval  : ([optional] , [int]    ) Sampling interval OPC Publisher requests the server to sample the node value. The value is in milliseconds."
    Write-Host "    OpcNodes_OpcPublishingInterval: ([optional] , [int]    ) Subscription will publish node value with this interval, it will only be published if the value has changed. The value is in milliseconds."
    Write-Host "    OpcNodes_DisplayName          : ([optional] , [string] ) Display Name for Node. This value overrides DisplayName values fetched from server with -fd=true switch."
    Write-Host "    OpcNodes_HeartbeatInterval    : ([optional] , [int]    ) If set, the last value will be sent again with an updated SourceTimestamp value after the given interval.The value is in milliseconds. "
    Write-Host "    OpcNodes_SkipFirst            : ([optional] , [boolean]) When true, first event will not generate a telemetry event, this is useful when publishing a large amount of data to prevent a event flood at startup of OPC Publisher"
    Exit 0
}


$global:stopWatch=[System.Diagnostics.Stopwatch]::StartNew()
$global:execStats=[System.Int64[]]::new(100)
$global:debug=$false

function MeasureStopWatch()
{
Param ([int]$index)
if ($global:debug -eq $true)
{
    if ($index -gt -1){$global:execStats[$index]+=$global:stopWatch.ElapsedTicks}
    $global:stopWatch.Restart()
}
}

function FormatDuration()
{
Param ([TimeSpan]$TimeSpan)
$TimeSpanStr = ""
if($TimeSpan.Days>0) {$TimeSpanStr = [String]::Format("{0} days ",$TimeSpan.Days)}
$TimeSpanStr = [String]::Format("{0:hh}:{0:mm}:{0:ss}",$TimeSpan)
return $TimeSpanStr
}




if (-not $(Test-Path $InputFileName -PathType Leaf))
{
    Write-Error "File '$InputFileName' not found" -Category ObjectNotFound
    Exit 1
}

$InputFileName = (Get-Item $InputFileName).FullName
if (-not $OutputFileName) {
    $OutputFileName = [System.IO.Path]::GetFileNameWithoutExtension($(Split-Path $InputFileName -Leaf))
    $destPath = Split-Path -Path $InputFileName
    $OutputFileName = Join-Path $destPath ($OutputFileName + ".json")
}

Write-Host "PARAMETERS:"
Write-Host "    InputFileName : $InputFileName"
Write-Host "    OutputFileName: $OutputFileName"
Write-Host "    Delimiter     : $Delimiter"
Write-Host ""

$nodeListInput = Get-Content $InputFileName | ConvertFrom-Csv -Delimiter $Delimiter
$nodeListOutput = [System.Collections.ArrayList][ordered]@{}
$currentServerNode=$null
$lineCount=$nodeListInput.Count

$nodeIdList = New-Object system.collections.hashtable #case sensitive hashtables
#$nodeIdList = @{} #case insensitive hashtables
$displayNameList = New-Object system.collections.hashtable #case sensitive hashtables
#$displayNameList = @{} #case insensitive hashtables

$prevEndpointUrl=$null
$prevUseSecurity=$null
$prevOpcAuthenticationMode=$null
$prevUsername=$null
$prevPassword=$null

$expandedNodeIdExists=$nodeListInput | where {-not [string]::IsNullOrWhiteSpace($_.OpcNodes_ExpandedNodeId)}

$timeStarted= Get-Date

$lineNum=0

foreach ($node in $nodeListInput) {

    MeasureStopWatch -index -1

    $lineNum=$lineNum+1
    if(-not $node.EndpointUrl)
    {
        $node.EndpointUrl=$prevEndpointUrl
        if(-not $node.UseSecurity){$node.UseSecurity=$prevUseSecurity}
        if(-not $node.OpcAuthenticationMode){$node.OpcAuthenticationMode=$prevOpcAuthenticationMode}
        if(-not $node.OpcAuthenticationUsername){$node.OpcAuthenticationUsername=$prevUsername}
        if(-not $node.OpcAuthenticationPassword){$node.OpcAuthenticationPassword=$prevPassword}
    }

    if($node.EndpointUrl)
    {
        $prevEndpointUrl=$node.EndpointUrl
        $prevUseSecurity=if(-not $node.UseSecurity){$null}else{$node.UseSecurity}
        $prevOpcAuthenticationMode=if(-not $node.OpcAuthenticationMode){$null}else{$node.OpcAuthenticationMode}
        $prevUsername=if(-not $node.OpcAuthenticationUsername){$null}else{$node.OpcAuthenticationUsername}
        $prevPassword=if(-not $node.OpcAuthenticationPassword){$null}else{$node.OpcAuthenticationPassword}

        MeasureStopWatch -index 0
         
        $currentServerNode=$nodeListOutput | `
                .{process{If(($_.EndpointUrl -eq $node.EndpointUrl) -and `
                       (($_.UseSecurity -eq $node.UseSecurity) -or ([string]::IsNullOrWhiteSpace($_.UseSecurity) -and [string]::IsNullOrWhiteSpace($node.UseSecurity))) -and `
                       (($_.OpcAuthenticationMode -eq $node.OpcAuthenticationMode) -or ([string]::IsNullOrWhiteSpace($_.OpcAuthenticationMode) -and [string]::IsNullOrWhiteSpace($node.OpcAuthenticationMode))) -and `
                       (($_.OpcAuthenticationUsername -eq $node.OpcAuthenticationUsername) -or ([string]::IsNullOrWhiteSpace($_.OpcAuthenticationUsername) -and [string]::IsNullOrWhiteSpace($node.OpcAuthenticationUsername))) -and `
                       (($_.OpcAuthenticationPassword -eq $node.OpcAuthenticationPassword) -or ([string]::IsNullOrWhiteSpace($_.OpcAuthenticationPassword) -and [string]::IsNullOrWhiteSpace($node.OpcAuthenticationPassword)))){$_}}}

        MeasureStopWatch -index 1
        if (-not $currentServerNode)
        {

            if ($nodeListOutput | .{process{If($_.EndpointUrl -eq $node.EndpointUrl){$_}}}) {Write-Host "[warn ] Line $lineNum : EndpointUrl '$($node.EndpointUrl)' appears more than once with different 'UseSecurity', 'OpcAuthenticationMode', 'OpcAuthenticationUsername' or 'OpcAuthenticationPassword' settings."  -ForegroundColor Yellow}

            $currentServerNode=[ordered]@{'EndpointUrl'=$node.EndpointUrl}
            if ($node.UseSecurity) {$currentServerNode.UseSecurity=[System.Convert]::ToBoolean($node.UseSecurity)}
            if ($node.OpcAuthenticationMode) 
            {
                if ($node.OpcAuthenticationMode -in "UsernamePassword", "Anonymous"){$currentServerNode.OpcAuthenticationMode=$node.OpcAuthenticationMode}
                else {Write-Host "[warn ] Line $lineNum : Invalid OpcAuthenticationMode '$($node.OpcAuthenticationMode)'" -ForegroundColor Red}
            }

            if ($node.OpcAuthenticationUsername) {$currentServerNode.OpcAuthenticationUsername=$node.OpcAuthenticationUsername}
            if ($node.OpcAuthenticationPassword) {$currentServerNode.OpcAuthenticationPassword=$node.OpcAuthenticationPassword}

            $dataPointNodes=[System.Collections.ArrayList]@{}
            $currentServerNode.OpcNodes=$dataPointNodes
            [void]$nodeListOutput.Add($currentServerNode)
        }
        MeasureStopWatch -index 2
        
        if ((([string]$node.OpcAuthenticationMode) -eq "Anonymous") -and (($node.OpcAuthenticationUsername) -or ($node.OpcAuthenticationPassword))){Write-Host "[warn ] Line $lineNum : 'OpcAuthenticationUsername' and 'OpcAuthenticationPassword' settings are not used with 'Anonymous' authentication (EndpointUrl : '$($node.EndpointUrl)')" -ForegroundColor Yellow}

        if(($node.OpcNodes_Id) -or ($node.OpcNodes_ExpandedNodeId))
        {
            $existingNode=$nodeIdList[$node.OpcNodes_Id]
            if((-not $existingNode) -and ($expandedNodeIdExists)) {$nodeIdList[$node.$node.OpcNodes_ExpandedNodeId]}
        MeasureStopWatch -index 3

            if(-not $existingNode)
            {
                if($node.OpcNodes_Id)
                {
                    $dataPointNode=[ordered]@{'Id'=$node.OpcNodes_Id}
                }
                elseif($node.OpcNodes_ExpandedNodeId)
                {
                    $dataPointNode=[ordered]@{'Id'=$node.OpcNodes_ExpandedNodeId}
                }

                if ($node.OpcNodes_OpcSamplingInterval){$dataPointNode.OpcSamplingInterval=[int]$node.OpcNodes_OpcSamplingInterval}
                if ($node.OpcNodes_OpcPublishingInterval){$dataPointNode.OpcPublishingInterval=[int]$node.OpcNodes_OpcPublishingInterval}
                if ($node.OpcNodes_DisplayName){$dataPointNode.DisplayName=$node.OpcNodes_DisplayName}
                if ($node.OpcNodes_HeartbeatInterval){$dataPointNode.HeartbeatInterval=[int]$node.OpcNodes_HeartbeatInterval}
                if ($node.OpcNodes_SkipFirst){$dataPointNode.SkipFirst=[System.Convert]::ToBoolean($node.OpcNodes_SkipFirst)}

                if(($dataPointNode.OpcPublishingInterval) -and ($dataPointNode.OpcSamplingInterval) -and ($dataPointNode.OpcPublishingInterval -lt $dataPointNode.OpcSamplingInterval))
                {
                    Write-Host "[warn ] Line $lineNum : There's no point in publishing a value more frequently than it is sampled. EndpointUrl=$($currentServerNode.EndpointUrl), NodeId=$($node.OpcNodes_Id)$($node.OpcNodes_ExpandedNodeId), PublishingInterval=$($dataPointNode.OpcPublishingInterval), SamplingInterval=$($dataPointNode.OpcSamplingInterval)"  -ForegroundColor Yellow
                }

        MeasureStopWatch -index 4
                if (($node.OpcNodes_DisplayName) -and $displayNameList.ContainsKey($node.OpcNodes_DisplayName))
                {
                    Write-Host "[warn ] Line $lineNum : 'DisplayName'='$($node.OpcNodes_DisplayName)' appears more than once under 'EndpointUrl'='$($node.EndpointUrl)'." -ForegroundColor Yellow
                }

        MeasureStopWatch -index 5
                [void]$currentServerNode.OpcNodes.Add($dataPointNode)
                [void]$nodeIdList.Add($dataPointNode.Id,$dataPointNode)
                [void]$displayNameList.Add($dataPointNode.DisplayName,$dataPointNode)
        MeasureStopWatch -index 6
            }
            else
            {
                Write-Host "[error] Line $lineNum : Duplicate NodeId $($node.OpcNodes_Id)$($node.OpcNodes_ExpandedNodeId) under 'EndpointUrl'='$($node.EndpointUrl)'" -ForegroundColor Red
            }
        }
        else
        {
            Write-Host "[error] Line $lineNum : Id and ExpandedNodeId are both empty." -ForegroundColor Red
        }

    }
    else
    {
        Write-Host "[error] Line $lineNum : EndpointUrl value is empty." -ForegroundColor Red
    }

    #if total lines >1000 then update status every 100 records
    if (($lineCount -le 1000) -or (($lineCount -gt 1000) -and ($lineNum%100 -eq 0)))
    {
        $completeratio = [double](($lineNum)/[double]$lineCount)
        $pct =[int] ($completeratio*100)
        $timeElapsed = New-TimeSpan -Start $timeStarted -End $(Get-Date)
        $totalTime = New-TimeSpan -Seconds ([Int64]([double]$timeElapsed.TotalSeconds * [double]$lineCount / [double]$lineNum))
        $timeRemaining=$totalTime - $timeElapsed

        $timeRemainingStr = FormatDuration -TimeSpan $timeRemaining
        $timeElapsedStr = FormatDuration -TimeSpan $timeElapsed
        $totalTimeStr = FormatDuration -TimeSpan $totalTime

        $statusMsg=[String]::Format("{0}% ({1}/{2}) Complete. Time Elapsed:{3} Time Remaining:{4} Total Time:{5}", $pct,$lineNum,$lineCount,$timeElapsedStr, $timeRemainingStr, $totalTimeStr)
        Write-Progress -Activity "Processing ..." -Status $statusMsg -PercentComplete $pct

        MeasureStopWatch -index 7
    }
    
    #if ($global:debug -eq $true){if ($lineNum -gt 30000){break}}
}

Write-Output "Writing to file: $OutputFileName..."
$jsonTop = [System.Collections.ArrayList][ordered]@{}
[void]$jsonTop.Add($nodeListOutput)
$jsonTop | ConvertTo-Json -depth 100 | Out-File $OutputFileName 


$timeElapsed = New-TimeSpan -Start $timeStarted -End $(Get-Date)
$statusMsg=[String]::Format("Finished. Total time: {0:hh}:{0:mm}:{0:ss}", $timeElapsed)
Write-Output $statusMsg

if ($global:debug -eq $true)
{
    For($i=0; $i -le 7; $i++) {Write-Host "$i : $($global:execStats[$i])"}
}


