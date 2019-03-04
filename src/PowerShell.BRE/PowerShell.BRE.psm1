$driver = [Microsoft.BizTalk.RuleEngineExtensions.RuleSetDeploymentDriver]::new()
$ruleStore = $driver.GetRuleStore()
Write-Verbose "Connected to BRE store on $($ruleStore.Location)"

function Import-Policy {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateScript({$_.Exists})]
        [System.IO.FileInfo]$Path,
        [Parameter()]
        [switch]$Deploy
    )
    Process {
        [xml]$policy = Get-Content -Path $Path.FullName
        $policy

        $driver.ImportAndPublishFileRuleStore($Path.FullName)
        if ($Deploy) {
            Write-Verbose "Deploying Policy"
        }
    }
}

function Remove-Vocabulary {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Name,
        [Parameter(Position=1)]
        [version]$Version
    )
    Process {
        $vocabs = $ruleStore.GetVocabularies($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
        $vocabs = if ($PSBoundParameters.ContainsKey("Version")) {
            $vocabs = $vocabs | Where-Object {$_.MajorRevision -eq $Version.Major }
        }
    }
}