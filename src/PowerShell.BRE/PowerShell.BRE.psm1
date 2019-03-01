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