[CmdletBinding()]
Param ()
Process {
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

    function Get-Vocabulary {
        [CmdletBinding()]
        Param (
            [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
            [string]$Name,
            [Parameter()]
            [version]$Version,
            [Parameter()]
            [string]$Output = "."
        )
        Process {
            $vocabs = $ruleStore.GetVocabularies($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
            if ($PSBoundParameters.ContainsKey("Version")) {
                $vocabs = $vocabs | Where-Object {($_.MajorRevision -eq $Version.Major) -and ($_.MinorRevision -eq $Version.Minor)}
            }
            Write-Verbose "Found $($vocabs.Count) vocabularies"
            return $vocabs
        }
    }
    
    function Export-Vocabulary {
        [CmdletBinding()]
        Param (
            [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
            [string]$Name,
            [Parameter()]
            [version]$Version,
            [Parameter()]
            [string]$Output = "."
        )
        Process {
            $vocabs = Get-Vocabulary @PSBoundParameters

            foreach ($v in $vocabs) {
                $fileName = "$($v.Name).$($v.MajorRevision).$($v.MinorRevision).xml"
                Write-Debug "FileName = $fileName"
                $filePath = Join-Path -Path $Output -ChildPath $fileName
                Write-Debug "FilePath = $filePath"
                Write-Output $filePath
                $driver.ExportVocabularyToFileRuleStore($v, $filePath)
            }
        }
    }
}