[CmdletBinding()]
param ()
process {
    $driver = [Microsoft.BizTalk.RuleEngineExtensions.RuleSetDeploymentDriver]::new()
    $ruleStore = $driver.GetRuleStore()
    Write-Verbose "Connected to BRE store on $($ruleStore.Location)"

    #region Rules
    function Export-Rule {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [Microsoft.RuleEngine.RuleSetInfo]$Rule,
            [Parameter()]
            [string]$Output = "."
        )
        process {
            $fileName = "$($Rule.Name).$($Rule.MajorRevision).$($Rule.MinorRevision).xml"
            Write-Debug "FileName = $fileName"
            $filePath = Join-Path -Path $Output -ChildPath $fileName
            Write-Debug "FilePath = $filePath"
            $driver.ExportRuleSetToFileRuleStore($Rule, $filePath)
            Write-Output $filePath
        }
    }

    function Get-Rule {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, ValueFromPipeline = $true)]
            [string]$Name,
            [Parameter()]
            [version]$Version
        )
        process {
            $rules = if ($PSBoundParameters.ContainsKey("Name")) {
                $ruleStore.GetRuleSets($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
            }
            else {
                $ruleStore.GetRuleSets([Microsoft.RuleEngine.RuleStore+Filter]::All)
            }
            Write-Verbose "Found $($rules.Count) rules"
            return $rules
        }
    }

    function Import-Rule {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [ValidateScript( {$_.Exists})]
            [System.IO.FileInfo]$Path,
            [Parameter()]
            [switch]$Deploy
        )
        process {
            [xml]$policy = Get-Content -Path $Path.FullName
            $policy
    
            $driver.ImportAndPublishFileRuleStore($Path.FullName)
            if ($Deploy) {
                Write-Verbose "Deploying Policy"
            }
        }
    }
    #endregion
    
    #region Vocabularies
    function Export-Vocabulary {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [Microsoft.RuleEngine.VocabularyInfo]$Vocabulary,
            [Parameter()]
            [string]$Output = "."
        )
        process {
            $fileName = "$($Vocabulary.Name).$($Vocabulary.MajorRevision).$($Vocabulary.MinorRevision).xml"
            Write-Debug "FileName = $fileName"
            $filePath = Join-Path -Path $Output -ChildPath $fileName
            Write-Debug "FilePath = $filePath"
            $driver.ExportVocabularyToFileRuleStore($Vocabulary, $filePath)
            Write-Output $filePath
        }
    }

    function Get-Vocabulary {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, ValueFromPipeline = $true)]
            [string]$Name,
            [Parameter()]
            [version]$Version
        )
        process {
            $vocabs = if ($PSBoundParameters.ContainsKey("Name")) {
                $ruleStore.GetVocabularies($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
            } 
            else {
                $ruleStore.GetVocabularies([Microsoft.RuleEngine.RuleStore+Filter]::All)
            }
            if ($PSBoundParameters.ContainsKey("Version")) {
                $vocabs = $vocabs | Where-Object {($_.MajorRevision -eq $Version.Major) -and ($_.MinorRevision -eq $Version.Minor)}
            }
            Write-Verbose "Found $($vocabs.Count) vocabularies"
            return $vocabs
        }
    }

    function Remove-Vocabulary {
        [CmdletBinding(SupportsShouldprocess = $true)]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [Microsoft.RuleEngine.VocabularyInfo]$Vocabulary
        )
        process {
            $dependantRules = $ruleStore.GetDependentRuleSets($Vocabulary)
            if ($dependantRules.Count -gt 0) {
                Write-Warning "Dependant rules found: $($dependantRules.Count)"
                Write-Debug ($dependantRules | Out-String)
            }
            $ruleStore.Remove($Vocabulary)
        }
    }
    #endregion
}