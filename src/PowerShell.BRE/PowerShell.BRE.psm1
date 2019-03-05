[CmdletBinding()]
param ()
process {
    $driver = [Microsoft.BizTalk.RuleEngineExtensions.RuleSetDeploymentDriver]::new()
    $ruleStore = $driver.GetRuleStore()
    Write-Verbose "Connected to BRE store on $($ruleStore.Location)"

    #region Policies
    function Export-Policy {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [Microsoft.RuleEngine.RuleSetInfo]$Policy,
            [Parameter()]
            [string]$Output = "."
        )
        process {
            $fileName = "$($Policy.Name).$($Policy.MajorRevision).$($Policy.MinorRevision).xml"
            Write-Debug "FileName = $fileName"
            $filePath = Join-Path -Path $Output -ChildPath $fileName
            Write-Debug "FilePath = $filePath"
            $driver.ExportRuleSetToFileRuleStore($Rule, $filePath)
            Write-Output $filePath
        }
    }

    function Get-Policy {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, ValueFromPipeline = $true)]
            [string]$Name,
            [Parameter()]
            [version]$Version
        )
        process {
            if (-not $PSBoundParameters.ContainsKey("Version")) {
                Write-Verbose "Looking for policy: $Name"
            }
            else {
                Write-Verbose "Looking for policy: $Name v$($version.ToString())"
            }
            $policies = if ($PSBoundParameters.ContainsKey("Name")) {
                $ruleStore.GetRuleSets($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
            }
            else {
                $ruleStore.GetRuleSets([Microsoft.RuleEngine.RuleStore+Filter]::All)
            }
            if ($PSBoundParameters.ContainsKey("Version")) {
                $policies = $policies | Where-Object {($_.MajorRevision -eq $Version.Major) -and ($_.MinorRevision -eq $Version.Minor)}
            }
            Write-Verbose "Found $($policies.Count) policy(s)"
            return $policies
        }
    }

    function Import-Policy {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [ValidateScript({$_.Exists})]
            [System.IO.FileInfo]$Path,
            [Parameter()]
            [switch]$Deploy,
            [Parameter()]
            [switch]$Force
        )
        process {
            Write-Verbose "Reading XML"
            [xml]$policyXml = Get-Content -Path $Path.FullName
            [System.Collections.Generic.List[Microsoft.RuleEngine.RuleSetInfo]]$policies = [System.Collections.Generic.List[Microsoft.RuleEngine.RuleSetInfo]]::new()
            foreach ($p in $policyXml.brl.ruleset) {
                $policies.Add([Microsoft.RuleEngine.RuleSetInfo]::new($p.name, $p.version.major, $p.version.minor))

                $policy = Get-Policy -Name $p.name -Version ([version]::new($p.version.major, $p.version.minor))
                if ($policy) {
                    Write-Warning "Policy already deployed"
                    Write-Debug ($policy | Out-String)
                    
                    if ($Force) {
                        Remove-Policy -Policy $policy -Delete
                    }
                }
            }

            Write-Verbose "Publishing XML policy(s)"
            $driver.ImportAndPublishFileRuleStore($Path.FullName)
            if ($Deploy) {
                foreach ($p in $policies) {
                    Write-Verbose "Deploying policy: $($p.Name) v$($p.MajorRevision).$($p.MinorRevision)"
                    Write-Debug ($p | Out-String)
                    $driver.Deploy($p)
                }
            }
        }
    }

    function Remove-Policy {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [Microsoft.RuleEngine.RuleSetInfo]$Policy,
            [Parameter()]
            [switch]$Delete
        )
        process {
            Write-Verbose "Removing policy: $($Policy.Name) v$($Policy.MajorRevision).$($Policy.MinorRevision)"
            if ($driver.IsRuleSetDeployed($Policy)) {
                if ($PSCmdlet.ShouldProcess($Policy, "Undeploying policy")) {
                    $driver.Undeploy($Policy)
                    Write-Verbose "Undeployed policy"
                }
            }
            if ($Delete) {
                if ($PSCmdlet.ShouldProcess($Policy, "Deleting policy")) {
                    $ruleStore.Remove($Policy)
                    Write-Verbose "Deleted policy"
                }
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

    function Import-Vocabulary {
        [CmdletBinding()]
        param (
            [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
            [ValidateScript({$_.Exists})]
            [System.IO.FileInfo]$Path,
            [Parameter()]
            [switch]$Force
        )
        process {
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