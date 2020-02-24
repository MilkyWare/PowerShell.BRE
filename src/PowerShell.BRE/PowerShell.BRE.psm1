#Requires -Assembly 'Microsoft.RuleEngine'
#Requires -Assembly 'Microsoft.BizTalk.RuleEngineExtensions'

$driver = [Microsoft.BizTalk.RuleEngineExtensions.RuleSetDeploymentDriver]::new()
$ruleStore = $driver.GetRuleStore()
Write-Verbose "Connected to BRE store on $($ruleStore.Location)"

#region Policies
function Clear-Policies
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param ()
    process
    {
        Get-Policy | Remove-Policy -Delete
    }
}

<#
    .SYNOPSIS
        Exports specified policy to file. The policies can be specified explicity or passed via the pipeline. These will be exported by default too the current folder, be can be changed and will be exported in the format Guid_Name.MajorVersion.MinorVersion.xml. FileInfo objects are returned for each exported policy.
    .EXAMPLE
        PS C:\> Export-Policy -Policy $policy
        Exports policy in variable to current folder
    .EXAMPLE
        PS C:\> Export-Policy -Policy $policy -OutPut D:\Temp
        Exports policy in variable to specified folder
    .EXAMPLE
        PS C:\> Get-Policy | Export-Policy
        Gets all BRE policies and exports each to the current folder
    .PARAMETER Policy
        Policy to be exported
    .PARAMETER Output
        Directory for the policy to be exported to. Default is current folder
    #>
function Export-Policy
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.RuleEngine.RuleSetInfo]$Policy,
        [Parameter()]
        [System.IO.DirectoryInfo]$Output = "."
    )
    process
    {
        $fileName = "$(New-Guid)_$($Policy.Name).$($Policy.MajorRevision).$($Policy.MinorRevision).xml"
        Write-Debug "FileName = $fileName"
        $filePath = Join-Path -Path $Output -ChildPath $fileName

        if (-not $Output.Exists)
        {
            Write-Verbose "Creating output directory"
            New-Item -Path $Output.FullName -ItemType Directory -Force
        }

        Write-Debug "FilePath = $filePath"
        $driver.ExportRuleSetToFileRuleStore($Policy, $filePath)
        Write-Output ([System.IO.FileInfo]::new($filePath))
    }
}

<#
    .SYNOPSIS
        Searches BRE for policies matching specified parameters. Calling with no paramters will return all BRE polcies, where specifying the name and/or version will filter this list
    .EXAMPLE
        PS C:\> Get-Policy
        Returns all BRE policies
    .EXAMPLE
        PS C:\> Get-Policy -Name Test
        Returns all versions of the BRE policy "Test"
    .EXAMPLE
        PS C:\> Get-Policy -Name -Version 1.0
        Returns version "1.0" of the BRE policy "Test"
    .PARAMETER Name
        Name of the BRE policy to filter on
    .PARAMETER Version
        Version of the BRE policy to filter on
    #>
function Get-Policy
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [string]$Name,
        [Parameter()]
        [version]$Version
    )
    process
    {
        if (-not $PSBoundParameters.ContainsKey("Version"))
        {
            Write-Verbose "Looking for policy: $Name"
        }
        else
        {
            Write-Verbose "Looking for policy: $Name v$($version.ToString())"
        }
        $policies = if ($PSBoundParameters.ContainsKey("Name"))
        {
            $ruleStore.GetRuleSets($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
        }
        else
        {
            $ruleStore.GetRuleSets([Microsoft.RuleEngine.RuleStore+Filter]::All)
        }
        if ($PSBoundParameters.ContainsKey("Version"))
        {
            $policies = $policies | Where-Object { ($_.MajorRevision -eq $Version.Major) -and ($_.MinorRevision -eq $Version.Minor) }
        }
        Write-Verbose "Found $($policies.Count) policy(s)"
        return $policies
    }
}

<#
    .SYNOPSIS
        Imports and publishes XML BRE policy into the store and optionally deploys the policy. The source XML files can also be cleaned up (removed) once the policy has been imported. Force can also be specified to remove an existing policy of the same name and version as part of the process
    .EXAMPLE
        PS C:\> Import-Policy -Path C:\BREPolicies\Test.1.0.xml
        Imports and publishes the policy "Test" version "1.0" into the BRE store
    .EXAMPLE
        PS C:\> Import-Policy -Path C:\BREPolicies\Test.1.0.xml -Deploy
        Imports, publishes and deploys the policy "Test" version "1.0" into the BRE store
    .EXAMPLE
        PS C:\> Import-Policy -Path C:\BREPolicies\Test.1.0.xml -Deploy -Force
        Imports, publishes and deploys the policy "Test" version "1.0" into the BRE store. If the policy already exists, it is removed to allow the new policy to be imported
    .EXAMPLE
        PS C:\> Import-Policy -Path C:\BREPolicies\Test.1.0.xml -CleanUp
        Imports and publishes the policy "Test" version "1.0" into the BRE store. Once imported, the XML file is removed
    .PARAMETER Path
        Path to the policy XML to be imported
    .PARAMETER Deploy
        Publish and Deploy the policies
    .PARAMETER Force
        Toggle for whether existing policies are to be handled during the import
    .PARAMETER CleanUp
        Toggle whether the source XML is deleted once the import process is complete
    #>
function Import-Policy
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript( { $_.Exists })]
        [System.IO.FileInfo]$Path,
        [Parameter()]
        [switch]$Deploy,
        [Parameter()]
        [switch]$Force,
        [Parameter()]
        [switch]$CleanUp
    )
    process
    {
        Write-Verbose "Reading XML"
        [xml]$policyXml = Get-Content -Path $Path.FullName
        [System.Collections.Generic.List[Microsoft.RuleEngine.RuleSetInfo]]$policies = [System.Collections.Generic.List[Microsoft.RuleEngine.RuleSetInfo]]::new()
        foreach ($p in $policyXml.brl.ruleset)
        {
            Write-Verbose "Processing policy: $($p.name) $($p.version.major).$($p.version.minor)"
            $policies.Add([Microsoft.RuleEngine.RuleSetInfo]::new($p.name, $p.version.major, $p.version.minor))

            $policy = Get-Policy -Name $p.name -Version ([version]::new($p.version.major, $p.version.minor))
            if ($policy)
            {
                Write-Warning "Policy already deployed"
                Write-Debug ($policy | Out-String)
                    
                if ($Force)
                {
                    Remove-Policy -Policy $policy -Delete
                }
            }
        }

        Write-Verbose "Publishing XML policy(s)"
        $driver.ImportAndPublishFileRuleStore($Path.FullName)
        if ($Deploy)
        {
            foreach ($p in $policies)
            {
                Write-Verbose "Deploying policy: $($p.Name) v$($p.MajorRevision).$($p.MinorRevision)"
                Write-Debug ($p | Out-String)
                $driver.Deploy($p)
            }
        }

        if ($CleanUp)
        {
            Write-Verbose "Removing XML"
            Remove-Item -Path $Path.FullName -Force
        }
    }
}

<#
    .SYNOPSIS
        Removed the specified policy from the BRE store. Policy can either be specified explicitly and passed from a pipeline. The default behaviour is just to undeploy the policy, it can also optionally be deleted from the store entirely
    .EXAMPLE
        PS C:\> Remove-Policy -Policy $policy
        Undeploys the specified policy from BRE
    .EXAMPLE
        PS C:\> Remove-Policy -Policy $policy -Delete
        Undeploys and deletes the specified policy from BRE
    .EXAMPLE
        PS C:\> Get-Policy | Remove-Policy
        Undeploys policies from the pipeline
    .PARAMETER Policy
        Policy to be removed
    .PARAMETER Delete
        Use to delete the policy from BRE instead of just undeploying
    #>
function Remove-Policy
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.RuleEngine.RuleSetInfo]$Policy,
        [Parameter()]
        [switch]$Delete
    )
    process
    {
        Write-Verbose "Removing policy: $($Policy.Name) v$($Policy.MajorRevision).$($Policy.MinorRevision)"
        if ($driver.IsRuleSetDeployed($Policy))
        {
            if ($PSCmdlet.ShouldProcess($Policy, "Undeploying policy"))
            {
                $driver.Undeploy($Policy)
                Write-Verbose "Undeployed policy"
            }
        }
        if ($Delete)
        {
            if ($PSCmdlet.ShouldProcess(($Policy | Out-String), "Deleting policy"))
            {
                $ruleStore.Remove($Policy)
                Write-Verbose "Deleted policy"
            }
        }
    }
}
#endregion
    
#region Vocabularies
function Clear-Vocabularies
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param ()        
    process
    {
        if ($PSCmdlet.ShouldProcess())
        {
            Get-Vocabulary | Remove-Vocabulary -Force
        }
    }
}

<#
    .SYNOPSIS
        Exports specified vocabulary to file. The vocabularies can be specified explicity or passed via the pipeline. These will be exported by default too the current folder, be can be changed and will be exported in the format Guid_Name.MajorVersion.MinorVersion.xml. FileInfo objects are returned for each exported vocabulary.
    .EXAMPLE
        PS C:\> Export-Vocabulary -Vocabulary $vocabulary
        Exports vocabulary in variable to current folder
    .EXAMPLE
        PS C:\> Export-Vocabulary -Vocabulary $vocabulary -OutPut D:\Temp
        Exports vocabulary in variable to specified folder
    .EXAMPLE
        PS C:\> Get-Vocabulary | Export-Vocabulary
        Gets all BRE vocabularies and exports each to the current folder
    .PARAMETER Vocabulary
        Vocabulary to be exported
    .PARAMETER Output
        Directory for the vocabulary to be exported to. Default is current folder
    #>
function Export-Vocabulary
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.RuleEngine.VocabularyInfo]$Vocabulary,
        [Parameter()]
        [System.IO.DirectoryInfo]$Output = "."
    )
    process
    {
        $fileName = "$(New-Guid)_$($Vocabulary.Name).$($Vocabulary.MajorRevision).$($Vocabulary.MinorRevision).xml"
        Write-Debug "FileName = $fileName"
        $filePath = Join-Path -Path $Output -ChildPath $fileName

        if (-not $Output.Exists)
        {
            Write-Verbose "Creating output directory"
            New-Item -Path $Output.FullName -ItemType Directory -Force
        }

        Write-Debug "FilePath = $filePath"
        $driver.ExportVocabularyToFileRuleStore($Vocabulary, $filePath)
        Write-Output ([System.IO.FileInfo]::new($filePath))
    }
}

<#
    .SYNOPSIS
        Searches BRE for vocabularies matching specified parameters. Calling with no paramters will return all BRE vocaularies, where specifying the name and/or version will filter this list
    .EXAMPLE
        PS C:\> Get-Vocabulary
        Returns all BRE vocabularies
    .EXAMPLE
        PS C:\> Get-Vocabulary -Name Test
        Returns all versions of the BRE vocabulary "Test"
    .EXAMPLE
        PS C:\> Get-Vocabulary -Name -Version 1.0
        Returns version "1.0" of the BRE vocabulary "Test"
    .PARAMETER Name
        Name of the BRE vocabulary to filter on
    .PARAMETER Version
        Version of the BRE vocabulary to filter on
    #>
function Get-Vocabulary
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [string]$Name,
        [Parameter()]
        [version]$Version
    )
    process
    {
        $vocabs = if ($PSBoundParameters.ContainsKey("Name"))
        {
            $ruleStore.GetVocabularies($Name, [Microsoft.RuleEngine.RuleStore+Filter]::All)
        } 
        else
        {
            $ruleStore.GetVocabularies([Microsoft.RuleEngine.RuleStore+Filter]::All)
        }
        if ($PSBoundParameters.ContainsKey("Version"))
        {
            $vocabs = $vocabs | Where-Object { ($_.MajorRevision -eq $Version.Major) -and ($_.MinorRevision -eq $Version.Minor) }
        }
        Write-Verbose "Found $($vocabs.Count) vocabularies"
        return $vocabs
    }
}

<#
    .SYNOPSIS
        Imports an exported BRE vocabulary whilst handling pre-exisitng vocabularies as well as policies that reference those. 
        
        A list of policies is taken from the XML and used to query the rule store. If vocabularies already exist, a list of dependant policies is retrieved and exported before deleting. Once the dependencies are removed, the XML vocabularies are imported and the dependencies restored.
    .EXAMPLE
        PS C:\> Import-Vocabulary -Path C:\Temp\0d54dc5b-e73e-4936-a751-6df7fb5f39f5_Vocab1.1.0.xml
        Imports specified BRE vocabulary XML
    .EXAMPLE
        PS C:\> Import-Vocabulary -Path C:\BREPolicies\Test.1.0.xml -CleanUp
        Imports and publishes the vocabulary "Test" version "1.0" into the BRE store. Once imported, the XML file is removed
    .EXAMPLE 
        PS C:\> Import-Vocabulary -Path C:\BREPolicies\Test.1.0.xml -Force
        Imports and publishes the vocabulary "Test" version "1.0" into the BRE store. Each vocabulary in the XML is looked up in the store. If the vocabulary already exists the dependencies are backed up and restored once the vocabularies have been imported
    .PARAMETER Path
        Path to the vocabulary XML to be imported
    .PARAMETER Force
        Toggle for whether dependent policies are to be handled during the import
    .PARAMETER CleanUp
        Toggle whether the source XML is deleted once the import process is complete
    #>
function Import-Vocabulary
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript( { $_.Exists })]
        [System.IO.FileInfo]$Path,
        [Parameter()]
        [switch]$Force,
        [Parameter()]
        [switch]$CleanUp
    )
    process
    {
        Write-Verbose "Reading XML"
        [xml]$vocabXml = Get-Content -Path $Path.FullName
        [System.Collections.Generic.List[Microsoft.RuleEngine.VocabularyInfo]]$vocabs = [System.Collections.Generic.List[Microsoft.RuleEngine.VocabularyInfo]]::new()
        [System.Collections.Generic.Stack[string]]$dependantPolicies = [System.Collections.Generic.List[string]]::new()
        foreach ($v in $vocabXml.brl.vocabulary)
        {
            Write-Verbose "Processing vocabulary: $($v.name) $($v.version.major).$($v.version.minor)"
            $vocabs.Add([Microsoft.RuleEngine.VocabularyInfo]::new($v.name, $v.version.major, $v.version.minor))

            $vocab = Get-Vocabulary -Name $v.name -Version ([version]::new($v.version.major, $v.version.minor))
            if ($vocab)
            {
                Write-Warning "Vocabulary already deployed"
                Write-Debug ($vocab | Out-String)

                Write-Verbose "Checking for dependencies"
                $policies = $ruleStore.GetDependentRuleSets($vocab)
                if ($policies)
                {
                    Write-Verbose "Found $($policies.Count) dependant policies"
                        
                    foreach ($p in $policies)
                    {
                        $policyExport = Export-Policy -Policy $p -Output $env:TEMP
                        $dependantPolicies.Push($policyExport)
                        Remove-Policy -Policy $p -Delete
                    }
                }

                Remove-Vocabulary -Vocabulary $vocab
            }
        }

        Write-Verbose "Publishing XML vocabulary(s)"
        if ($PSCmdlet.ShouldProcess($Path.FullName, "Published vocaulary(s)"))
        {
            $driver.ImportAndPublishFileRuleStore($Path.FullName)
        }

        if ($dependantPolicies.Count -gt 0)
        {
            Write-Verbose "Restoring dependant policy(s)"
            while ($dependantPolicies -gt 0)
            {
                Import-Policy -Path $dependantPolicies.Pop() -Deploy
            }
        }

        if ($CleanUp)
        {
            Write-Verbose "Removing XML"
            Remove-Item -Path $Path.FullName -Force
        }
    }
}


<#
    .SYNOPSIS
        Removed the specified vocabulary from the BRE store. Vocabulary can either be specified explicitly and passed from a pipeline. The default behaviour is just to undeploy the vocabulary, it can also optionally be deleted from the store entirely
    .EXAMPLE
        PS C:\> Remove-Vocabulary -Vocabulary $vocabulary
        Undeploys the specified vocabulary from BRE
    .EXAMPLE
        PS C:\> Remove-Vocabulary -Vocabulary $vocabulary -Delete
        Undeploys and deletes the specified vocabulary from BRE
    .EXAMPLE
        PS C:\> Get-Vocabulary | Remove-Vocabulary
        Undeploys policies from the pipeline
    .PARAMETER Vocabulary
        Vocabulary to be removed
    .PARAMETER Force
        Use to remove all dependent policies to allow the Vocabulary to be removed cleanly
    #>
function Remove-Vocabulary
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.RuleEngine.VocabularyInfo]$Vocabulary,
        [Parameter()]
        [switch]$Force
    )
    process
    {
        $dependantRules = $ruleStore.GetDependentRuleSets($Vocabulary)
        if ($dependantRules.Count -gt 0)
        {
            Write-Warning "Dependant rules found: $($dependantRules.Count)"
            Write-Debug ($dependantRules | Out-String)
            $dependantRules | Remove-Policy -Delete
        }

        if ($PSCmdlet.ShouldProcess(($Vocabulary | Out-String), "Removing vocabulary"))
        {
            $ruleStore.Remove($Vocabulary)
        }
    }
}
#endregion