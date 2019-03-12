# PowerShell.BRE

## Introduction

PowerShell.BRE is a PowerShell module to aid in the deployment and administration of BizTalk BRE policies and vocabularies.

Bundled with BizTalk is the **Business Rules Engine Deployment Wizard** which is limited to importing single policy/vocabulary XML files at a time (although the files can include multiple policies/vocabularies) and does not handle if the artifact being import already exists and is referenced by other artifacts.

## Description

To deal with this, the **Microsoft.RuleEngine** and **Microsoft.BizTalk.RuleEngineExtensions** are used to create typical read and delete operations as well as the import and export of artifacts. This is especially important for vocabularies. The BRE traditionally uses versioning, but during development there can often be a need to update an existing vocabulary without incrementing the version to avoid big version jumps during deployment. The import of vocabularies will check for existing vocabularies of the same name and version and check for dependent policies, the policies can then be exported temporarily to allow the vocabulary to be updated and then restore the policies