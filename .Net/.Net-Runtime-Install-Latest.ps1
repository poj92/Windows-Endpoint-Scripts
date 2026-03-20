#Requires -Version 5.1
<#
DotNet-MinKeep-ConditionalInstall-Cleanup.ps1

What it does:
- Reports installed .NET inventory: SDK + Runtime + ASP.NET Core + Windows Desktop (x64 + optional x86).
- Removes versions below MinKeepVersion (best effort: uninstall tool + folder cleanup).
- Installs latest for a family ONLY if that family has below-min versions detected (i.e. would be removed).

Channel selection:
- TargetChannel defaults to major.minor of MinKeepVersion (prevents unintended upgrades to higher channels like 10.0).

Datto env vars:
  DotNet_MinKeepVersion (required)
  DotNet_TargetChannel
  DotNet_ReportOnly
  DotNet_IncludeX86
  DotNet_LatestLTSOnly
  DotNet_ForceUninstallTool
  DotNet_LogPath

Exit codes:
  0 = no action needed (no below-min versions found)
  1 = changes made (installed and/or removed)
  2 = report only; changes would be made
  3 = error
#>

At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:246 
char:59
+       $out += [pscustomobject]@{ Family=$fam; Arch=($arch ?? ''); Ver ...
+                                                           ~~
Unexpected token '??' in expression or statement.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:246 
char:58
+       $out += [pscustomobject]@{ Family=$fam; Arch=($arch ?? ''); Ver ...
+                                                          ~
Missing closing ')' in expression.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:246 
char:58
+       $out += [pscustomobject]@{ Family=$fam; Arch=($arch ?? ''); Ver ...
+                                                          ~
The hash literal was incomplete.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:245 
char:26
+     if($ver -lt $MinKeep){
+                          ~
Missing closing '}' in statement block or type definition.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:224 
char:32
+   foreach($e in Get-ArpEntries){
+                                ~
Missing closing '}' in statement block or type definition.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:221 
char:70
+ ... unction Get-DotNetArpBelowMin([Version]$MinKeep,[switch]$IncludeX86){
+                                                                         ~
Missing closing '}' in statement block or type definition.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:246 
char:64
+ ...    $out += [pscustomobject]@{ Family=$fam; Arch=($arch ?? ''); Versio ...
+                                                                 ~
Unexpected token ')' in expression or statement.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:246 
char:90
+ ... omobject]@{ Family=$fam; Arch=($arch ?? ''); Version=$ver; Entry=$e }
+                                                                         ~
Unexpected token '}' in expression or statement.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:247 
char:5
+     }
+     ~
Unexpected token '}' in expression or statement.
At C:\ProgramData\CentraStage\Packages\7e6b4020-8e06-4cf6-8245-e59f9574d88d#\DotNet-Runtime-Install-Latest.ps1:248 
char:3
+   }
+   ~
Unexpected token '}' in expression or statement.
Not all parse errors were reported.  Correct the reported errors and try again.
    + CategoryInfo          : ParserError: (:) [], ParseException
    + FullyQualifiedErrorId : UnexpectedToken