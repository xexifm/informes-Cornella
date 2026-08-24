#requires -Version 5.1
<#
.SYNOPSIS
  Dobles de prova de les funcions Format-* (Format.ps1), que escriuen a Word.

.DESCRIPTION
  Substitueixen les Format-* reals per versions que NOMES apunten a
  $global:emitCalls que se'ls ha cridat i amb quins arguments. Aixi es pot
  provar el motor de composicio (_WriteCatalegBody, Build-Document...) sense
  Word ni COM, i comprovar EXACTAMENT que rep cada funcio de format.

  Cada crida es una cadena "TIPUS|args". Els fills (IsChild) es marquen amb
  el sufix /CH (o " (fill)" a ITEM) per poder distingir-los del pare.

  US: dot-source ABANS del bloc de proves que el necessita. Cada dot-source
  reinicia $global:emitCalls, aixi que cada bloc comenca de zero:

      . (Join-Path $PSScriptRoot 'FormatDoubles.ps1')
      ... munta seccions i crida _WriteCatalegBody ...
      $global:emitCalls | Where-Object { $_ -like 'ITEM|*' }
#>

$global:emitCalls = New-Object System.Collections.ArrayList

function Format-Section    { param($s,$t) [void]$global:emitCalls.Add("SECT|$t") }
function Format-Subsection { param($s,$t) [void]$global:emitCalls.Add("SUB|$t") }
function Format-Item       { param($s,$n,$t,[switch]$IsChild) [void]$global:emitCalls.Add("ITEM|$n|$t" + $(if($IsChild){' (fill)'}else{''})) }
function Format-Body       { param($s,$t,[switch]$IsChild,[switch]$Bold,[switch]$Separat) [void]$global:emitCalls.Add('BODY'   + $(if($IsChild){'/CH'}else{''}) + $(if($Bold){'/N'}else{''}) + $(if($Separat){'/SEP'}else{''}) + "|$t") }
function Format-Bullet     { param($s,$t,[switch]$IsChild,[switch]$First) [void]$global:emitCalls.Add('BULLET' + $(if($IsChild){'/CH'}else{''}) + $(if($First){'/1r'}else{''}) + "|$t") }
function Format-Url        { param($s,$u,[switch]$IsChild) [void]$global:emitCalls.Add('URL'    + $(if($IsChild){'/CH'}else{''}) + "|$u") }
function Format-Spacer     { param($s) }
function Format-Append     { param($s,$t) [void]$global:emitCalls.Add("APPEND|$t") }
function Format-ListItem   { param($s,$t) [void]$global:emitCalls.Add("LLISTA|$t") }
function Format-Conclusion { param($s,$t) [void]$global:emitCalls.Add("CONCL|$t") }
function Format-ConclusionHeader { param($s,$t) [void]$global:emitCalls.Add("CONCLCAP|$t") }
function Format-Note       { param($s,$t) [void]$global:emitCalls.Add("NOTE|$t") }
function Format-Label      { param($s,$t) [void]$global:emitCalls.Add("LABEL|$t") }
function Format-Plain      { param($s,$t,[switch]$Bold,[int]$Size=0) [void]$global:emitCalls.Add('PLA' + $(if($Bold){'/N'}else{''}) + $(if($Size -gt 0){"/sz$Size"}else{''}) + "|$t") }
