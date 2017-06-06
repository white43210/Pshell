$src = Import-Clixml c:\temp\source.xml #exported data from "good source"
$trgt = Get-Process
$prop = name
compare-object -ReferenceObject ($src) -DifferenceObject ($tget) -Property $prop
