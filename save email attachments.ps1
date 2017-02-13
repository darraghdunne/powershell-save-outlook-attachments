#<-------- Script --------->
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")
$f = $n.PickFolder()
$filepath = "c:\scripts\email"

$f.Items| foreach {
  $_.attachments|foreach {
   $a = $_.filename
    If ($a.Contains("xlsx")) {	
    $_.saveasfile((Join-Path $filepath $_.filename)) 
  }
 }
}

# <------- End Script ------->




