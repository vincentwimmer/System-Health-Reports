#Use Static path for Task Scheduler
#$fpath = "C:\inetpub\wwwroot\healthreports"

#Use Local path for testing scripts
$fpath = "."

#Create HTML file
$Style = "
<style>
#all-buttons {
margin: auto;
width: 950px;
}

#header {
  position: relative;
  background-color: #403d4e;
  border: none;
  font-size: 28px;
  color: #FFFFFF;
  padding: 30px;
  text-align: center;
}

#info {
  position: relative;
  background-color: #eae7f7;
  border: none;
  font-size: 16px;
  color: #403d4e;
  padding: 20px;
  margin-bottom: 15px;
}

#madeby {
  position: relative;
  border: none;
  text-align: center;
  font-size: 12px;
  color: #a9a6b7;
  padding: 20px;
  margin-bottom: 15px;  
}

body {
  margin-top:0px;
  background-color:#292533;
  color:#a9a6b7;
  font-family: Arial, Helvetica, sans-serif;
}

.button-new {
  position: relative;
  background-color: #3f9fff;
  border: none;
  font-size: 28px;
  color: #FFFFFF;
  padding: 20px;
  width: 100%;
  text-align: left;
  -webkit-transition-duration: 0.4s; /* Safari */
  transition-duration: 0.4s;
  text-decoration: none;
  overflow: hidden;
  cursor: pointer;
  margin-bottom: 15px;
}

.button-new:after {
  background: #f1f1f1;
  background-color: #3f9fff;
  display: block;
  position: absolute;
  padding-top: 300%;
  padding-left: 350%;
  margin-left: -20px !important;
  margin-top: -120%;
  opacity: 0;
  transition: all 0.8s
}

.button-new:active:after {
  padding: 0;
  margin: 0;
  opacity: 1;
  transition: 0s
}

.button-old {
  position: relative;
  background-color: #3e3c47;
  border: none;
  font-size: 28px;
  color: #FFFFFF;
  padding: 20px;
  width: 100%;
  text-align: left;
  -webkit-transition-duration: 0.4s; /* Safari */
  transition-duration: 0.4s;
  text-decoration: none;
  overflow: hidden;
  cursor: pointer;
  margin-bottom: 15px;
}

.button-old:after {
  background: #f1f1f1;
  background-color: #3e3c47;
  display: block;
  position: absolute;
  padding-top: 300%;
  padding-left: 350%;
  margin-left: -20px !important;
  margin-top: -120%;
  opacity: 0;
  transition: all 0.8s
}

.button-old:active:after {
  padding: 0;
  margin: 0;
  opacity: 1;
  transition: 0s
}
</style>
"

$bodytop = "
<body>
<div id='all-buttons'>
<div id='header'>System Health Reports</div>
<div id='info'>
<br>
Notes:
<ul>
  <li>Latest reports are highlighted blue.</li>
  <li>This page is recompiled after retrieval completes.</li>
  <li>Reports stored in:
"

$reportslocation = "$fpath\Reports" # | Resolve-Path

$bodybreakforpath = "</li>
  <li>Powered by Powershell.</li>
</ul>
</div>
"

$bodybottom = "
<div id='madeby'>Created by Vincent Wimmer.</div>
</div>

</body>
"

$SourceFile = "$fpath\FileList.txt"
$today     = Get-Date -format M-dd-yyyy
$File = Get-Content $SourceFile
$FileLine = @()
Foreach ($Line in $File) {
  $MyObject = New-Object -TypeName PSObject
  
  Add-Member -InputObject $MyObject -Type NoteProperty -Name HealthCheck -Value "$Line"
  echo $Line

  if ($Line -like "*$($today)*") {
  $link = ".\Reports\$Line"
  $FileLine += "
  <div>
  <a href='$link'>
  <button class='button-new'>
  "
  $FileLine += $Line
  $FileLine += "
  </button>
  </a>
  </div>
  "
  }
  else {
  $link = ".\Reports\$Line"
  $FileLine += "
  <div>
  <a href='$link'>
  <button class='button-old'>
  "
  $FileLine += $Line
  $FileLine += "
  </button>
  </a>
  </div>
  "
  }
}

ConvertTo-HTML -Head $Style -PostContent "$bodytop $reportslocation $bodybreakforpath $FileLine $bodybottom" -Title "System Health Check" | Out-File "$fpath\index.html"