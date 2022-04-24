<################################################################################ 

  Name    : Atera - On Call Text Alerts
  Version : 0.2
  Author  : Robert Brown
  Email   : robertdcbrown@gmail.com
  Date    : 4/20/2022
  
################################################################################>

### Loading external assemblies ###
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'


### Definitons ###
$AppDataFolder = "$env:APPDATA\AteraTXT"
$ConfigFile = "ateraTXT.config"
$TicketLogCSV = "ticketLog.csv"
$FullConfigPath = "$AppDataFolder\$ConfigFile"
$FullLogPath = "$AppDataFolder\$TicketLogCSV"
$timer = New-Object System.Windows.Forms.Timer
$timer2 = New-Object System.Windows.Forms.Timer
#OnIT Logo in base 64 format
$base64ImageString = "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"
$imageBytes = [Convert]::FromBase64String($base64ImageString)
$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$ms.Write($imageBytes, 0, $imageBytes.Length);
$img = [System.Drawing.Image]::FromStream($ms, $true)


$FMmain = New-Object System.Windows.Forms.Form
$box_console = New-Object System.Windows.Forms.GroupBox
$txt_console = New-Object System.Windows.Forms.TextBox
$CB_onCallTech = New-Object System.Windows.Forms.ComboBox
$label1 = New-Object System.Windows.Forms.Label
$CB_Holiday = New-Object System.Windows.Forms.ComboBox
$label2 = New-Object System.Windows.Forms.Label
$dataGridView1 = New-Object System.Windows.Forms.DataGridView
$label3 = New-Object System.Windows.Forms.Label
$button_save = New-Object System.Windows.Forms.Button
$txt_IsItOnCall = New-Object System.Windows.Forms.TextBox
$label4 = New-Object System.Windows.Forms.Label
$txt_onCallTech = New-Object System.Windows.Forms.TextBox
$label5 = New-Object System.Windows.Forms.Label
$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$label6 = New-Object System.Windows.Forms.Label
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$box_oncall = New-Object System.Windows.Forms.GroupBox
$Main = New-Object System.Windows.Forms.TabControl
$tabPage_Main = New-Object System.Windows.Forms.TabPage
$tabPage_Settings = New-Object System.Windows.Forms.TabPage
$button_start = New-Object System.Windows.Forms.Button
$button_stop = New-Object System.Windows.Forms.Button
$dataGridView2 = New-Object System.Windows.Forms.DataGridView
$label7 = New-Object System.Windows.Forms.Label
$label8 = New-Object System.Windows.Forms.Label
$label9 = New-Object System.Windows.Forms.Label
$text_AteraToken = New-Object System.Windows.Forms.TextBox
$text_TwilioToken = New-Object System.Windows.Forms.TextBox
$label10 = New-Object System.Windows.Forms.Label
$text_TwilioSID = New-Object System.Windows.Forms.TextBox
$text_TwilioNumber = New-Object System.Windows.Forms.TextBox
$label11 = New-Object System.Windows.Forms.Label
$label12 = New-Object System.Windows.Forms.Label
$button_save2 = New-Object System.Windows.Forms.Button
$label15 = New-Object System.Windows.Forms.Label
$backgroundWorker1 = New-Object System.ComponentModel.BackgroundWorker
$UpDown_Interval = New-Object System.Windows.Forms.NumericUpDown
$label16 = New-Object System.Windows.Forms.Label
$process1 = New-Object System.Diagnostics.Process
$label13 = New-Object System.Windows.Forms.Label
$label14 = New-Object System.Windows.Forms.Label
$label19 = New-Object System.Windows.Forms.Label
$label20 = New-Object System.Windows.Forms.Label
$label21 = New-Object System.Windows.Forms.Label
$label22 = New-Object System.Windows.Forms.Label
$label23 = New-Object System.Windows.Forms.Label
$label28 = New-Object System.Windows.Forms.Label
$label29 = New-Object System.Windows.Forms.Label
$label30 = New-Object System.Windows.Forms.Label
$LinkLabel1 = New-Object System.Windows.Forms.Label
$LinkLabel2 = New-Object System.Windows.Forms.Label
$text_SundayClose = New-Object System.Windows.Forms.TextBox
$text_SundayOpen = New-Object System.Windows.Forms.TextBox
$text_TuesdayClose = New-Object System.Windows.Forms.TextBox
$text_TuesdayOpen = New-Object System.Windows.Forms.TextBox
$text_MondayClose = New-Object System.Windows.Forms.TextBox
$text_MondayOpen = New-Object System.Windows.Forms.TextBox
$text_SatClose = New-Object System.Windows.Forms.TextBox
$text_SatOpen = New-Object System.Windows.Forms.TextBox
$text_FridayClose = New-Object System.Windows.Forms.TextBox
$text_FridayOpen = New-Object System.Windows.Forms.TextBox
$text_ThursClose = New-Object System.Windows.Forms.TextBox
$text_ThursOpen = New-Object System.Windows.Forms.TextBox
$text_WedClose = New-Object System.Windows.Forms.TextBox
$text_WedOpen = New-Object System.Windows.Forms.TextBox
$label24 = New-Object System.Windows.Forms.Label
$label25 = New-Object System.Windows.Forms.Label
$button_clear = New-Object System.Windows.Forms.Button
$label26 = New-Object System.Windows.Forms.Label
$checkBox_AutoRun = New-Object System.Windows.Forms.CheckBox
$checkBox_Sunday = New-Object System.Windows.Forms.CheckBox
$checkBox_Monday = New-Object System.Windows.Forms.CheckBox
$checkBox_Wed = New-Object System.Windows.Forms.CheckBox
$checkBox_Tues = New-Object System.Windows.Forms.CheckBox
$checkBox_Friday = New-Object System.Windows.Forms.CheckBox
$checkBox_Thurs = New-Object System.Windows.Forms.CheckBox
$checkBox_Sat = New-Object System.Windows.Forms.CheckBox
$label27 = New-Object System.Windows.Forms.Label
$textBox1 = New-Object System.Windows.Forms.TextBox



### Functions ###
Function FirstRun {
    if (Test-Path -Path $AppDataFolder) {
    }
    else {
        New-Item -Path "$env:APPDATA" -Name "AteraTXT" -ItemType "directory" | out-null
        $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
        $txt_console.AppendText("$logDate - Creating app data folder ($AppDataFolder\AteraTXT)`r`n")
    }

    if (Test-Path -Path "$FullConfigPath") {
    }
    else {
        New-Item -Path "$env:APPDATA\AteraTXT" -Name "$ConfigFile" -ItemType File | out-null
        $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
        $txt_console.AppendText("$logDate - Creating config file with defaults ($FullConfigPath)`r`n")
        "AteraAPIKey:Enter Key;TwilioToken:Enter Token;TwilioSID:Enter SID;TwilioNumber:Enter Number;AutoRun:0;Refresh:5;LastTech:Example1;Sunday:Closed;Monday:0800,1700;Tuesday:0800,1700;Wednesday:0800,1700;Thursday:0800,1700;Friday:0800,1700;Saturday:Closed;Tech:Example1,+12223334444;Tech:Example2,+19998887777" | add-content -path "$FullConfigPath"
    }

    if (Test-Path -Path "$FullLogPath") {
    }
    else {
        $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
        $txt_console.AppendText("$logDate - Creating lof file ($FullLogPath)`r`n")
        New-Item -Path "$env:APPDATA\AteraTXT" -Name "$TicketLogCSV" -ItemType File | out-null

        "Ticket Number,Ticket Title,Client,Contact,Ticket Created Date,Ticket Created Time,On Call Hours,On Call Tech" | add-content -path "$FullLogPath"
        "00000,Test Ticket,Test Client,Test Contact,2022-04-11,19:18:27,No,Stinky Pete" | add-content -path "$FullLogPath"
    }
}



Function Get-TinyURL { 
    #PowerShell - Get-TinyURL API Call
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] 
        [String]
        $sHTTPLink
    )

    if ($sHTTPLink.StartsWith("http://") -eq $true -or $sHTTPLink.StartsWith("https://") -eq $true)
        {
            if ($WebClient -eq $null) {$Global:WebClient=new-object System.Net.WebClient }
            $webClient.DownloadString("http://tinyurl.com/api-create.php?url="  + [System.Web.HttpUtility]::UrlEncode($sHTTPLink))
        }
}


function Reload {
    #Copy Tech Names to Array
    $CurrentTech = $CB_onCallTech.SelectedItem
    $RowCount = $dataGridView2.RowCount
    $RowCount -= 2
    $Main = 1
    $Techs = @()
    While ($Main -eq 1){
        $Name = $dataGridView2[0,$RowCount].Value
        $Techs += "$Name"
        $RowCount -= 1
        If ($RowCount -eq -1){
            $Main = 0
        }
    }
    $Techs = $Techs | Sort-Object
    $AllTechs = ""
    Foreach ($TechName in $Techs){
        $AllTechs += "$TechName;"

    }
    $CB_onCallTech.Items.Clear()
    $CB_onCallTech.Items.AddRange($Techs)
    $CB_onCallTech.SelectedItem = "$CurrentTech"

}




function UpdateOnCallTech {

    $tech = $CB_onCallTech.SelectedItem

    $RowCount = $dataGridView2.RowCount
    $RowCount -= 2
    $Main = 1
    $Techs = @()
    While ($Main -eq 1){
        $Name = $dataGridView2[0,$RowCount].Value
        $Number = $dataGridView2[1,$RowCount].Value
        $Techs += "$Name,$Number"
        $RowCount -= 1
        If ($RowCount -eq -1){
            $Main = 0
        }
    }
    $Techs = $Techs | Sort-Object
    
    foreach ($Technician in $techs){
        if ($Technician -like "*$tech*"){
            $techphone = $Technician.Split(",")[1]
        }
    }
    $txt_onCallTech.Text = "$tech"
    return "$tech-$techphone"
}


function ForceHoliday {
    if ($CB_Holiday.SelectedIndex -eq 0){
        $Holiday = 0
    }
    elseif ($CB_Holiday.SelectedIndex -eq 1){
        $Holiday = 1
    }
    return $Holiday
}

function IsItOnCall {
    $now = Get-Date
    $CurrentDay = Get-Date -f dddd
    $Holiday = ForceHoliday
    #Determine if it is oncall hours
    If ($holiday -eq 1){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, holiday hours active"
    }
    elseif ($CurrentDay -like "*Sunday*"){
        $SundayStatus = $checkBox_Sunday.CheckState
        $SunOpenTime = $text_SundayOpen.Text
        $SunCloseTime = $text_SundayClose.Text
        if ($SunOpenTime){
            $SunOpenTime = $SunOpenTime.Insert(2,':')
        }
        if ($SunCloseTime){
            $SunCloseTime = $SunCloseTime.Insert(2,':')
        }
        If ($SundayStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Sunday"
        }

        Elseif ($now -lt (Get-Date $SunOpenTime) -or $now -gt (Get-Date $SunCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    elseif ($CurrentDay -like "*Monday*"){
        $MondayStatus = $checkBox_Monday.CheckState
        $MonOpenTime = $text_MondayOpen.Text
        $MonCloseTime = $text_MondayClose.Text
        if ($MonOpenTime){
            $MonOpenTime = $MonOpenTime.Insert(2,':')
        }
        if ($MonCloseTime){
            $MonCloseTime = $MonCloseTime.Insert(2,':')
        }
        If ($MondayStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Monday"
        }

        Elseif ($now -lt (Get-Date $MonOpenTime) -or $now -gt (Get-Date $MonCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    elseif ($CurrentDay -like "*Tuesday*"){
        $TuesdayStatus = $checkBox_Tues.CheckState
        $TuesOpenTime = $text_TuesdayOpen.Text
        $TuesCloseTime = $text_TuesdayClose.Text
        if ($TuesOpenTime){
            $TuesOpenTime = $TuesOpenTime.Insert(2,':')
        }
        if ($TuesCloseTime){
            $TuesCloseTime = $TuesCloseTime.Insert(2,':')
        }
        If ($TuesdayStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Tuesday"
        }
        Elseif ($now -lt (Get-Date $TuesOpenTime) -or $now -gt (Get-Date $TuesCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    elseif ($CurrentDay -like "*Wedn*"){
        $WedStatus = $checkBox_Wed.CheckState
        $WedOpenTime = $text_WedOpen.Text
        $WedCloseTime = $text_WedClose.Text
        if ($WedOpenTime){
            $WedOpenTime = $WedOpenTime.Insert(2,':')
        }
        if ($WedCloseTime){
            $WedCloseTime = $WedCloseTime.Insert(2,':')
        }
        If ($WedStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Wednesday"
        }
        Elseif ($now -lt (Get-Date $WedOpenTime) -or $now -gt (Get-Date $WedCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    elseif ($CurrentDay -like "*Thursday*"){
        $ThurStatus = $checkBox_Thurs.CheckState
        $ThursOpenTime = $text_ThursOpen.Text
        $ThursCloseTime = $text_ThursClose.Text
        if ($ThursOpenTime){
            $ThursOpenTime = $ThursOpenTime.Insert(2,':')
        }
        if ($ThursCloseTime){
            $ThursCloseTime = $ThursCloseTime.Insert(2,':')
        }
        If ($ThurStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Thursday"
        }
        Elseif ($now -lt (Get-Date $ThursOpenTime) -or $now -gt (Get-Date $ThursCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
        elseif ($CurrentDay -like "*Friday*"){
        $FridayStatus = $checkBox_Friday.CheckState
        $FriOpenTime = $text_FridayOpen.Text
        $FriCloseTime = $text_FridayClose.Text
        if ($FriOpenTime){
            $FriOpenTime = $FriOpenTime.Insert(2,':')
        }
        if ($FriCloseTime){
            $FriCloseTime = $FriCloseTime.Insert(2,':')
        }
        If ($FridayStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Friday"
        }
        Elseif ($now -lt (Get-Date $FriOpenTime) -or $now -gt (Get-Date $FriCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    elseif ($CurrentDay -like "*Saturday*"){
        $SaturdayStatus = $checkBox_Sat.CheckState
        $SatOpenTime = $text_SatOpen.Text
        $SatCloseTime = $text_SatClose.Text
        if ($SatOpenTime){
            $SatOpenTime = $SatOpenTime.Insert(2,':')
        }
        if ($SatCloseTime){
            $SatCloseTime = $SatCloseTime.Insert(2,':')
        }
        If ($SaturdayStatus -like "Checked"){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, closed on Saturday"
        }
        Elseif ($now -lt (Get-Date $SatOpenTime) -or $now -gt (Get-Date $SatCloseTime)){
            $OnCall = 1
            $txt_IsItOnCall.Text = "Yes, outside of business hours"
        }
        Else {
            $OnCall = 0
            $txt_IsItOnCall.Text = "No"
        }
    }
    return $OnCall
}

Function SendText{
    param(
        [String]$techPhone,
        [String]$body
    )
    ### Send Text Message ###
    $sid = $text_TwilioSID.Text
    $token = $text_TwilioToken.Text
    $number = $text_TwilioNumber.Text

    # Twilio API endpoint and POST params
    $url = "https://api.twilio.com/2010-04-01/Accounts/$sid/Messages.json"
    $params = @{ To = $techPhone; From = $number; Body = $body }

    # Create a credential object for HTTP basic auth
    $p = $token | ConvertTo-SecureString -asPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($sid, $p)

    # Make API request, selecting JSON properties from response
    Invoke-WebRequest $url -Method Post -Credential $credential -Body $params -UseBasicParsing |
    ConvertFrom-Json | Select sid, body
}

function startTimer { 
   $timer.start()
}

function stopTimer {
    $timer.Enabled = $false
}

function startTimer2 { 
   $timer2.start()
}

function stopTimer2 {
    $timer2.Enabled = $false
    $progressBar1.Value = 1
}


function Hide-Console{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}


function SaveConfig {
    $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
    $txt_console.AppendText("$logDate - Saving current config ($FullConfigPath)`r`n")
    $AteraTokenNew = $text_AteraToken.Text
    $TwilioTokenNew = $text_TwilioToken.Text
    $TwilioSIDNew = $text_TwilioSID.Text
    $TwilioNumberNew = $text_TwilioNumber.Text
    $RefreshNew = $UpDown_Interval.Text
    $AutoRun = $checkBox_AutoRun.CheckState
    If ($AutoRun -like "Checked"){
        $Autorun = 1
    }
    ElseIf ($AutoRun -like "Unchecked"){
        $Autorun = 0
    }

    if($checkBox_Sunday.checked -eq $true) {
        $SundayNew = "Closed"
    }
    Else{
        $SundayOpen = $text_SundayOpen.Text
        $SundayClose = $text_SundayClose.Text
        $SundayNew = "$SundayOpen,$SundayClose"
        $SundayNew = $SundayNew.Replace(":","")
    }
    if($checkBox_Monday.checked -eq $true) {
        $MondayNew = "Closed"
    }
    Else{
        $MondayOpen = $text_MondayOpen.Text
        $MondayClose = $text_MondayClose.Text
        $MondayNew = "$MondayOpen,$MondayClose"
        $MondayNew = $MondayNew.Replace(":","")
    }
    if($checkBox_Tues.checked -eq $true) {
        $TuesdayNew = "Closed"
    }
    Else{
        $TuesdayOpen = $text_TuesdayOpen.Text
        $TuesdayClose = $text_TuesdayClose.Text
        $TuesdayNew = "$TuesdayOpen,$TuesdayClose"
        $TuesdayNew = $TuesdayNew.Replace(":","")
    }
    if($checkBox_Wed.checked -eq $true) {
        $WednesdayNew = "Closed"
    }
    Else{
        $WedOpen = $text_WedOpen.Text
        $WedClose = $text_WedClose.Text
        $WednesdayNew = "$WedOpen,$WedClose"
        $WednesdayNew = $WednesdayNew.Replace(":","")
    }
    if($checkBox_Thurs.checked -eq $true) {
        $ThursdayNew = "Closed"
    }
    Else{
        $ThursOpen = $text_ThursOpen.Text
        $ThursClose = $text_ThursClose.Text
        $ThursdayNew = "$ThursOpen,$ThursClose"
        $ThursdayNew = $ThursdayNew.Replace(":","")
    }
    if($checkBox_Friday.checked -eq $true) {
        $FridayNew = "Closed"
    }
    Else{
        $FriOpen = $text_FridayOpen.Text
        $FriClose = $text_FridayClose.Text
        $FridayNew = "$FriOpen,$FriClose"
        $FridayNew = $FridayNew.Replace(":","")
    }

    if($checkBox_Sat.checked -eq $true) {
        $SaturdayNew = "Closed"
    }
    Else{
        $SatOpen = $text_SatOpen.Text
        $SatClose = $text_SatClose.Text
        $SaturdayNew = "$SatOpen,$SatClose"
        $SaturdayNew = $SaturdayNew.Replace(":","")
    }
    

    #Save Techs and Numbers
    $RowCount = $dataGridView2.RowCount
    $RowCount -= 2
    $Main = 1
    $Techs = @()
    While ($Main -eq 1){
        $Name = $dataGridView2[0,$RowCount].Value
        $Number = $dataGridView2[1,$RowCount].Value
        $Techs += "Tech:$Name,$Number"
        $RowCount -= 1
        If ($RowCount -eq -1){
            $Main = 0
        }
    }
    $Techs = $Techs | Sort-Object
    $AllTechs = ""
    Foreach ($TechNum in $Techs){
        $AllTechs += "$TechNum;"
    }

    #Get Last Tech
    $LastTech = $CB_onCallTech.SelectedItem

    $configNew = "AteraAPIKey:$AteraTokenNew;TwilioToken:$TwilioTokenNew;TwilioSID:$TwilioSIDNew;TwilioNumber:$TwilioNumberNew;AutoRun:$AutoRun;Refresh:$RefreshNew;LastTech:$LastTech;Sunday:$SundayNew;Monday:$MondayNew;Tuesday:$TuesdayNew;Wednesday:$WednesdayNew;Thursday:$ThursdayNew;Friday:$FridayNew;Saturday:$SaturdayNew;$AllTechs"

    Remove-Item -Path $FullConfigPath -Force
    New-Item -Path "$env:APPDATA\AteraTXT" -Name "$ConfigFile" -ItemType File | out-null
    "$configNew" | add-content -path "$FullConfigPath"



}


Function LoadConfig {
    $txt_console.AppendText("$logDate - Loading Config File ($FullConfigPath)`r`n")
    $Config = Get-Content -Path "$FullConfigPath"

    $ConfigSplit = $config.Split(";")
    $AteraAPIKey = $ConfigSplit[0].Split(":")[1]
    $TwilioToken = $ConfigSplit[1].Split(":")[1]
    $TwilioSID = $ConfigSplit[2].Split(":")[1]
    $TwilioNumber = $ConfigSplit[3].Split(":")[1]
    $AutoRunStatus = $ConfigSplit[4].Split(":")[1]
    $RefreshRate = $ConfigSplit[5].Split(":")[1]
    $LastTech = $ConfigSplit[6].Split(":")[1]
    $SundayHours = $ConfigSplit[7].Split(":")[1]
    $MondayHours = $ConfigSplit[8].Split(":")[1]
    $TuesdayHours = $ConfigSplit[9].Split(":")[1]
    $WednesdayHours = $ConfigSplit[10].Split(":")[1]
    $ThursdayHours = $ConfigSplit[11].Split(":")[1]
    $FridayHours = $ConfigSplit[12].Split(":")[1]
    $SaturdayHours = $ConfigSplit[13].Split(":")[1]
    $text_AteraToken.Text = $AteraAPIKey
    $text_TwilioToken.Text = $TwilioToken
    $text_TwilioSID.Text = $TwilioSID
    $text_TwilioNumber.Text = $TwilioNumber


    If ($SundayHours -like "*Closed*"){
        "Closed"
        $checkBox_Sunday.checked = $true
    }
    Else {
        $SundayOpen = $SundayHours.Split(",")[0]
        $SundayClose = $SundayHours.Split(",")[1]
        $text_SundayOpen.Text = "$SundayOpen"
        $text_SundayClose.Text = "$SundayClose"
    }
    If ($MondayHours -like "*Closed*"){
        $checkBox_Monday.checked = $true
    }
    Else {
        $MondayOpen = $MondayHours.Split(",")[0]
        $MondayClose = $MondayHours.Split(",")[1]
        $text_MondayOpen.Text = "$MondayOpen"
        $text_MondayClose.Text = "$MondayClose"
    }
    If ($TuesdayHours -like "*Closed*"){
        $checkBox_Tues.checked = $true
    }
    Else {
        $TuesdayOpen = $TuesdayHours.Split(",")[0]
        $TuesdayClose = $TuesdayHours.Split(",")[1]
        $text_TuesdayOpen.Text = "$TuesdayOpen"
        $text_TuesdayClose.Text = "$TuesdayClose"
    }
    If ($WednesdayHours -like "*Closed*"){
        $checkBox_Wed.checked = $true
    }
    Else {
        $WedOpen = $WednesdayHours.Split(",")[0]
        $WedClose = $WednesdayHours.Split(",")[1]
        $text_WedOpen.Text = "$WedOpen"
        $text_WedClose.Text = "$WedClose"
    }
    If ($ThursdayHours -like "*Closed*"){
        $checkBox_Thurs.checked = $true
    }
    Else {
        $ThursdayOpen = $ThursdayHours.Split(",")[0]
        $ThursdayClose = $ThursdayHours.Split(",")[1]
        $text_ThursOpen.Text = "$ThursdayOpen"
        $text_ThursClose.Text = "$ThursdayClose"
    }
    If ($FridayHours -like "*Closed*"){
        $checkBox_Friday.checked = $true
    }
    Else {
        $FridayOpen = $FridayHours.Split(",")[0]
        $FridayClose = $FridayHours.Split(",")[1]
        $text_FridayOpen.Text = "$FridayOpen"
        $text_FridayClose.Text = "$FridayClose"
    }
    If ($SaturdayHours -like "*Closed*"){
        $checkBox_Sat.checked = $true
    }
    Else {
        $SatOpen = $SaturdayHours.Split(",")[0]
        $SatClosed = $SaturdayHours.Split(",")[1]
        $text_SatOpen.Text = "$SatOpen"
        $text_SatClose.Text = "$SatClosed"
    }


    #Update AutoRun
    If ($AutoRunStatus -eq 1){
        $checkBox_AutoRun.Checked = $true
    }

    #Update Interval Timers
    $UpDown_Interval.Value = $RefreshRate
    $timer1MS = [int]$RefreshRate*60000
    $time2MS = $timer1MS/240
    $timer.Interval = $timer1MS
    $timer2.Interval = $time2MS



    #Load Tech List
    $dataGridView2.Rows.Clear();
    $CB_onCallTech_List = @()
    foreach ($tech in $ConfigSplit){
        If ($tech -like "Tech:*"){
            $techName = $tech.Split(",")[0]
            $techName = $techName.Split(":")[1]
            $techNumber = $tech.Split(",")[1]
            $dataGridView2.Rows.Add("$techName","$techNumber")
            $CB_onCallTech_List += $techName
        }
        $dataGridView2.Refresh();
    }
    Set-AteraAPIKey "$AteraAPIKey"

}

Function AutoRun {
    $Config = Get-Content -Path "$FullConfigPath"
    $ConfigSplit = $config.Split(";")
    $AutoRunStatus = $ConfigSplit[4].Split(":")[1]
    $LastTech = $ConfigSplit[6].Split(":")[1]

    If ($AutoRunStatus -eq 1){
        $CB_onCallTech.Text = "$LastTech"
        $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
        $txt_console.AppendText("$logDate - LastTech was $LastTech`r`n")

        $txt_console.AppendText("Starting Script`r`n")
        $button_stop.Enabled   = $true
        $button_start.Enabled   = $false
        UpdateOnCallTech
        ForceHoliday
        IsItOnCall
        startTimer
        startTimer2

    }

}



###################################################################################################################################



#Main Program
$timer2.add_tick({
    $progressBar1.PerformStep()
})

$timer.add_tick({
    $progressBar1.Value = 1
    ### Update Last 10 On Call Tickets ###
    $dataGridView1.Rows.Clear();
    $dataGridView1.Refresh();
    $csvdata10 = Import-Csv "$FullLogPath"
    $csvdata10 = $csvdata10 | Select-Object -Last 10
    foreach ($csvTicket in $csvdata10){
        $csvTicketNumber = $csvTicket.'Ticket Number'
        $csvTicketTitle = $csvTicket.'Ticket Title'
        $csvTicketClient = $csvTicket.'Client'
        $csvTicketContact = $csvTicket.'Contact'
        $csvTicketDate = $csvTicket.'Ticket Created Date'
        $csvTicketTime = $csvTicket.'Ticket Created Time'
        $csvTicketOnCall = $csvTicket.'On Call Hours'
        $csvTicketTech = $csvTicket.'On Call Tech'
        $dataGridView1.Rows.Add("$csvTicketNumber","$csvTicketTitle","$csvTicketClient","$csvTicketContact","$csvTicketDate","$csvTicketTime","$csvTicketOnCall","$csvTicketTech")
    }

    $techAndPhone = UpdateOnCallTech
    $tech = $techAndPhone.Split("-")[0]
    $techPhone = $techAndPhone.Split("-")[1]

    $logDate = Get-Date -f "MM/dd/yyyy HH:mm"
    $holiday = ForceHoliday
    $NewTicket = 0
    $csvdata = Import-Csv "$FullLogPath"
    $csvdata = $csvdata | Select-Object -Last 1
    $csvTicket = $csvdata.'Ticket Number'

    $oncall = IsItOnCall
    if ($oncall -eq 0){
        $IsOnCall = "No"
    }
    Else {
        $IsOnCall = "Yes"
    }

    #Get Last Ticket from Atera
    $txt_console.AppendText("$logDate - Getting last ticket from Atera`r`n")
    $OpenTickets = Get-AteraTickets -Open
    if ($OpenTickets -eq $null){
        $txt_console.AppendText("$logDate - No open tickets in Atera`r`n")
    }
    Else {
        $OpenTickets = $OpenTickets | Sort-Object -Property TicketID -Descending
        $TicketNumber = $OpenTickets[0].TicketID
        $TicketCustomer = $OpenTickets[0].CustomerName
        $TicketCustomer = $TicketCustomer.Replace(',','')
        $TicketTitle = $OpenTickets[0].TicketTitle
        $TicketTitle = $TicketTitle.Replace(',','')
        $TicketFN = $OpenTickets[0].EndUserFirstName
        $TicketLN = $OpenTickets[0].EndUserLastName
        $TicketContact = "$TicketFN $TicketLN"
        $TicketContact = $TicketContact.Replace(',','')
        $TicketURL = "https://app.atera.com/Admin#/ticket/$TicketNumber"
        $TicketDateTime = $OpenTickets[0].TicketCreatedDate
        $TicketDate = $TicketDateTime.Split("T")[0]
        $TicketTime = $TicketDateTime.Split("T")[1]
        $TicketTime = $TicketTime.Replace("Z","")

        #Compare ticket number to last ticket in CSV
        If ($TicketNumber -gt $csvTicket){
            $NewTicket = 1
            $txt_console.AppendText("$logDate - ***New ticket***`r`n                   Client: $TicketCustomer`r`n                   Contact: $TicketContact`r`n                   Title: $TicketTitle`r`n")
            $txt_console.AppendText("$logDate - Checking if it is on call`r`n")
        }
        Else {
            $txt_console.AppendText("$logDate - No new tickets found since last check`r`n")
        }

        If ($NewTicket -eq 1){
            #Add Ticket to CSV
            "$TicketNumber,$TicketTitle,$TicketCustomer,$TicketContact,$TicketDate,$TicketTime,$IsOnCall,$tech" | add-content -path "$FullLogPath"
        }



        If (($OnCall -eq 1) -and ($NewTicket -eq 1)){
            $txt_console.AppendText("$logDate - On Call, preparing to alert on call tech`r`n")

            #Generate short URL
            $txt_console.AppendText("$logDate - Generating Short URL for ticket # $TicketNumber`r`n")
            $ShortURL = Get-TinyURL "$TicketURL"
            $txt_console.AppendText("$logDate - Link generated - $ShortURL `r`n")

            #Send text to on call tech
            $body = "New On Call Ticket
Client: $TicketCustomer
User: $TicketContact
Subject: $TicketTitle
Link: $ShortURL"



            ### Send Text Message ###
            $txt_console.AppendText("$logDate - Sending text to $tech for ticket # $TicketNumber`r`n")

            SendText -techPhone "$techphone" -body "$body"
        }
        elseIf (($OnCall -eq 0) -and ($NewTicket -eq 1)){
            $txt_console.AppendText("$logDate - Not on call hours`r`n")
        }

    }

})


###################################################################################################################################



### Boxes ###

# box_console
$box_console.Controls.Add($button_clear)
$box_console.Controls.Add($txt_console)
$box_console.Location = New-Object System.Drawing.Point(16, 471)
$box_console.Name = "box_console"
$box_console.Size = New-Object System.Drawing.Size(825, 292)
$box_console.TabIndex = 2
$box_console.TabStop = $false
$box_console.Text = "Console"

# box_oncall
$box_oncall.Controls.Add($button_stop)
$box_oncall.Controls.Add($button_start)
$box_oncall.Controls.Add($pictureBox1)
$box_oncall.Controls.Add($label6)
$box_oncall.Controls.Add($progressBar1)
$box_oncall.Controls.Add($label5)
$box_oncall.Controls.Add($txt_onCallTech)
$box_oncall.Controls.Add($label4)
$box_oncall.Controls.Add($txt_IsItOnCall)
$box_oncall.Controls.Add($button_save)
$box_oncall.Controls.Add($label3)
$box_oncall.Controls.Add($dataGridView1)
$box_oncall.Controls.Add($label2)
$box_oncall.Controls.Add($CB_Holiday)
$box_oncall.Controls.Add($label1)
$box_oncall.Controls.Add($CB_onCallTech)
$box_oncall.Location = New-Object System.Drawing.Point(3, 3)
$box_oncall.Name = "box_oncall"
$box_oncall.Size = New-Object System.Drawing.Size(815, 424)
$box_oncall.TabIndex = 3
$box_oncall.TabStop = $false
$box_oncall.Text = "On Call Options"


### Picture Box ###

# pictureBox1
$pictureBox1.Location = New-Object System.Drawing.Point(609, 14)
$pictureBox1.Name = "pictureBox1"
$pictureBox1.Size = New-Object System.Drawing.Size(200, 100)
$pictureBox1.TabIndex = 12
$pictureBox1.TabStop = $false
$pictureBox1.Image = $img


### Labels ###

# label1
$label1.AutoSize = $true
$label1.Location = New-Object System.Drawing.Point(48, 22)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(102, 13)
$label1.TabIndex = 1
$label1.Text = "Select On Call Tech"
# label2
$label2.AutoSize = $true
$label2.Location = New-Object System.Drawing.Point(6, 48)
$label2.Name = "label2"
$label2.Size = New-Object System.Drawing.Size(144, 13)
$label2.TabIndex = 3
$label2.Text = "Force On Call Texts (Holiday)"
# label3
$label3.AutoSize = $true
$label3.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label3.Location = New-Object System.Drawing.Point(300, 255)
$label3.Name = "label3"
$label3.Size = New-Object System.Drawing.Size(194, 20)
$label3.TabIndex = 5
$label3.Text = "Last 10 On Call Tickets"
# label4
$label4.AutoSize = $true
$label4.Location = New-Object System.Drawing.Point(75, 75)
$label4.Name = "label4"
$label4.Size = New-Object System.Drawing.Size(75, 13)
$label4.TabIndex = 8
$label4.Text = "On call hours?"
# label5
$label5.AutoSize = $true
$label5.Location = New-Object System.Drawing.Point(48, 101)
$label5.Name = "label5"
$label5.Size = New-Object System.Drawing.Size(102, 13)
$label5.TabIndex = 10
$label5.Text = "Current tech on call:"
# label6
$label6.AutoSize = $true
$label6.Location = New-Object System.Drawing.Point(103, 194)
$label6.Name = "label6"
$label6.Size = New-Object System.Drawing.Size(52, 13)
$label6.TabIndex = 11
$label6.Text = "Next Run"
# label7
$label7.AutoSize = $true
$label7.Location = New-Object System.Drawing.Point(68, 50)
$label7.Name = "label7"
$label7.Size = New-Object System.Drawing.Size(86, 13)
$label7.TabIndex = 1
$label7.Text = "Atera API Key"
# label8
$label8.AutoSize = $true
$label8.Location = New-Object System.Drawing.Point(55, 81)
$label8.Name = "label8"
$label8.Size = New-Object System.Drawing.Size(88, 13)
$label8.TabIndex = 2
$label8.Text = "Twilio Auth Token"
# label9
$label9.AutoSize = $true
$label9.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label9.Location = New-Object System.Drawing.Point(623, 10)
$label9.Name = "label9"
$label9.Size = New-Object System.Drawing.Size(50, 16)
$label9.TabIndex = 3
$label9.Text = "Techs"
# label10
$label10.AutoSize = $true
$label10.Location = New-Object System.Drawing.Point(50, 109)
$label10.Name = "label10"
$label10.Size = New-Object System.Drawing.Size(55, 13)
$label10.TabIndex = 6
$label10.Text = "Twilio Account SID"
# label11
$label11.AutoSize = $true
$label11.Location = New-Object System.Drawing.Point(70, 135)
$label11.Name = "label11"
$label11.Size = New-Object System.Drawing.Size(74, 13)
$label11.TabIndex = 9
$label11.Text = "Twilio Number"
# label12
$label12.AutoSize = $true
$label12.Location = New-Object System.Drawing.Point(350, 135)
$label12.Name = "label12"
$label12.Size = New-Object System.Drawing.Size(114, 13)
$label12.TabIndex = 10
$label12.Text = "Format +12223334444"
# label13
$label13.AutoSize = $true
$label13.Location = New-Object System.Drawing.Point(547, 226)
$label13.Name = "label13"
$label13.Size = New-Object System.Drawing.Size(43, 13)
$label13.TabIndex = 22
$label13.Text = "Sunday"
# label14
$label14.AutoSize = $true
$label14.Location = New-Object System.Drawing.Point(546, 252)
$label14.Name = "label14"
$label14.Size = New-Object System.Drawing.Size(45, 13)
$label14.TabIndex = 23
$label14.Text = "Monday"
# label15
$label15.AutoSize = $true
$label15.Location = New-Object System.Drawing.Point(92, 228)
$label15.Name = "label15"
$label15.Size = New-Object System.Drawing.Size(128, 13)
$label15.TabIndex = 15
$label15.Text = "Refresh Interval (Minutes)"
# label16
$label16.AutoSize = $true
$label16.Location = New-Object System.Drawing.Point(282, 228)
$label16.Name = "label16"
$label16.Size = New-Object System.Drawing.Size(53, 13)
$label16.TabIndex = 18
$label16.Text = "Default: 5"
# label19
$label19.AutoSize = $true
$label19.Location = New-Object System.Drawing.Point(544, 278)
$label19.Name = "label19"
$label19.Size = New-Object System.Drawing.Size(48, 13)
$label19.TabIndex = 24
$label19.Text = "Tuesday"
# label20
$label20.AutoSize = $true
$label20.Location = New-Object System.Drawing.Point(526, 304)
$label20.Name = "label20"
$label20.Size = New-Object System.Drawing.Size(64, 13)
$label20.TabIndex = 25
$label20.Text = "Wednesday"
# label21
$label21.AutoSize = $true
$label21.Location = New-Object System.Drawing.Point(542, 382)
$label21.Name = "label21"
$label21.Size = New-Object System.Drawing.Size(49, 13)
$label21.TabIndex = 28
$label21.Text = "Saturday"
# label22
$label22.AutoSize = $true
$label22.Location = New-Object System.Drawing.Point(555, 356)
$label22.Name = "label22"
$label22.Size = New-Object System.Drawing.Size(35, 13)
$label22.TabIndex = 27
$label22.Text = "Friday"
# label23
$label23.AutoSize = $true
$label23.Location = New-Object System.Drawing.Point(544, 330)
$label23.Name = "label23"
$label23.Size = New-Object System.Drawing.Size(46, 13)
$label23.TabIndex = 26
$label23.Text = "Thurday"
# label24
$label24.AutoSize = $true
$label24.Location = New-Object System.Drawing.Point(603, 200)
$label24.Name = "label24"
$label24.Size = New-Object System.Drawing.Size(33, 13)
$label24.TabIndex = 45
$label24.Text = "Open"
# label25
$label25.AutoSize = $true
$label25.Location = New-Object System.Drawing.Point(674, 200)
$label25.Name = "label25"
$label25.Size = New-Object System.Drawing.Size(33, 13)
$label25.TabIndex = 46
$label25.Text = "Close"
# label26
$label26.AutoSize = $true
$label26.Location = New-Object System.Drawing.Point(732, 207)
$label26.Name = "label26"
$label26.Size = New-Object System.Drawing.Size(39, 13)
$label26.TabIndex = 47
$label26.Text = "Closed"
# label27
$label27.AutoSize = $true
$label27.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label27.Location = New-Object System.Drawing.Point(552, 175)
$label27.Name = "label27"
$label27.Size = New-Object System.Drawing.Size(226, 16)
$label27.TabIndex = 55
$label27.Text = "Business Hours (24 hour format)"
# label28
$label28.AutoSize = $true
$label28.Location = New-Object System.Drawing.Point(90, 170)
$label28.Name = "label28"
$label28.Size = New-Object System.Drawing.Size(74, 13)
$label28.TabIndex = 9
$label28.Text = "Auto Start"
# label29
$label29.AutoSize = $true
$label29.Location = New-Object System.Drawing.Point(170, 160)
$label29.Name = "label29"
$label29.Size = New-Object System.Drawing.Size(74, 13)
$label29.TabIndex = 9
$label29.Text = 'If Checked, program will run without needing to press "Start" button'
# label30
$label30.AutoSize = $true
$label30.Location = New-Object System.Drawing.Point(170, 180)
$label30.Name = "label30"
$label30.Size = New-Object System.Drawing.Size(74, 13)
$label30.TabIndex = 9
$label30.Text = 'On-call tech will default to last tech used.'

### Link Labels ###
$LinkLabel1 = New-Object System.Windows.Forms.LinkLabel
$LinkLabel1.Location = New-Object System.Drawing.Size(360,52)
$LinkLabel1.Size = New-Object System.Drawing.Size(130,20)
$LinkLabel1.LinkColor = "BLUE"
$LinkLabel1.ActiveLinkColor = "RED"
$LinkLabel1.Text = "Get Atera API Key Here"
$LinkLabel1.add_Click({[system.Diagnostics.Process]::start("https://app.atera.com/Admin#/admin/api")})

$LinkLabel2 = New-Object System.Windows.Forms.LinkLabel
$LinkLabel2.Location = New-Object System.Drawing.Size(360,82)
$LinkLabel2.Size = New-Object System.Drawing.Size(130,20)
$LinkLabel2.LinkColor = "BLUE"
$LinkLabel2.ActiveLinkColor = "RED"
$LinkLabel2.Text = "Get Twilio API Info Here"
$LinkLabel2.add_Click({[system.Diagnostics.Process]::start("https://console.twilio.com/?frameUrl=%2Fconsole%3Fx-target-region%3Dus1")})


### Progress Bar ###

# progressBar1
$progressBar1.Location = New-Object System.Drawing.Point(161, 189)
$progressBar1.Name = "progressBar1"
$progressBar1.Size = New-Object System.Drawing.Size(495, 23)
$progressBar1.TabIndex = 4
$progressBar1.Style="Continuous"
$progressBar1.Maximum = 240 #CODEME - Calculate from Up/Down Interval
$progressBar1.Minimum = 1
$progressBar1.Step = 1


### Drop Down Boxes ###

# CB_onCallTech
$CB_onCallTech.FormattingEnabled = $true
$CB_onCallTech.Location = New-Object System.Drawing.Point(161, 18)
$CB_onCallTech.Name = "CB_onCallTech"
$CB_onCallTech.Size = New-Object System.Drawing.Size(151, 21)
$CB_onCallTech.TabIndex = 0


# CB_Holiday
$CB_Holiday.FormattingEnabled = $true
$CB_Holiday.Location = New-Object System.Drawing.Point(161, 45)
$CB_Holiday.Name = "CB_Holiday"
$CB_Holiday.Size = New-Object System.Drawing.Size(151, 21)
$CB_Holiday.TabIndex = 2
# I added this:
$CB_Holiday_List = "No","Yes"
$CB_Holiday.Items.AddRange($CB_Holiday_List)
$CB_Holiday.SelectedItem = "No"



### Check Boxes ###
#checkBox_AutoRun
$checkBox_AutoRun.AutoSize = $true
$checkBox_AutoRun.Location = New-Object System.Drawing.Point(150, 170)
$checkBox_AutoRun.Name = "checkBox_AutoRun"
$checkBox_AutoRun.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_AutoRun.TabIndex = 49
$checkBox_AutoRun.UseVisualStyleBackColor = $true



# checkBox_Sunday
$checkBox_Sunday.AutoSize = $true
$checkBox_Sunday.Checked = $true
$checkBox_Sunday.CheckState = [System.Windows.Forms.CheckState]::Checked
$checkBox_Sunday.Location = New-Object System.Drawing.Point(742, 227)
$checkBox_Sunday.Name = "checkBox_Sunday"
$checkBox_Sunday.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Sunday.TabIndex = 48
$checkBox_Sunday.UseVisualStyleBackColor = $true
$checkBox_Sunday.Add_CheckStateChanged({
    $text_SundayOpen.ReadOnly = $checkBox_Sunday.Checked
    $text_SundayClose.ReadOnly = $checkBox_Sunday.Checked
})

# checkBox_Monday
$checkBox_Monday.AutoSize = $true
$checkBox_Monday.Location = New-Object System.Drawing.Point(742, 252)
$checkBox_Monday.Name = "checkBox_Monday"
$checkBox_Monday.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Monday.TabIndex = 49
$checkBox_Monday.UseVisualStyleBackColor = $true
$checkBox_Monday.Add_CheckStateChanged({
    $text_MondayOpen.ReadOnly = $checkBox_Monday.Checked
    $text_MondayClose.ReadOnly = $checkBox_Monday.Checked
})

# checkBox_Tues
$checkBox_Tues.AutoSize = $true
$checkBox_Tues.Location = New-Object System.Drawing.Point(742, 277)
$checkBox_Tues.Name = "checkBox_Tues"
$checkBox_Tues.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Tues.TabIndex = 50
$checkBox_Tues.UseVisualStyleBackColor = $true
$checkBox_Tues.Add_CheckStateChanged({
    $text_TuesdayOpen.ReadOnly = $checkBox_Tues.Checked
    $text_TuesdayClose.ReadOnly = $checkBox_Tues.Checked
})

# checkBox_Wed
$checkBox_Wed.AutoSize = $true
$checkBox_Wed.Location = New-Object System.Drawing.Point(742, 302)
$checkBox_Wed.Name = "checkBox_Wed"
$checkBox_Wed.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Wed.TabIndex = 51
$checkBox_Wed.UseVisualStyleBackColor = $true
$checkBox_Wed.Add_CheckStateChanged({
    $text_WedOpen.ReadOnly = $checkBox_Wed.Checked
    $text_WedClose.ReadOnly = $checkBox_Wed.Checked
})


# checkBox_Thurs
$checkBox_Thurs.AutoSize = $true
$checkBox_Thurs.Location = New-Object System.Drawing.Point(742, 330)
$checkBox_Thurs.Name = "checkBox_Thurs"
$checkBox_Thurs.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Thurs.TabIndex = 52
$checkBox_Thurs.UseVisualStyleBackColor = $true
$checkBox_Thurs.Add_CheckStateChanged({
    $text_ThursOpen.ReadOnly = $checkBox_Thurs.Checked
    $text_ThursClose.ReadOnly = $checkBox_Thurs.Checked
})


# checkBox_Friday
$checkBox_Friday.AutoSize = $true
$checkBox_Friday.Location = New-Object System.Drawing.Point(742, 355)
$checkBox_Friday.Name = "checkBox_Friday"
$checkBox_Friday.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Friday.TabIndex = 53
$checkBox_Friday.UseVisualStyleBackColor = $true
$checkBox_Friday.Add_CheckStateChanged({
    $text_FridayOpen.ReadOnly = $checkBox_Friday.Checked
    $text_FridayClose.ReadOnly = $checkBox_Friday.Checked
})


# checkBox_Sat
$checkBox_Sat.AutoSize = $true
$checkBox_Sat.Checked = $true
$checkBox_Sat.CheckState = [System.Windows.Forms.CheckState]::Checked
$checkBox_Sat.Location = New-Object System.Drawing.Point(742, 381)
$checkBox_Sat.Name = "checkBox_Sat"
$checkBox_Sat.Size = New-Object System.Drawing.Size(15, 14)
$checkBox_Sat.TabIndex = 54
$checkBox_Sat.UseVisualStyleBackColor = $true
$checkBox_Sat.Add_CheckStateChanged({
    $text_SatOpen.ReadOnly = $checkBox_Sat.Checked
    $text_SatClose.ReadOnly = $checkBox_Sat.Checked
})



### Data Grid View ###

# dataGridView1
$dataGridView1.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridView1.Location = New-Object System.Drawing.Point(3, 280)
$dataGridView1.Name = "dataGridView1"
$dataGridView1.Size = New-Object System.Drawing.Size(808, 141)
$dataGridView1.TabIndex = 4
# I added this:
$dataGridView1.RowHeadersVisible = $false
$dataGridView1.AutoSizeColumnsMode = 'Fill'
$dataGridView1.AllowUserToResizeRows = $true
$dataGridView1.selectionmode = 'FullRowSelect'
$dataGridView1.MultiSelect = $false
$dataGridView1.AllowUserToAddRows = $false
$dataGridView1.ReadOnly = $true
$dataGridView1.ColumnCount = 8
$dataGridView1.ColumnHeadersVisible = $true
$dataGridView1.Columns[0].Name = "Ticket"
$dataGridView1.Columns[1].Name = "Title"
$dataGridView1.Columns[2].Name = "Client"
$dataGridView1.Columns[3].Name = "Contact"
$dataGridView1.Columns[4].Name = "Ticket Date"
$dataGridView1.Columns[5].Name = "Ticket Time"
$dataGridView1.Columns[6].Name = "On Call Hours"
$dataGridView1.Columns[7].Name = "On Call Tech"
$dataGridView1.Sort($dataGridView1.Columns['Ticket'],'Descending')


# dataGridView2
$dataGridView2.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridView2.Location = New-Object System.Drawing.Point(518, 37)
$dataGridView2.Name = "dataGridView2"
$dataGridView2.Size = New-Object System.Drawing.Size(280, 124)
$dataGridView2.TabIndex = 0
# I added this:
$dataGridView2.RowHeadersVisible = $false
$dataGridView2.AutoSizeColumnsMode = 'Fill'
$dataGridView2.AllowUserToResizeRows = $true
$dataGridView2.selectionmode = 'FullRowSelect'
$dataGridView2.MultiSelect = $false
$dataGridView2.AllowUserToAddRows = $true
$dataGridView2.ReadOnly = $false
$dataGridView2.ColumnCount = 2
$dataGridView2.ColumnHeadersVisible = $true
$dataGridView2.Columns[0].Name = "Tech"
$dataGridView2.Columns[1].Name = "Number"
$dataGridView2.Sort($dataGridView2.Columns['Tech'],'Ascending')




### Buttons ###

# button_save
$button_save.Location = New-Object System.Drawing.Point(528, 19)
$button_save.Name = "button_save"
$button_save.Size = New-Object System.Drawing.Size(75, 23)
$button_save.TabIndex = 6
$button_save.Text = "Save"
$button_save.UseVisualStyleBackColor = $true
# I added this:
$button_save.add_click({
    $txt_console.AppendText("$logDate - Options Saved`r`n")
    UpdateOnCallTech
    ForceHoliday
    IsItOnCall
})

# button_save2
$button_save2.Location = New-Object System.Drawing.Point(15, 11)
$button_save2.Name = "button_save2"
$button_save2.Size = New-Object System.Drawing.Size(75, 23)
$button_save2.TabIndex = 13
$button_save2.Text = "Save"
$button_save2.UseVisualStyleBackColor = $true
# I added this:
$button_save2.add_click({
    SaveConfig
    LoadConfig
    reload
})




# button_start
$button_start.Location = New-Object System.Drawing.Point(528, 51)
$button_start.Name = "button_start"
$button_start.Size = New-Object System.Drawing.Size(75, 23)
$button_start.TabIndex = 13
$button_start.Text = "Start"
$button_start.UseVisualStyleBackColor = $true
# I added this:
$button_start.add_click({
    #If on tech field empty, return error
    $tech = $CB_onCallTech.SelectedItem
    if (!$tech){
        $txt_console.AppendText("Please select an on-call technican first`r`n")
    }
    Else{
        $txt_console.AppendText("Starting Script`r`n")
        $button_stop.Enabled   = $true
        $button_start.Enabled   = $false
        UpdateOnCallTech
        ForceHoliday
        IsItOnCall
        startTimer
        startTimer2
    }
})

# button_stop
$button_stop.Location = New-Object System.Drawing.Point(528, 81)
$button_stop.Name = "button_stop"
$button_stop.Size = New-Object System.Drawing.Size(75, 23)
$button_stop.TabIndex = 14
$button_stop.Text = "Stop"
$button_stop.UseVisualStyleBackColor = $true
# I added this:
$button_stop.add_click({
    $main_loop = 0
    $txt_console.AppendText("Script Stopped`r`n")
    $button_start.Enabled   = $true
    $button_stop.Enabled   = $false
    StopTimer
    StopTimer2
})



# button_clear
$button_clear.Location = New-Object System.Drawing.Point(700, 263)
$button_clear.Name = "button_clear"
$button_clear.Size = New-Object System.Drawing.Size(90, 23)
$button_clear.TabIndex = 15
$button_clear.Text = "Clear Console"
$button_clear.UseVisualStyleBackColor = $true
# I added this:
$button_clear.add_click({
    #Function to clear Console TXT Box
})



### Text Boxes ###

# txt_console
$txt_console.Location = New-Object System.Drawing.Point(7, 19)
$txt_console.Multiline = $true
$txt_console.ScrollBars = 'Vertical'
$txt_console.Name = "txt_console"
$txt_console.ReadOnly = $true
$txt_console.Size = New-Object System.Drawing.Size(812, 273)
$txt_console.TabIndex = 0

# txt_onCallTech
$txt_onCallTech.Location = New-Object System.Drawing.Point(161, 98)
$txt_onCallTech.Name = "txt_onCallTech"
$txt_onCallTech.ReadOnly = $true
$txt_onCallTech.Size = New-Object System.Drawing.Size(151, 20)
$txt_onCallTech.TabIndex = 9

# txt_IsItOnCall
$txt_IsItOnCall.Location = New-Object System.Drawing.Point(161, 72)
$txt_IsItOnCall.Name = "txt_IsItOnCall"
$txt_IsItOnCall.ReadOnly = $true
$txt_IsItOnCall.Size = New-Object System.Drawing.Size(151, 20)
$txt_IsItOnCall.TabIndex = 7

# text_SundayOpen
$text_SundayOpen.Location = New-Object System.Drawing.Point(593, 223)
$text_SundayOpen.Name = "text_SundayOpen"
$text_SundayOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_SundayOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_SundayOpen.TabIndex = 31
$text_SundayOpen.ReadOnly = $true

# text_SundayClose
$text_SundayClose.Location = New-Object System.Drawing.Point(663, 223)
$text_SundayClose.Name = "text_SundayClose"
$text_SundayClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_SundayClose.Size = New-Object System.Drawing.Size(58, 20)
$text_SundayClose.TabIndex = 32
$text_SundayClose.ReadOnly = $true

# text_MondayOpen
$text_MondayOpen.Location = New-Object System.Drawing.Point(593, 249)
$text_MondayOpen.Name = "text_MondayOpen"
$text_MondayOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_MondayOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_MondayOpen.TabIndex = 33


# text_MondayClose
$text_MondayClose.Location = New-Object System.Drawing.Point(663, 249)
$text_MondayClose.Name = "text_MondayClose"
$text_MondayClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_MondayClose.Size = New-Object System.Drawing.Size(58, 20)
$text_MondayClose.TabIndex = 34


# text_TuesdayOpen
$text_TuesdayOpen.Location = New-Object System.Drawing.Point(593, 275)
$text_TuesdayOpen.Name = "text_TuesdayOpen"
$text_TuesdayOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_TuesdayOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_TuesdayOpen.TabIndex = 35

# text_TuesdayClose
$text_TuesdayClose.Location = New-Object System.Drawing.Point(663, 275)
$text_TuesdayClose.Name = "text_TuesdayClose"
$text_TuesdayClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_TuesdayClose.Size = New-Object System.Drawing.Size(58, 20)
$text_TuesdayClose.TabIndex = 36

# text_WedOpen
$text_WedOpen.Location = New-Object System.Drawing.Point(593, 301)
$text_WedOpen.Name = "text_WedOpen"
$text_WedOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_WedOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_WedOpen.TabIndex = 37

# text_WedClose
$text_WedClose.Location = New-Object System.Drawing.Point(663, 301)
$text_WedClose.Name = "text_WedClose"
$text_WedClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_WedClose.Size = New-Object System.Drawing.Size(58, 20)
$text_WedClose.TabIndex = 38

# text_ThursOpen
$text_ThursOpen.Location = New-Object System.Drawing.Point(593, 327)
$text_ThursOpen.Name = "text_ThursOpen"
$text_ThursOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_ThursOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_ThursOpen.TabIndex = 39

# text_ThursClose
$text_ThursClose.Location = New-Object System.Drawing.Point(663, 327)
$text_ThursClose.Name = "text_ThursClose"
$text_ThursClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_ThursClose.Size = New-Object System.Drawing.Size(58, 20)
$text_ThursClose.TabIndex = 40

# text_FridayOpen
$text_FridayOpen.Location = New-Object System.Drawing.Point(593, 353)
$text_FridayOpen.Name = "text_FridayOpen"
$text_FridayOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_FridayOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_FridayOpen.TabIndex = 41

# text_FridayClose
$text_FridayClose.Location = New-Object System.Drawing.Point(663, 353)
$text_FridayClose.Name = "text_FridayCLose"
$text_FridayClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_FridayClose.Size = New-Object System.Drawing.Size(58, 20)
$text_FridayClose.TabIndex = 42

# text_SatOpen
$text_SatOpen.Location = New-Object System.Drawing.Point(593, 379)
$text_SatOpen.Name = "text_SatOpen"
$text_SatOpen.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_SatOpen.Size = New-Object System.Drawing.Size(58, 20)
$text_SatOpen.TabIndex = 43
$text_SatOpen.ReadOnly = $true

# text_SatClose
$text_SatClose.Location = New-Object System.Drawing.Point(663, 379)
$text_SatClose.Name = "text_SatClose"
$text_SatClose.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$text_SatClose.Size = New-Object System.Drawing.Size(58, 20)
$text_SatClose.TabIndex = 44
$text_SatClose.ReadOnly = $true

# textBox1
$textBox1.Location = New-Object System.Drawing.Point(95, 278)
$textBox1.Multiline = $true
$textBox1.Name = "textBox1"
$textBox1.ReadOnly = $true
$textBox1.TextAlign = "Center"
$textBox1.Size = New-Object System.Drawing.Size(280, 98)
$textBox1.TabIndex = 56
$textBox1.Text = "Created by Robert Brown
robertdcbrown@gmail.com

Special thanks to David Long
for the PowerShell Atera Module
https://github.com/davejlong/PSAtera"


# text_AteraToken
$text_AteraToken.Location = New-Object System.Drawing.Point(153, 47)
$text_AteraToken.Name = "text_AteraToken"
$text_AteraToken.Size = New-Object System.Drawing.Size(191, 20)
$text_AteraToken.TabIndex = 4

# text_TwilioToken
$text_TwilioToken.Location = New-Object System.Drawing.Point(153, 78)
$text_TwilioToken.Name = "text_TwilioToken"
$text_TwilioToken.Size = New-Object System.Drawing.Size(191, 20)
$text_TwilioToken.TabIndex = 5

# text_TwilioSID
$text_TwilioSID.Location = New-Object System.Drawing.Point(153, 106)
$text_TwilioSID.Name = "text_TwilioSID"
$text_TwilioSID.Size = New-Object System.Drawing.Size(191, 20)
$text_TwilioSID.TabIndex = 7

# text_TwilioNumber
$text_TwilioNumber.Location = New-Object System.Drawing.Point(153, 132)
$text_TwilioNumber.Name = "text_TwilioNumber"
$text_TwilioNumber.Size = New-Object System.Drawing.Size(191, 20)
$text_TwilioNumber.TabIndex = 8






### Up/Down Interval ###

# UpDown_Interval
$UpDown_Interval.Location = New-Object System.Drawing.Point(234, 226)
$UpDown_Interval.Maximum = 60
$UpDown_Interval.Minimum = 1
$UpDown_Interval.Name = "UpDown_Interval"
$UpDown_Interval.Size = New-Object System.Drawing.Size(42, 20)
$UpDown_Interval.TabIndex = 17








### Main Window ###
# Main
$Main.Controls.Add($tabPage_Main)
$Main.Controls.Add($tabPage_Settings)
$Main.Location = New-Object System.Drawing.Point(13, 12)
$Main.Name = "Main"
$Main.SelectedIndex = 0
$Main.Size = New-Object System.Drawing.Size(828, 456)
$Main.TabIndex = 4

# tabPage_Main
$tabPage_Main.Controls.Add($box_oncall)
$tabPage_Main.Location = New-Object System.Drawing.Point(4, 22)
$tabPage_Main.Name = "tabPage_Main"
$tabPage_Main.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage_Main.Size = New-Object System.Drawing.Size(820, 430)
$tabPage_Main.TabIndex = 0
$tabPage_Main.Text = "Main"
$tabPage_Main.UseVisualStyleBackColor = $true

# tabPage_Settings
$tabPage_Settings.Controls.Add($textBox1)
$tabPage_Settings.Controls.Add($label27)
$tabPage_Settings.Controls.Add($checkBox_AutoRun)
$tabPage_Settings.Controls.Add($checkBox_Sat)
$tabPage_Settings.Controls.Add($checkBox_Friday)
$tabPage_Settings.Controls.Add($checkBox_Thurs)
$tabPage_Settings.Controls.Add($checkBox_Wed)
$tabPage_Settings.Controls.Add($checkBox_Tues)
$tabPage_Settings.Controls.Add($checkBox_Monday)
$tabPage_Settings.Controls.Add($checkBox_Sunday)
$tabPage_Settings.Controls.Add($label26)
$tabPage_Settings.Controls.Add($label25)
$tabPage_Settings.Controls.Add($label24)
$tabPage_Settings.Controls.Add($text_SatClose)
$tabPage_Settings.Controls.Add($text_SatOpen)
$tabPage_Settings.Controls.Add($text_FridayClose)
$tabPage_Settings.Controls.Add($text_FridayOpen)
$tabPage_Settings.Controls.Add($text_ThursClose)
$tabPage_Settings.Controls.Add($text_ThursOpen)
$tabPage_Settings.Controls.Add($text_WedClose)
$tabPage_Settings.Controls.Add($text_WedOpen)
$tabPage_Settings.Controls.Add($text_TuesdayClose)
$tabPage_Settings.Controls.Add($text_TuesdayOpen)
$tabPage_Settings.Controls.Add($text_MondayClose)
$tabPage_Settings.Controls.Add($text_MondayOpen)
$tabPage_Settings.Controls.Add($text_SundayClose)
$tabPage_Settings.Controls.Add($text_SundayOpen)
$tabPage_Settings.Controls.Add($LinkLabel1)
$tabPage_Settings.Controls.Add($LinkLabel2)
$tabPage_Settings.Controls.Add($label21)
$tabPage_Settings.Controls.Add($label22)
$tabPage_Settings.Controls.Add($label23)
$tabPage_Settings.Controls.Add($label20)
$tabPage_Settings.Controls.Add($label19)
$tabPage_Settings.Controls.Add($label28)
$tabPage_Settings.Controls.Add($label29)
$tabPage_Settings.Controls.Add($label30)
$tabPage_Settings.Controls.Add($label14)
$tabPage_Settings.Controls.Add($label13)
$tabPage_Settings.Controls.Add($label16)
$tabPage_Settings.Controls.Add($UpDown_Interval)
$tabPage_Settings.Controls.Add($label15)
$tabPage_Settings.Controls.Add($button_save2)
$tabPage_Settings.Controls.Add($label12)
$tabPage_Settings.Controls.Add($label11)
$tabPage_Settings.Controls.Add($text_TwilioNumber)
$tabPage_Settings.Controls.Add($text_TwilioSID)
$tabPage_Settings.Controls.Add($label10)
$tabPage_Settings.Controls.Add($text_TwilioToken)
$tabPage_Settings.Controls.Add($text_AteraToken)
$tabPage_Settings.Controls.Add($label9)
$tabPage_Settings.Controls.Add($label8)
$tabPage_Settings.Controls.Add($label7)
$tabPage_Settings.Controls.Add($dataGridView2)
$tabPage_Settings.Location = New-Object System.Drawing.Point(4, 22)
$tabPage_Settings.Name = "tabPage_Settings"
$tabPage_Settings.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage_Settings.Size = New-Object System.Drawing.Size(820, 430)
$tabPage_Settings.TabIndex = 1
$tabPage_Settings.Text = "Settings"
$tabPage_Settings.UseVisualStyleBackColor = $true


# FMmain
$FMmain.ClientSize = New-Object System.Drawing.Size(853, 794)
$FMmain.Controls.Add($Main)
$FMmain.Controls.Add($box_console)
$FMmain.Name = "FMmain"
$FMmain.Text = "Atera Ticket Text Alerts"




### Load Config ###

FirstRun
LoadConfig
reload
Autorun
Hide-Console
$FMmain.ShowDialog()

