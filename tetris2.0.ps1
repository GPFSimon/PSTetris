#
# Filename “tetrisInPowershell.ps1”
# Example for rendering screen.
#
# Original by:   Magic Mao
# Created date:  2011-07-14
# Restored and Updated by: Simon Howlett
# DateLastModified:        2022-05-11

$Host.UI.RawUI.BackgroundColor = “Black” 

$speaker = New-Object -ComObject “SAPI.SpVoice”
$voices = $speaker.GetVoices()
$winPosition = $host.ui.rawui.WindowPosition
$winSize = $host.ui.rawui.WindowSize
$host.UI.RawUI.BufferSize = New-Object -TypeName System.Management.Automation.Host.Size -Argument 150,$host.UI.RawUI.BufferSize.Height
$winSize.Height = 1500
$host.UI.RawUI.WindowSize = New-Object -TypeName System.Management.Automation.Host.Size -Argument 75,40

$blockX = 10; $blockY = 10; $blockW = 10; $blockH = 20
$nextBlockX = $blockX + $blockW + 10 
$nextBlockY = $blockY
$framework = New-Object Management.Automation.Host.Rectangle
$blockFrame =  New-Object Management.Automation.Host.Rectangle
$point = New-Object Management.Automation.Host.Coordinates
$rect = New-Object Management.Automation.Host.Rectangle

$fgc = $Host.UI.RawUI.ForegroundColor
$bgc = $Host.UI.RawUI.BackgroundColor

$border = $null # handle for border buffer
$blocks = @() # seven basic blocks
[System.ConsoleColor[]]$blockColors = @(
[System.ConsoleColor]::Gray,       #The color gray.
#[System.ConsoleColor]::Blue,        #The color blue.
[System.ConsoleColor]::Green,       #The color green.
[System.ConsoleColor]::Cyan,       #The color cyan (blue-green).
[System.ConsoleColor]::Red,        #The color red.
[System.ConsoleColor]::Magenta,    #The color magenta (purplish-red).
[System.ConsoleColor]::Yellow,      #The color yellow.
[System.ConsoleColor]::White        #The color white.
) # handle of block colors
$blockChar = “#”
$nextBlock = $null # handle for next block
[Int32]$nextBlockRotation = 0 # rotation of the next block
$nextBlockBuffer = $null # handle for next block buffer
[Int32]$nextBlockColor = -1 # next block color index
$currentBlock = $null #handle for current block
[Int32]$currentBlockRotation = 0 # rotation of the current block
$holdBlockBuffer = $null
$holdBlock = $null
$holdBlockColor = @()
$holdBlockRotation = @()
$newholdBlock = @()
[Int32]$currentBlockX = 0 
[Int32]$currentBlockY = 0 # location of current block
[Int32]$currentBlockColor = -1 # current block color index
$turndelay = 500
$currentRunFrameBuffer = $null # handle for current run block
[Int32]$totalScore = 0
$elapseTime = 0
$lineScore = 0
$levels = 1
$points = 40
$elapsed = 0
$countt = 1
$enableA = 0
$runOnce = 1


function main(){
  Introduction
  Start-Sleep -s 1
  InitEnvironment
  # Backup UI
  $script:uiBackup = $Host.UI.RawUI.GetBufferContents($framework)
  StartGame
  # Restore UI
  Draw 0 0 $uiBackup
}

#
# Brief introduction
#
function Introduction(){
  Speak 2 3 “Welcome! Enjoy the game!”
 # Speak 1 5 “Can't Stop, Won't Stop”
 # Speak 0 2 “Waddah Pat pat, with my glat patta blat blat”
}
function Speak([Int32]$index, [Int32]$rate, [String]$words){
  $speaker.Rate = $rate
  $speaker.Voice = $voices.Item($index)
  $color = switch($index){0 {“Red”} 1 {"Green"} default {“Yellow”}}
  $speaker.Speak($words, 1) | Out-Null
  Write-Host $words -ForegroundColor $color
  $speaker.WaitUntilDone(60000) | Out-Null
}

#
# Draw the formated buffer start from the [x,y]
#
function Draw($x, $y, $buffer){
  $script:elapseTime = $elapsedT.Elapsed.ToString()
  $point.x = $x 
  $point.y = $y
  $Host.UI.RawUI.SetBufferContents($point, $buffer)
}
function DrawText($x, $y, $message){
  $point.x = $x 
  $point.y = $y
  $Host.UI.RawUI.CursorPosition = $point
  Write-Host $message
}

function SelectNewBlock(){
  $script:currentBlock = $nextBlock
  $script:currentBlockRotation = $nextBlockRotation
  $script:currentBlockColor = $nextBlockColor
  $script:index = random(7)
  $script:nextBlock = $blocks[$index]
  $script:nextBlockRotation = random(4)
  switch($index){
  0 {$script:nextBlockColor = 0}
  1 {$script:nextBlockColor = 1}
  2 {$script:nextBlockColor = 2}
  3 {$script:nextBlockColor = 3}
  4 {$script:nextBlockColor = 4}
  5 {$script:nextBlockColor = 5}
  6 {$script:nextBlockColor = 6}
  }
  #$script:nextBlockColor = random($blockColors.Length)
  $nextBlockBuffer = $Host.UI.RawUI.NewBufferCellArray(@(”    “,”    “,”    “,”    “),[System.ConsoleColor]’Green’,$bgc)
  0..3 | %{$char = $nextBlock[$nextBlockRotation][$_];$nextBlockBuffer[$char[1], $char[0]] = $Host.UI.RawUI.NewBufferCellArray(@($blockChar),$blockColors[$nextBlockColor],$bgc)[0,0];}
  $script:nextBlockBuffer = $nextBlockBuffer

  $script:currentBlockX = $blockX + $blockW/2 -2 
  $script:currentBlockY = $blockY

  $script:currentRunFrameBuffer = $Host.UI.RawUI.GetBufferContents($blockFrame)
  $script:erase = $currentRunFrameBuffer.Clone()
  DrawStaticInfo
}

#
# Prepare invironment variables
#
function InitEnvironment(){
  cls
  pausemenu
  $script:elapsedT = [System.Diagnostics.Stopwatch]::StartNew()
  $script:totalScore = 0
  $script:lineScore = 0
  $script:levels = 1
  $script:turndelay = 500 
  $script:points = 40
  $script:elapsed = 0
  $script:startTime2=Get-Date

  # init framework
  $framework.Left = $framework.Top = 0
  $framework.Right =  $framework.Left + $winSize.Width
  $framework.Bottom = $framework.Top + $winSize.Height
  # build border
  [String[]]$border = @()
  $line1 = “|” 
  $line2 = “\”
  1..$blockW | % { 
    $line1 += ” “ 
    $line2 += “=” }
  $line1 += “|” 
  $line2 += “/”
  1..$blockH | %{ $border += $line1 }
  $border += $line2
  $script:border = $Host.UI.RawUI.NewBufferCellArray($border,[system.consolecolor]’Green’,$script:bgc)
  Draw $blockX $blockY $script:border
 $script:blockFrame.Left = $blockX; $script:blockFrame.Top = $blockY; $script:blockFrame.Right = $blockX + $blockW; $script:blockFrame.Bottom = $blockY + $blockH;
  #init blocks
  [Int32[][][][]]$script:blocks = @(
    @(((1,0),(1,1),(1,2),(1,3)), ((0,1),(1,1),(2,1),(3,1)), ((1,0),(1,1),(1,2),(1,3)), ((0,1),(1,1),(2,1),(3,1))),
    @(((1,1),(1,2),(2,1),(2,2)), ((1,1),(1,2),(2,1),(2,2)), ((1,1),(1,2),(2,1),(2,2)), ((1,1),(1,2),(2,1),(2,2))),
    @(((1,0),(1,1),(1,2),(2,1)), ((1,0),(0,1),(1,1),(2,1)), ((1,0),(0,1),(1,1),(1,2)), ((0,1),(1,1),(2,1),(1,2))),
    @(((1,0),(1,1),(2,1),(2,2)), ((1,1),(2,1),(0,2),(1,2)), ((1,0),(1,1),(2,1),(2,2)), ((1,1),(2,1),(0,2),(1,2))),
    @(((2,0),(1,1),(2,1),(1,2)), ((0,1),(1,1),(1,2),(2,2)), ((2,0),(1,1),(2,1),(1,2)), ((0,1),(1,1),(1,2),(2,2))),
    @(((1,0),(2,0),(1,1),(1,2)), ((0,0),(0,1),(1,1),(2,1)), ((1,0),(1,1),(0,2),(1,2)), ((0,1),(1,1),(2,1),(2,2))),
    @(((0,0),(1,0),(1,1),(1,2)), ((0,1),(1,1),(2,1),(0,2)), ((1,0),(1,1),(1,2),(2,2)), ((2,0),(0,1),(1,1),(2,1)))
    )
  #$script:blocks = $blocks
  #init current and next block
  sleep -Milliseconds 200
  SelectNewBlock
  SelectNewBlock
}

function StartGame(){
  DrawStaticInfo
  MoveBlock
}

function DrawStaticInfo(){
  #Draw $blockX $blockY $currentRunFrameBuffer
  DrawText $nextBlockX $nextBlockY (“Total Score: ” + $totalScore)
  DrawText $nextBlockX ($nextBlockY + 1)  ("Lines Cleared: " + $lineScore)
  DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels)
  DrawText $nextBlockX ($nextBlockY + 3)  ("Total Elapsed Time: " + $elapseTime)
  DrawText $nextBlockX ($nextBlockY + 5) “Next Block:”
  Draw ($nextBlockX + 6) ($nextBlockY + 6) $nextBlockBuffer
  DrawText $nextBlockX ($nextBlockY + 10) “Hold Block:”

}


function CheckAndDrawCurrentBlock(){
$ErrorActionPreference = "Stop"
  $isStuck = $false

  $buffer = $currentRunFrameBuffer.Clone()
  0..3 | %{
    $x = $currentBlockX + $currentBlock[$currentBlockRotation][$_][0] – $blockX; $y = $currentBlockY + $currentBlock[$currentBlockRotation][$_][1] – $blockY;
    $bufferCell = $buffer[$y, $x];
    if( $bufferCell.Character -ne ” ” ) { $isStuck = $true; }
    else { $bufferCell.Character = $blockChar; $bufferCell.ForegroundColor = $blockColors[$currentBlockColor]; $buffer[$y, $x] = $bufferCell } 
  }

  if ($isStuck -eq $true) {
    return $false
  } else {
    Draw $blockX $blockY $buffer
    return $true
  }
}

#
# Remove full lines
#
function ClearBlocks(){

  $buffer = $Host.UI.RawUI.GetBufferContents($blockFrame)
  $x = $blockFrame.Left 
  $y = $blockFrame.Top
  #$blockW$blockH
  $lastLine = “”
  [Int32[]]$lines = @()
  foreach($i in $blockH..0){
    $isPassed = $true
    foreach($j in 1..$blockW){
      if($buffer[$i,$j].Character -ne $blockChar) { $isPassed = $false }
    }
    if($isPassed){
      $lines += $i
    }
  }

  if($lines.Length -gt 0){
    # Blink the lines
    0..3 | %{
      foreach($i in $lines){
        foreach($j in 1..$blockW){
          $bufferCell = $buffer[$i, $j]
          $bufferCell.Character = “*”
          $buffer[$i, $j] = $bufferCell
        }
      }
      Draw $x $y $buffer
      Start-Sleep -m 100
      foreach($i in $lines){
        foreach($j in 1..$blockW){
          $bufferCell = $buffer[$i, $j]
          $bufferCell.Character = “X”
          $buffer[$i, $j] = $bufferCell
        }
      }
      Draw $x $y $buffer
      Start-Sleep -m 100
    }
    #Clear the lines
    $sourceLine = $lines[0]
    for($k = $sourceLine; $k -gt 0; $k–-){
      if($sourceLine -gt 0){
        $sourceLine–-
      }
      while($lines -Contains $sourceLine){
        if($sourceLine -gt 0){
          $sourceLine–-
        }
      }
      foreach($l in 1..$blockW){
        $bufferCell = $buffer[$k, $l]
        $bufferCell.Character = $buffer[$sourceLine, $l].Character
        $buffer[$k, $l] = $bufferCell
      }
    }
    Draw $x $y $buffer
    $script:totalScore += $points * [Math]::Pow($lines.Length,2) * $levels
    $script:lineScore += $lines.Length
    levels
  }
}


#CLEARHOLDS TOO CHANGE THE CURRENTBLOCK TO THE HOLD BLOCK!!!!~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ClearHolds(){
if($countt -ge 2){  

    $script:runOnce = 0
    $script:currentRunFrameBuffer = $Host.UI.RawUI.GetBufferContents($blockFrame)
    $buffer = $Host.UI.RawUI.GetBufferContents($blockFrame)
    $x = $blockFrame.Left 
    $y = $blockFrame.Top
    $script:currentBlockRotation = $holdBlockRotation[-2]
    $script:currentBlockColor = $holdBlockColor[-2]
    if($countt -eq 2){
        
        $script:holding = $currentBlock
        $script:currentBlock = $newholdBlock | select -last 15

        }
    if($countt -ge 3){

        $script:newholdBlock = $currentBlock
        $script:currentBlock = $holding
        
        
        $script:countt = 1
        }
    

    <#
    $newholdBlock = $Host.UI.RawUI.NewBufferCellArray(@("    ","    ","    ","    "),[System.ConsoleColor]'Green',$bgc)
    0..3 | %{
    $char2 = $newholdBlock[0][$_]
    $newholdBlock[$char2[1], $char2[0]] = $Host.UI.RawUI.NewBufferCellArray(@($blockChar),$blockColors[$currentBlockColor],$bgc)[0,0] 
  }
  #>

    Draw $x $y $buffer

}
}

function KeyToCommand(){
  $code = 0
      #$Host.UI.RawUI.FlushInputBuffer()
      #Start-Sleep -milliseconds 50
  if ($Host.UI.RawUi.KeyAvailable){
    $key = $Host.UI.RawUI.ReadKey(“NoEcho, IncludeKeyUp”)
    $code = switch($key.VirtualKeyCode){
      81 { -1 } # ‘q’ quit the game
      32 { 32 } # map space to down arrow
      80 {  $elapsedT.Stop()
            pausemenu
            $elapsedT.Start()
         }
      65 { If($enableA -eq 0) {
            HoldBlocks #press A
            }
         }
      {(37..40) -contains $_ } {$key.VirtualKeyCode} # arrow key
      default { 0 } # do nothing
    }
  }
  return $code
}

#HOLDBLOCKS @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function HoldBlocks(){
    
    Draw $blockX $blockY $erase

   
    $script:newholdBlock += $currentBlock
    $script:holdBlockRotation += $currentBlockRotation
    $script:holdBlockColor += $currentBlockColor

    $holdBlock = $Host.UI.RawUI.NewBufferCellArray(@("    ","    ","    ","    "),[System.ConsoleColor]'Green',$bgc)
    0..3 | %{   
        $char2 = $currentBlock[$currentBlockRotation][$_]
        $holdBlock[$char2[1], $char2[0]] = $Host.UI.RawUI.NewBufferCellArray(@($blockChar),$blockColors[$currentBlockColor],$bgc)[0,0]
            } 
    Draw ($nextBlockX + 11) ($nextBlockY + 11) $holdBlock

    $script:currentBlockX = $blockX + $blockW/2 -2 
    $script:currentBlockY = $blockY


    ClearHolds
    If ($runOnce -eq 1) { SelectNewBlock }

        
   
    $script:currentRunFrameBuffer = $Host.UI.RawUI.GetBufferContents($blockFrame)
    $script:countt++
    $script:enableA = 1
}

<#
function Rotate($tempBlock){
  $tempBlock
 0..3 | %{
    $y = $tempBlock[$_][0] – 1; $x = $tempBlock[$_][1] – 1;
    $x = $x -BXOR $y; $y = $x -BXOR $y; $x = $x -BXOR $y;
    if(((1 – $y) -lt 0) -or (($x + 1) -lt 0)) {exit; 1 – $y; $x + 1;}
    $tempBlock[$_][0] = 1 – $y; $tempBlock[$_][1] = $x + 1
    }

  return $tempBlock
}#>

function MoveBlock(){
  #$leftKey = 37 $upKey = 38 $rightKey = 39 $downKey = 40
  $code = KeyToCommand
  $startTime=Get-Date
  while($code -ne -1){
    do{
      $isDown = $false
      $code = KeyToCommand
      $x = $currentBlockX
      $lastBlockRotation = $currentBlockRotation
      switch($code){
        37 { $script:currentBlockX-– } #<=
        38 { $script:currentBlockRotation = ($currentBlockRotation + 1) % 4 } #rotate
        39 { $script:currentBlockX++ } #=>
        40 { $isDown = $true} #striaght down
        32 { $isDown = $true
          $result = CheckAndDrawCurrentBlock
          $y = $currentBlockY
          while($result){
            $y = $script:currentBlockY++
            $result = CheckAndDrawCurrentBlock
            }
          $script:currentBlockY = $y
          }
        default {}
      }

      $result = CheckAndDrawCurrentBlock
      if($result -ne $true){
        $script:currentBlockX = $x
        $script:currentBlockRotation = $lastBlockRotation
      }
      $elapsed=((Get-Date).Subtract($startTime)).TotalMilliseconds
      [int]$secs = ((Get-Date).Subtract($startTime2)).TotalSeconds
      DrawText $nextBlockX ($nextBlockY + 3)  ("Total Elapsed Time: " + $elapseTime)
      if($isDown -OR (($isDown -eq $false) -AND ($elapsed -ge $turndelay))){
        $isDown = $true 
        $script:currentBlockY++ 
        $startTime=Get-Date
        $result = CheckAndDrawCurrentBlock
        if($result -ne $true){
          ClearBlocks
          #Score for placing down peieces
          if($secs -lt 60) {$script:totalScore += (60 - $secs)}
          $script:enableA = 0
          SelectNewBlock
          $startTime2=Get-Date
          $result = CheckAndDrawCurrentBlock
          if($result -ne $true) {
            break
          }
        }
      }
      Start-Sleep -m 10
    } until($code -eq -1)

    #End game here
    $highscorelow = 1000
    $highscorelist = "$PSScriptRoot\highscore.csv"
    if (!(Test-Path $highscorelist))
    {
$highscorefile = @"
"Name","Total Score","Lines Cleared","Level"
"Need","25000","999","99"
"At","20000","999","99"
"Least","15000","999","99"
"1000","10000","999","99"
"Score","5000","999","99"
"@
    $highscorefile | Out-File $highscorelist -Encoding ascii
    }

    $highscores = Import-Csv $highscorelist
    $Headers = "Name,Total Score,Lines Cleared,Level"
    

    If($totalScore -gt $highscorelow){
        write-host "Congrats!" -ForegroundColor Green
        sleep 1
        $highscoreName = read-host "Highscore! Please enter your name"
        $highscoreNew = "$highscoreName,$totalScore,$lineScore,$levels"
        $highscoreNew | Out-File $highscorelist -Encoding ascii -Append
        $highscores = Import-Csv $highscorelist
        }

    $highscores | Sort-Object {[int]$_."total score"} -Descending | Export-Csv -path $highscorelist -NoTypeInformation
    $highscores = Import-Csv $highscorelist

#Get cursor back at the bottom
    "`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n"
    $highscores | select -first 10 | Format-Table 
    $continue = read-host "Type: 1 for full Highscore list`nQ to quit`nOr any key to Continue... "
    If($continue -eq "1") {$highscores | more; $Host.UI.RawUI.ReadKey(“NoEcho, IncludeKeyUp”)}
    elseif($continue -eq "q") {cls;exit}
    elseif($continue -eq "echo") {echo1}
    else { InitEnvironment }
    #CheckAndDrawCurrentBlock

  }   # end of main loop
  #exit 1
  return
}

function levels(){
    switch($lineScore) {
    {$_ -ge 10}{ $script:levels = 1; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 500; $script:points = 45 }
    {$_ -ge 20}{ $script:levels = 2; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 400; $script:points = 50 }
    {$_ -ge 30}{ $script:levels = 3; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 400; $script:points = 80 }
    {$_ -ge 40}{ $script:levels = 4; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 300; $script:points = 83 }
    {$_ -ge 50}{ $script:levels = 5; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 300; $script:points = 120 }
    {$_ -ge 60}{ $script:levels = 6; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 200; $script:points = 123 }
    {$_ -ge 70}{ $script:levels = 7; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 200; $script:points = 160 }
    {$_ -ge 80}{ $script:levels = 8; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 180; $script:points = 163 }
    {$_ -ge 90}{ $script:levels = 9; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels); $script:turndelay = 150; $script:points = 200 }
    {$_ -ge 100}{ $script:levels = 10; DrawText $nextBlockX ($nextBlockY + 2)  ("Current Level: " + $levels +"(Max)"); $script:turndelay = 100; $script:points = 202}
    }
}

function echo1() {
$done = "Y"

while ($done -eq "Y") {
    $string = read-host "What do you want to say"
    $voice = read-host "In what voice? [0-2]"
    Speak $voice 2 $string

    $done = read-host "Again? (Y/N)"
    }
    InitEnvironment
}

function pausemenu(){
$a = new-object -comobject wscript.shell
$b = $a.popup(“Controls`n---------`nArrows keys:`n`tLeft/Right to move`n`tUp to Rotate`n`tDown to increase drop speed `nA to Hold piece `nP to Pause `nSpacebar to Fast drop`nQ to Quit`nType Echo on menu for fun`n`nPress OK to Continue...“,0,“Pause Menu”,0)
}

. main