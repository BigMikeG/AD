; Auto DPS archive generator
; Mike G.
; The application automatically creates DPS Tool archives. You must create a configuration file 
; that describes the archive that you want to create (name and contents).
;
; Year/Vehicle/SW_Level/Country/Comment,HW Code,Rev,System,Display,Wireless,Nav System,Navn Map DB,Brand,Country,Speech Rec,Hands Free,Rear Camera
; HMI14_Y1XX_N1_USA,BC00,Rev1,22906036,22906019,22906002,22906000,22905992,22905990,12031201,22905975,22905966,22905950
; HMI14_Y1XX_N1_Korea,B400,Rev1,22906029,22906007,22906002,22905998,22905992,22905990,22905984,22905968,22905966,22905950
;
; Notes:
;  The ini file is saved when the Go button is pressed and all the inputs are valid.
;  The FileExists function seems to be the biggest time drag. It is used to verify every binary file. 

#include <GUIConstantsEx.au3>

Opt("MustDeclareVars", 1) ; 1=Variables must be pre-declared, 0=Variables don't need to be pre-declared 

Const $VER                       = "12.04.10"
Const $DEAD_TIME                 = 0

Const $SEC_1                     = 1
Const $ERROR_TMO                 = 30   ; my display fails but recovers sometimes, but it takes more than 5 seconds
Const $MS_5000                   = 5000
Const $ARCHIVE_COMMENTS_LINE1_ID = 1004
Const $ARCHIVE_COMMENTS_LINE2_ID = 1005
Const $ARCHIVE_COMMENTS_LINE3_ID = 1006
Const $ARCHIVE_COMMENTS_LINE4_ID = 1007
Const $NEW_ARCHIVE_NAME_ID       = 1009
Const $TEMPLATE_FILENAME_ID      = 1023
Const $NEW_ARCHIVE_FOLDER_ID     = 1025
Const $BUILD_CAL_ARCHIVE_ID      = 1065
Const $OPEN_FILENAME_INPUT_ID    = 1048
Const $CAL_OFFSET                = 4    ; column number in the config file where the parts start
Const $X_INIT                    = 564  ; Set the x and y coordinates where the "Select File" button is located.
Const $Y_INIT                    = 13   
Const $Y_OFFSET                  = 20   ; offset down to the "Select File" button for the next part

Const $DPS_TOOL    = "DPS_TOOL"
Const $TEMPLATE    = "TEMPLATE"
Const $CONFIG      = "CONFIG"
Const $UTILITY     = "UTILITY"
Const $BINARY_PATH = "BINARY_PATH"
Const $OUTPUT_PATH = "OUTPUT_PATH"
Const $OVERWRITE   = "OVERWRITE"
Const $INI_FILE    = "C:\Temp\AutoDPS.ini"
Const $LOG_FILE    = "AutoDpsLog.txt"

Global $dpsTool      = ""
Global $templateFile = ""
Global $configFile   = ""
Global $utilityFile  = ""
Global $binaryPath   = ""
Global $outputPath   = ""
Global $lineNumber   = 1
Global $numArchivesCreated = 0
Global $numLogMessages = 0

Opt("GUIOnEventMode", 1)  ; Change to OnEvent mode 
GUICreate("Auto DPS Ver " & $VER, 450, 190)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;                                           x    y
GUICtrlCreateLabel("DPS Tool Executable:",  10,  14)
GUICtrlCreateLabel("DPS Template File:",    10,  34)
GUICtrlCreateLabel("Auto DPS Config File:", 10,  54)
GUICtrlCreateLabel("Utility File:",         10,  74)
GUICtrlCreateLabel("Binary Path:",          10,  94)
GUICtrlCreateLabel("Output Path:",          10,  114)

LoadIniFileSettings()

Local $dpsInputId       = GUICtrlCreateInput($dpsTool,      120,  10, 300, 20)
Local $templateInputId  = GUICtrlCreateInput($templateFile, 120,  30, 300, 20)
Local $configInputId    = GUICtrlCreateInput($configFile,   120,  50, 300, 20)
Local $utilityInputId   = GUICtrlCreateInput($utilityFile,  120,  70, 300, 20)
Local $binaryInputId    = GUICtrlCreateInput($binaryPath,   120,  90, 300, 20)
Local $outputInputId    = GUICtrlCreateInput($outputPath,   120, 110, 300, 20)
Local $overwriteChkbxId = GUICtrlCreateCheckbox("Overwrite Existing Archives", 10, 130)

Local $dpsButton        = GUICtrlCreateButton("...", 422,  10, 20, 20)
Local $templateButton   = GUICtrlCreateButton("...", 422,  30, 20, 20)
Local $configButton     = GUICtrlCreateButton("...", 422,  50, 20, 20)
Local $utilityButton    = GUICtrlCreateButton("...", 422,  70, 20, 20)
Local $binaryButton     = GUICtrlCreateButton("...", 422,  90, 20, 20)
Local $outputButton     = GUICtrlCreateButton("...", 422, 110, 20, 20)

Local $helpButton       = GUICtrlCreateButton("Help",  10, 160, 60)
Local $goButton         = GUICtrlCreateButton("Go",   380, 160, 60)

GUICtrlSetOnEvent($dpsButton,      "DpsToolButton")
GUICtrlSetOnEvent($templateButton, "TemplateButton")
GUICtrlSetOnEvent($configButton,   "ConfigButton")
GUICtrlSetOnEvent($utilityButton,  "UtilityButton")
GUICtrlSetOnEvent($binaryButton,   "BinaryButton")
GUICtrlSetOnEvent($outputButton,   "OutputButton")
GUICtrlSetOnEvent($helpButton,     "HelpButton")
GUICtrlSetOnEvent($goButton,       "GoButton")

GUISetState(@SW_SHOW)

While 1
  Sleep(1000)  ; Idle around
WEnd

Func DpsToolButton()
   $dpsTool = FileOpenDialog ("Select the DPS Tool executable (dps.exe) to use", "", "(*.exe)", 1)
   GUICtrlSetData($dpsInputId, $dpsTool)
EndFunc

Func TemplateButton()
   $templateFile = FileOpenDialog ("Select the Template file to use", "", "Template files (*.dtm)", 1)
   GUICtrlSetData($templateInputId, $templateFile)
EndFunc

Func ConfigButton()
   $configFile = FileOpenDialog ("Select the Archive Config File to use", "", "Archive Config files (*.csv)", 1)
   GUICtrlSetData($configInputId, $configFile)
EndFunc

Func UtilityButton()
   $utilityFile = FileOpenDialog ("Select the Utility file to use", "", "Utility files (*.*)", 1)
   GUICtrlSetData($utilityInputId, $utilityFile)
EndFunc

Func BinaryButton()
   $binaryPath  = FileSelectFolder("Select the folder where your binary files are located.", "", 4, "")
   GUICtrlSetData($binaryInputId, $binaryPath)
EndFunc

Func OutputButton()
   $outputPath = FileSelectFolder("Select the output folder where your archives will be stored.", "", 4, "")
   GUICtrlSetData($outputInputId, $outputPath)
EndFunc

; Launch the help site in a web browser
Func HelpButton2()
   MsgBox(0, "Help", "Help is currently unavailable until the Share Point site is up and running. Sorry.")
EndFunc

Func HelpButton()
   Local $username = EnvGet("USERNAME")
   Local $chrome   = "C:\Users\" & $username & "\AppData\Local\Google\Chrome\Application\chrome.exe" 
   Local $explorer = "C:\Program Files\Internet Explorer\iexplore.exe" 
   ;Local $site     = " https://sites.google.com/site/autodpstool/user-s-guide"
   Local $site     = " https://gmweb.gm.com/sites/CalSupport/Auto%20DPS"
   
   If FileExists($chrome) Then
	  Run($chrome & $site)
   ElseIf FileExists($explorer) Then
	  Run($explorer & $site)
   Else
	  MsgBox(0, "Unable To Launch Web Browser", "Help is here:" & $site)
   EndIf
EndFunc

Func GoButton()
   Local $time = TimerInit() ; start a timer to see how long it takes to build the archives
   Local $timePerArchive
   Local $line2 = ""
   Local $line3 = " Review file '" & $outputPath & "\" & $LOG_FILE & "' for any issues."
   
   $numArchivesCreated = 0

   ; Validate the GUI input boxes before continuing.
   if AreGuiInputsValid() Then
	  SaveIniFileSettings()
	  LaunchDpsTool()
	  ArchiveHandler()
	  WinClose("Development Programming System")

	  ; Calculate how long it took to create the archives.
	  Local $dif = Round((TimerDiff($time) / 1000), 0)
	  Local $min = Floor($dif / 60)
	  Local $sec = $dif - ($min * 60)
	  
	  If $min > 0 Then
		 Local $line1 = $numArchivesCreated & " Archive(s) Created in " & $min & " minutes and " & $sec & " seconds."
	  Else
		 Local $line1 = $numArchivesCreated & " Archive(s) Created in " & $sec & " seconds."
	  Endif
	  
	  If $numArchivesCreated <> 0 Then
		 $timePerArchive = $dif / $numArchivesCreated
		 $line2 = " " & $timePerArchive & " seconds per archive."
	  EndIf
	  
	  Local $msg = $line1 & $line2 & $line3
	  
	  ; If any messages were logged, open the log file in Notepad.
	  If $numLogMessages <> 0 Then
		 Run("notepad.exe " & $outputPath & "\" & $LOG_FILE)
	  EndIf
	  
	  MsgBox(0, "DPS Archive Generator", $msg)
   EndIf ; if AreGuiInputsValid() Then
EndFunc

Func CLOSEClicked()
   WinClose("Development Programming System")
   Exit
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This functions verifies that all of the GUI inputs exist.
; Returns 1 if everything is valid, 0 otherwise.
;//////////////////////////////////////////////////////////////////////////////
Func AreGuiInputsValid()
   ; Read the edit boxes
   $dpsTool      = GUICtrlRead($dpsInputId)
   $templateFile = GUICtrlRead($templateInputId)
   $configFile   = GUICtrlRead($configInputId)
   $utilityFile  = GUICtrlRead($utilityInputId)
   $binaryPath   = GUICtrlRead($binaryInputId)
   $outputPath   = GUICtrlRead($outputInputId)

   ; Validate
   If FileExists($dpsTool) = 0 Then
	  MsgBox(0, "Error", "Please select a valid DPS Tool")
	  Return 0
   ElseIf FileExists($templateFile) = 0 Then
	  MsgBox(0, "Error", "Please select a valid template file")
	  Return 0
   ElseIf FileExists($configFile) = 0 Then
	  MsgBox(0, "Error", "Please select a valid config file")
	  Return 0
   ElseIf FileExists($utilityFile) = 0 Then
	  MsgBox(0, "Error", "Please select a valid utility file")
	  Return 0
   ElseIf FileExists($binaryPath) = 0 Then
	  MsgBox(0, "Error", "Please select a valid binary folder")
	  Return 0
   ElseIf FileExists($outputPath) = 0 Then
	  MsgBox(0, "Error", "Please select a valid output folder")
	  Return 0
   Else
	  Return 1
   EndIf
EndFunc

; This function reads the settings from the AutoDPS.ini and set the global variables.
Func LoadIniFileSettings()
   Local $lineNumber = 1
   Local $line
   Local $array

   Local $file = FileOpen($INI_FILE, 0)

   ; Check if file opened for reading OK
   If $file <> -1 Then
	  $line = FileReadLine($file, $lineNumber) ; Read a line from the config file
	  
	  ; Read lines until there are none left
	  While $line <> ""
		 ; split on the equal sign and save each field in an array of strings
		 $array = StringSplit($line, "=")

		 ; Verify that there are 2 elements in the line. 
		 If $array[0] = 2 Then
			SetIniParameter($array)
		 EndIf
		 
		 $lineNumber += 1                              ; inc the config file line number
		 $line = FileReadLine($file, $lineNumber)   ; read the next line in the config file

	  WEnd ; End of the main while loop that processes archive records
   EndIf
   
   ; Close the ini file.
   FileClose($file)
EndFunc

; This function parses a line from the ini file and sets variables that will be copied into
; the input boxes on the GUI.
Func SetIniParameter($array)
   Switch $array[1]
	  Case $DPS_TOOL
		 $dpsTool = $array[2]
	  Case $TEMPLATE
		 $templateFile = $array[2]
	  Case $CONFIG
		 $configFile = $array[2]
	  Case $UTILITY
		 $utilityFile = $array[2]
	  Case $BINARY_PATH
		 $binaryPath = $array[2]
	  Case $OUTPUT_PATH
		 $outputPath = $array[2]
   EndSwitch
EndFunc

Func SaveIniFileSettings()
   Local $handle = FileOpen ($INI_FILE, 2 ) ; open for Write Mode (erase previous contents) 
   
   ; Check if file opened for writing OK
   If $handle = -1 Then
	  MsgBox(0, "Error", "Unable to write ini file: " & $INI_FILE)
   Else
	  FileWriteLine($handle, $DPS_TOOL    & "=" & $dpsTool)
	  FileWriteLine($handle, $TEMPLATE    & "=" & $templateFile)
	  FileWriteLine($handle, $CONFIG      & "=" & $configFile)
	  FileWriteLine($handle, $UTILITY     & "=" & $utilityFile)
	  FileWriteLine($handle, $BINARY_PATH & "=" & $binaryPath)
	  FileWriteLine($handle, $OUTPUT_PATH & "=" & $outputPath)
	  
	  FileClose($handle)
   EndIf
EndFunc

; This function shuts down the DPS tool when an error is detected.
Func ExitGracefully($error)
   WinClose("Development Programming System")
   MsgBox(0, "Error", $error)
EndFunc

; This function shuts down everything when a fatal error is detected.
Func FatalError($error)
   ExitGracefully($error)
   Exit
EndFunc

; Launch the DPS Tool and get it to the Create New Archive window.
Func LaunchDpsTool()
   Local $rv;
   
   ; Start the Development Programming System tool.
   If Run($dpsTool) = 0 Then
	  ExitGracefully("Unable to launch DPS Tool: '" & $dpsTool & "'. Please verify it exists.")
   Else
	  ; Wait for the DPS tool to open.
	  $rv = WinWaitActive("Development Programming System", "", $ERROR_TMO)
	  If $rv = 0 Then
		 ExitGracefully("DPS tool failed to become active.")
	  Else
		 ; Set the option to do a sub-string title match.
		 Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase

		 ; Send keys "Alt-a" then "s" to open the Spat Command Center.
		 ClickSpatCommandCenterAndVerify()

		 ; Click on the "Create a New Archive" button.
		 ClickNewArchiveAndVerify()

		 ; Paste the Output Path into the "Folder Where New Archive Will Be Stored" input box.
		 SetEditBoxNoVerify("New Archive Utility", $NEW_ARCHIVE_FOLDER_ID, $outputPath)
		 
		 LoadTemplateFile()
	  EndIf ; If $rv = 0 Then
   EndIf ; If Run($dpsTool) = 0 Then
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; Click on the "Create a New Archive" button.
;//////////////////////////////////////////////////////////////////////////////
Func ClickNewArchiveAndVerify()
   Local $rv = 0
   Local $elapsed = 0
   Local $start = TimerInit() ; start a timer to break us out or the loop
   
   Sleep($DEAD_TIME) ; pause for the window needs a little time to settle
   
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ControlClick("Spat","Create a New Archive",1080)
  	  $rv = WinWaitActive("New Archive Utility", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd

   ; Were we unsuccessful in opening the "New Archive Utility" window?
   If $rv = 0 Then
  	  ExitGracefully("Failed to opem the New Archive Utility window.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function sends keys "Alt-a" then "s" to open the Spat Command Center.
;//////////////////////////////////////////////////////////////////////////////
Func ClickSpatCommandCenterAndVerify()
   Local $rv = 0
   Local $elapsed = 0
   Local $start = TimerInit() ; start a timer to break us out or the loop
   
   Sleep($DEAD_TIME) ; the window needs a little time to settle
   
   While (($rv = 0) And ($elapsed < $MS_5000))
	  Send("!a") ; Alt-A 
	  Send("s")  ; s
  	  $rv = WinWaitActive("Spat1", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd

   ; Were we unsuccessful in opening the Spat1 window?
   If $rv = 0 Then
	  ExitGracefully("Failed to opem the Spat1 window.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function creates an archive for each row in the config file.
;//////////////////////////////////////////////////////////////////////////////
Func ArchiveHandler()
   Local $archiveFileName = "";
   Local $lineNumber
   Local $line
   Local $numCalsets
   Local $array
   Local $logFile = $outputPath & "\" & $LOG_FILE

   Local $cfgFh = FileOpen($configFile, 0)
   Local $logFh = FileOpen($logFile, 2) ; open for Write Mode (erase previous contents) 

   ; Check if the log file opened for writing OK
   If $logFh = -1 Then
	  MsgBox(0, "Error", "Unable to write log file: " & $LOG_FILE)
   EndIf

   ; Check if file opened for reading OK
   If $cfgFh = -1 Then
	  ExitGracefully("Unable to open file: " & $configFile)
   Else
	  ; Read the header file to get the number of calsets.
	  $line = FileReadLine($configFile, 1)  ; Read a line from the config file
	  $line = StringReplace($line, '"', '') ; remove adouble qoutes
	  $array = StringSplit($line, ",")      ; split on the comma and save each field in an array of strings
	  $numCalsets = $array[0] - 3           ; number of columns minus the 3 used for the archive name
	  
	  ; If the header row ends with a comma, don't count the last column. 
	  if $array[$array[0]] == "" Then
		 $numCalsets--
	  EndIf
	  
	  ; Read the first archive row
	  $lineNumber = 2                                ; first line is a header, start at line 2
	  $line = FileReadLine($configFile, $lineNumber) ; Read a line from the config file
	  $line = StringReplace($line, '"', '')          ; remove adouble qoutes

	  ; One arch is created each time through the loop.
	  While $line <> ""
		 $array = StringSplit($line, ",") ; split on the comma and save each field in an array of strings
		 
		 ; Verify that the number of calsets in the line matches the header file. 
		 If $array[0] >= ($numCalsets + $CAL_OFFSET - 1) Then
			; Get the archive file name.
			$archiveFileName = ValidateArchiveFileName($array, $lineNumber, $logFh);
			
			; If archive file name is not blank, process the archive.
			if $archiveFileName <> "" Then
			   ; Build full archive name with the path.
			   Local $fullArchiveName = $outputPath & "\" & $archiveFileName & ".zip"

			   ; Create archives as long as there are no blank parts in the line.
			   If IsAnyPartBlank($array, $numCalsets) Then
				  FileWriteLineToLog($logFh, "Archive not created for line " & $lineNumber & " of '" & $configFile & "' because missing part detected.")
			   Else  
				  ; Create the archive if it doesn't exist or the overwrite checkbox is checked.
				  If (FileExists($fullArchiveName) = 0) Or (IsOverwriteChecked()) Then
					 CreateArchive($archiveFileName, $array, $lineNumber, $numCalsets)
				  EndIf
			   EndIf ; If IsAnyPartBlank($array) Then
			EndIf ; if $archiveFileName <> "" Then
			
			$lineNumber += 1                               ; inc the config file line number
			$line = FileReadLine($configFile, $lineNumber) ; read the next line in the config file
			$line = StringReplace($line, '"', '')          ; remove adouble qoutes
		 Else
			ExitGracefully("Config file: " & $configFile & ", Line: " & $lineNumber & " - Too few part numbers compared to header row.")
			ExitLoop
		 EndIf ; If $array[0] = ($numCalsets + $CAL_OFFSET - 1) Then
	  WEnd ; End of the main while loop that processes archive records
   EndIf ; If $file = -1 Then

   ; Close the Archive Configuration file and log file.
   FileClose($cfgFh)
   FileClose($logFh)
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This functions validates the 3 columns in the config file that form the archive name.
; Name must be 32 chars or less.
; First column must not be blank.
; If HW Code or Rev are blank don't insert the underscores.
;//////////////////////////////////////////////////////////////////////////////
Func ValidateArchiveFileName($array, $lineNumber, $logFh)
   Local $name = ""
   
   ; First column can't be blank.
   If $array[1] = "" Then
	  FileWriteLineToLog($logFh, "Archive not created for line " & $lineNumber & " of '" & $configFile & "' because the first column can't be blank.")
   Else
	  $name = $array[1]
	  
	  ; Check if something is in the HW Code column.
	  If $array[2] <> "" Then
		 Local $s = $name & "_" & $array[2] ; insert an underscore and append to the archive name
		 
		 ; Auto-correct the Special Number conversion issue.
		 ; Excel does a conversion on numbers that have an E in the second character.
		 ; For example, 4E20 becomes 4.00E+20 in Excel and 4.00E+020 in Open Office.
		 $name = StringRegExpReplace($s, '\.00E\+.*(\d{2}$)', 'E$1')
		 
		 ; Send a message to the log file if a replace occurred.
		 If @error = 0 And @extended = 1 Then
			FileWriteLineToLog($logFh, "Archive file name was auto-corrected on line " & $lineNumber & " of '" & $configFile & ".")
		 EndIf
	  EndIf
		 
	  ; Check if something is in the Rev column.
	  If $array[3] <> "" Then
		 ; Append it to the filename with an underscore in front.
		 ; No longer adding Rev to it. Let the user put it in the column if they want do it.
		 $name &= "_" & $array[3]
	  EndIf

	  ; Archive file name must be less than 32 characters
	  If StringLen($name) > 32 Then
		 FileWriteLineToLog($logFh, "Archive not created for line " & $lineNumber & " of '" & $configFile & "' because file name '" & $name & "' must be <= 32 characters.")
		 $name = ""
	  EndIf
   EndIf ; If $array[1] = "" Then
   
   Return $name;
EndFunc

Func FileWriteLineToLog($logFh, $s)
   FileWriteLine($logFh, $s)
   $numLogMessages += 1
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function returns the status of the "Overwrite Existing Archives" checkbox.
; 1 = Checked, 0 = Not Checked.
;//////////////////////////////////////////////////////////////////////////////
Func IsOverwriteChecked()
   Local $rv
   
   If BitAnd(GUICtrlRead($overwriteChkbxId),$GUI_CHECKED) Then
	  Return 1 ; overwrite existing archives
   Else
	  Return 0 ; do not overwrite existing archives
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function checks the parts in a row to make sure they are not blank.
;//////////////////////////////////////////////////////////////////////////////
Func IsAnyPartBlank($array, $numCalsets)
   Local $rv = 0
   Local $i 
   
   For $i = 0 to ($numCalsets - 1)
      If $array[$i + $CAL_OFFSET] == "" Then
         $rv = 1
         ExitLoop
      EndIf
   Next
   
   Return $rv
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function creates an archive.
;//////////////////////////////////////////////////////////////////////////////
Func CreateArchive($file, $array, $lineNumber, $numCalsets)
   Local $rv = 0
   Local $start
   Local $elapsed = 0
   Local $y
   
   ; Verify that the Service window is active before starting.
   CheckForSpatWindowActive()
   
   ; Paste the new archive filename into the edit box. 
   SetEditBoxNoVerify("Service Programming Archive Tool(SPAT)", $NEW_ARCHIVE_NAME_ID, $file)
   
   ; Click the "Next" button to select the archive parts.
   ClickTheNextButtonAndVerify()
   
   ; Select the utility file for the archive.
   $y = $Y_INIT                     ; init the y coordinate
   SelectBinaryFile($y, "", $utilityFile, $lineNumber)

   ; Now select each of the 10 cal sets to add to the archive.
   For $i = 0 to ($numCalsets - 1)
      $y += $Y_OFFSET   ; move the y coordinate down to the next file
      SelectBinaryFile($y, ($binaryPath & "\"), $array[$i + $CAL_OFFSET], $lineNumber)
   Next

   ; Click the Build button and wait for the Archive Successful message.
   $rv = 0
   $elapsed = 0
   $start = TimerInit() ; start a timer to break us out or the loop
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ControlClick("Build Calibration  Archive/Template File", "Build", 1)
	  
	  ; Wait 1 sec for the Zip File Already Exists message. If it appears click OK to close it.
	  ; The Message won't show if the "Overwite Existing Archives" checkbox is not checked.
	  If (IsOverwriteChecked() = 1) Then
		 $rv = WinWaitActive("ZIP FILE ERROR", "", $SEC_1)
		 If $rv <> 0 Then
			ControlClick("ZIP FILE ERROR", "OK", 2)
		 EndIf
	  EndIf

	  ; Click OK when the "Archive File Was Created Successfully" message appears. 
	  $rv = CloseArchiveFileWasCreatedSuccessfullyBox()
	  $elapsed = TimerDiff($start)
   WEnd

   ; Click the Cancel button to return to the "Service Programming Archive Tool(SPAT)" window.
   $rv = 0
   $elapsed = 0
   $start = TimerInit() ; start a timer to break us out or the loop
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ControlClick("Build Calibration  Archive/Template File", "Cancel", 2)
  	  $rv = WinWaitActive("Service Programming Archive Tool(SPAT)", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd

   ; Were we unsuccessful in returning to the SPAT window?
   If $rv = 0 Then
	  FatalError("Failed to return to the Service Programming Archive Tool window. Exiting.")
   EndIf
   
   $numArchivesCreated += 1 ; increment the number of archives created.
EndFunc


;//////////////////////////////////////////////////////////////////////////////
; This function clicks the Next button and verifies that the Build Calibration 
; window has opened.
;//////////////////////////////////////////////////////////////////////////////
Func ClickTheNextButtonAndVerify()
   Local $rv = 0
   Local $start
   Local $elapsed = 0
   
   $start = TimerInit() ; start a timer to break us out or the loop
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ControlClick("Service Programming Archive Tool(SPAT)", "", "[CLASS:Button; TEXT:Next; INSTANCE:1]")
  	  $rv = WinWaitActive("Build Calibration  Archive/Template File", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd

   ; Were we unsuccessful in opening the Build Calibration window?
   If $rv = 0 Then
	  FatalError("Failed to open the Build Calibration window. Exiting.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function closes the "Archive File Was Created Successfully" message box 
; Returns 1 if successful.
;//////////////////////////////////////////////////////////////////////////////
Func CloseArchiveFileWasCreatedSuccessfullyBox()
   Local $rv = WinWaitActive("dps", "Archive File Was Created Successfully", $ERROR_TMO)
   
   If $rv = 0 Then
	  Return 0
   Else
      ControlClick("dps", "OK", 2)
	  Return 1
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function loads the template file.
;//////////////////////////////////////////////////////////////////////////////
Func LoadTemplateFile()
   Local $rv = 0
   Local $start
   Local $elapsed = 0
   
   ; Click the Template button and wait for the Open window to display.
   $start = TimerInit() ; start a timer to break us out or the loop
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ; Click the template button. The window title changes from New to Service after the first click. 
	  ControlClick("New Archive Utility", "", "[CLASS:Button; TEXT:Template; INSTANCE:1]")
	  ;ControlClick("Service Programming Archive Tool(SPAT)", "", "[CLASS:Button; TEXT:Template; INSTANCE:1]")
	  $rv = WinWaitActive("Open", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd

   ; Did the Open window open?
   If $rv <> 0 Then
	  ; Return was successful. Paste the template filename into the filename box.
	  SetEditBoxNoVerify("Open", "[CLASS:Edit; INSTANCE:1]", $templateFile)

	  ; Click the "Open" button to open the template file and verify the we have returned to the 
	  ; "Service Programming Archive Tool(SPAT)" window.
	  $start = TimerInit() ; start a timer to break us out or the loop
	  $elapsed = 0
	  $rv = 0
	  While (($rv = 0) And ($elapsed < $MS_5000))
		 WinActivate ("Open")
		 ControlClick("Open", "", "[CLASS:Button; TEXT:&Open; INSTANCE:1]")
		 $rv = WinWaitActive("Service Programming Archive Tool(SPAT)", "", $SEC_1)
		 $elapsed = TimerDiff($start)
	  WEnd

	  ; Wait until we return back to the "Service Programming Archive Tool(SPAT)" window.
	  $rv = WinWaitActive("Service Programming Archive Tool(SPAT)", "", $ERROR_TMO)
	  If $rv = 0 Then
		 FatalError("We did not return to the new archive utility window as expected. Exiting.")
	  EndIf
   Else   
	  FatalError("Open window did not open to allow us to select the template file. Exiting.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function pastes text into the Open window's filename box.
; It keeps trying until it reads back a match of what was sent.
; It will timeout in 5 seconds.
; $title - The title of the window.
; $id    - The control ID of the edit box.
; $text  - The text to paste into the box.
;//////////////////////////////////////////////////////////////////////////////
Func SetEditBoxNoVerify($title, $id, $text)
   Local $rv = ""
   Local $start
   Local $elapsed = 0

   Sleep($DEAD_TIME) ; wait a little for the window to settle

   $start = TimerInit() ; start a timer to break us out or the loop
   While (($text <> $rv) And ($elapsed < $MS_5000))
	  ControlSend($title, "", $id, "^{HOME}")
	  ControlSend($title, "", $id, "+{END}")
	  ControlSend($title, "", $id, $text)
	  $rv = ControlGetText ($title, "", $id)
	  $elapsed = TimerDiff($start)
   WEnd
   
   If $text <> $rv Then
	  FatalError("Can't seem to paste into the edit box. Exiting.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function pastes a binary filename into the edit box then reads it back
; to verify that the file was pasted correctly and that it exists.
; $title - The window title.
; $id    - The control ID for the file name box.
; $text  - The text to paste into the file name box.
;//////////////////////////////////////////////////////////////////////////////
Func SetEditBoxVerifyFile($title, $id, $text)
   Local $start
   Local $elapsed = 0
   Local $file
   Local $exists = 0

   Sleep($DEAD_TIME) ; wait a little for the window to settle
   
   $start = TimerInit() ; start a timer to break us out or the loop
   While (($exists = 0) And ($elapsed < $MS_5000))
	  ControlSend($title, "", $id, "^{HOME}")
	  ControlSend($title, "", $id, "+{END}")
	  ControlSend($title, "", $id, $text)
	  ControlSend($title, "", $id, "{DOWN}")
	  $file = ControlGetText ($title, "", $id)
	  $exists = FileExists($file)
	  $elapsed = TimerDiff($start)
   WEnd
   
   If $exists = 0 Then
	  FatalError("Part " & $text & " does not exist. Exiting.")
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; Selects a binary file.
; $y        - The y pixel value of the Select File button.
; $path     - The path
; $fileName - The binary filename.
; $row      - The row number in the config file.
;//////////////////////////////////////////////////////////////////////////////
Func SelectBinaryFile($y, $path, $name, $row)
   Local $rv = 0
   Local $start
   Local $elapsed = 0
   
   ; Verify that the Open window has displayed. If not retry for 5 seconds.
   $start = TimerInit() ; start a timer to break us out or the loop
   $elapsed = 0 
   While (($rv = 0) And ($elapsed < $MS_5000))
	  ControlClick("Build Calibration  Archive/Template File", "", $BUILD_CAL_ARCHIVE_ID, "", 1, $X_INIT, $y)
	  ControlClick("Build Calibration  Archive/Template File", "Select File", 1067)
	  $rv = WinWaitActive("Open", "", $SEC_1)
	  $elapsed = TimerDiff($start)
   WEnd
	  
   ; If the Open window did not display then exit.
   If $rv = 0 Then
	  FatalError("Open window did not open to allow us to select the binary file. Exiting.")
   Else
	  ; The Open window is active. Paste in the binary name
	  SetEditBoxVerifyFile("Open", "[CLASS:Edit; INSTANCE:1]", ($path & $name))
	  ;SetEditBox("Open", "[CLASS:Edit; INSTANCE:1]", ($path & $name), $DOWN)
	  
	  ; Click the "Open" button to open the binary file and verify the we returned to the "Build Calibration" window.
	  $start = TimerInit() ; start a timer to break us out or the loop
	  $elapsed = 0
	  $rv = 0
	  While (($rv = 0) And ($elapsed < $MS_5000))
		 WinActivate ("Open")
		 ControlClick("Open", "", "[CLASS:Button; TEXT:&Open; INSTANCE:1]")
		 $rv = WinWaitActive("Build Calibration  Archive/Template File", "", $SEC_1)
		 $elapsed = TimerDiff($start)
	  WEnd
   EndIf
EndFunc

;//////////////////////////////////////////////////////////////////////////////
; This function checks if the SPAT window is active.
;//////////////////////////////////////////////////////////////////////////////
Func CheckForSpatWindowActive()
   ; Wait 1 second for the SPAT window to become active
   Local $rv = WinWaitActive("Service Programming Archive Tool(SPAT)", "", $SEC_1)
   
   ; Did the timeout occur?
   If $rv = 0 Then
      ; Spat window is not active. It could be because a status message box is open.
	  FatalError("Closes any open message boxes. Then close this one. Exiting.")
   EndIf
EndFunc

#cs
Change Log
V12.04.10 - Changed the main error timeout from 15 seconds to 30. 
            Fixed the Auto-Correct message. It was being logged all the time.
V12.04.04 - Changed the main error timeout from 5 seconds to 15 because my display fails sometimes 
            but recovers though it takes more than 5 seconds.
            Ignoring the last column of the header row if it ends with a comma.
            Strip off any double qoutes in the archive name in ArchiveHandler().
			If there are issues open the log in Notepad, else don't open the log.
V12.03.11 - Added validation of the archive file name read from the config.
            Must be 32 chars or less. First column can't be blank. Cols 2 & 3 can be blank.
			Remove prepend of Rev for column 3. Must add Rev or R now.
V2.0_120302 - Added an "Overwrite Existing Archives" checkbox and functionality.
              Added Help.
			  Added functionality to save the GUI settings.
V2.0_120301 - Fixed the problem of the "Archive File Was Created Successfully" message not always closing.
V2.0_120229 - Added functionality to skip any row with a blank part entry.
V2.0_120228 - Added functionality to check if archive already exists and allow the user to skip it.
#ce