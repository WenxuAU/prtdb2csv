; This program runs ROAMES program to convert the prtdb files into csv files.

; ----------------------------initialize the program--------------------------------------------------
;-------------input foders and files to be specified through the inputs.ini file--------------------
;1. prtdb_dir: the directory to prtdb files to be converted
;2. fout: the output folder where the converted csv files from prtdb files will be stored.
;3. fconfig: the configuration files for csv outputs, depending on the vehicles used.
;---------------------------------------------------------------------------------------------------
#include <File.au3>
#include <GuiListView.au3>

; Read the INI section labelled 'General'. This will return a 2 dimensional array.
$sFolder = @WorkingDir ; inpired by https://www.autoitscript.com/forum/topic/12042-how-to-know-the-current-directory/
$sFilePath = $sFolder&"\inputs.ini"
Local $sVehicles,$aArray
$sNames = IniReadSectionNames ($sFilePath)

;~ If Not  @error Then
;~ 	ConsoleWrite($sNames[0]&@LF)
;~ 	For $iI = 1 to $sNames[0] 
;~ 		ConsoleWrite($sNames[$iI]&@LF)
;~ 	Next
;~ EndIf

; read the exe path of ROAMES
$aArray=IniReadSection($sFilePath, $sNames[1])
If Not @error Then	
	$sfROAMES=$aArray[1][1]
EndIf
ConsoleWrite("$ROMES_dir= "&$sfROAMES&@LF) ; test

For $iI = 2 to $sNames[0] ; loop over all vehicles
	$aArray = IniReadSection($sFilePath, $sNames[$iI])
	; Check if an error occurred.
	If Not @error Then
		; read the file paths and folders	
		$prtdb_dir = $aArray[1][1]
		$fconfig = $aArray[2][1]
		$fout =  $aArray[3][1]
		ConsoleWrite("$prtdb_dir= "&$prtdb_dir&@LF&"$fconfig="&$fconfig&@LF&"$fout="&$fout&@LF)
		prtdb2csv($sfROAMES, $prtdb_dir,$fconfig,$fout,$sNames[$iI])
	EndIf	
Next

;;; Directly specify the file  and folder paths
;~ ;$prtdb_dir = "D:\oTCI\01_PROD_IN_USE\02_QC_Manned_Vehicles\Files\prtdbs\MMY038" 
;~ ;$fconfig = "Q:\_Charts\_System_Reporter_Reports\QC_Reports\MMY038 IRRISYS.rep"
;~ ;$fout = "D:\oTCI\01_PROD_IN_USE\02_QC_Manned_Vehicles\Files\Input\Raw_CSVs\MMY038" ; output folder


Func prtdb2csv(ByRef $sfROAMES, ByRef $prtdb_dir,ByRef  $fconfig,ByRef $fout, ByRef $vehicle)
	
	Local $iI, $fcount
	;;;; ================use the following to test wether the item exits in the TRzShellTree================
	;$fout = "Desktop\This PC\DATA (D:)"
	;;;;;=============================end of the test==================================
	; select the prtdb files in the folder for conversion into csv files.
	$aFileList = _FileListToArray($prtdb_dir,"*.prtdb",1,True )
	; Display the results returned by _FileListToArray.
	;_ArrayDisplay($aFileList, "$aFileList") ; this will prompt a window pop up, commented here
	If  @error==0 Then
		if $aFileList[0] Then
			$fcount = $aFileList[0]  ; count the number of prtdbf files in the selected folder
			ConsoleWrite("The number of the prtdb files in the folder is "&$fcount&"."&@LF)
			Dim $aSplit[$fcount]
			$aSplit = _ArrayToString($aFileList, @LF)
			;ConsoleWrite("The number of the prtdb files in the folder is "&$aSplit&@LF)
			
			; ----------------------------------------Run ROMAMES-------------------------------------------------------------------
			Run($sfROAMES) ; run ROAMES
			
			WinWaitActive("[CLASS:TReportMainForm]")

			;----------------------------------- Load all of the prtdb files in the specified folder path------------------------------
			WinMenuSelectItem("[CLASS:TReportMainForm]","","&File","&Open Files...") ; configuration file
			ControlSetText ( "[CLASS:TFileListForm]", "", 1001, $prtdb_dir)
			GUICtrlSetState(1001,$GUI_FOCUS)
			Local $iEnd = StringLen(GUICtrlRead(1001))
			_GUICtrlEdit_SetSel(1001, $iEnd, $iEnd)
			_GUICtrlEdit_Scroll(1001, $SB_SCROLLCARET)
			ControlClick( "[CLASS:TFileListForm]", "", 1001, button = "left" , clicks = 2)
			send(SPACE)
		
			; get the handle to the listview the active window
			$ListView = ControlGetHandle("[CLASS:TFileListForm]","", "[CLASS:TListView; INSTANCE: 1]")

			; select all the prtdb files 
			_GUICtrlListView_GetItemCount($ListView)
			ConsoleWrite("The number of the items in the listview is "&_GUICtrlListView_GetItemCount ($ListView)&"."&@LF)

			; loop over all of the prtdb files to select all of them
			For $iI =0 To $fcount
				_GUICtrlListView_SetItemSelected( $ListView, $iI, True)
			Next 
;~ 			; click the open button to load the prtdb files into ROAMES software
;~ 			ControlClick("[CLASS:TFileListForm]","","[CLASS:TButton; INSTANCE:2]")

;~ 			;----------------load the configuration file--------------------------------------------------------------------------------------
;~ 			; the pop up is a UIHWND class, to select items from this type of window is not possible. In addition the top window is 
;~ 			; #32770 type,  not an Explore|Cabinet type, therefore the post here which is based on the shell window Explorer and 
;~ 			; Cabinet: https://www.autoitscript.com/forum/topic/155542-how-can-i-select-an-item-from-directuihwnd2-type-window/
;~ 			;----------------------------------------------------------------------------------------------------------------------------------
;~ 			; 1. remove all configuration files already loaded in the memory if there is any
;~ 			$RemoveAllBtn = ControlGetHandle("[CLASS:TReportMainForm]","","[CLASS:TButton; INSTANCE: 1]")
;~ 			ControlClick("[CLASS:TReportMainForm]","",$RemoveAllBtn)

;~ 			;2. load the new configuration file according to the vehicle type used
;~ 			; get the handle of the Load Report button
;~ 			$LoadReportBtn = ControlGetHandle("[CLASS:TReportMainForm]","","[CLASS:TButton; INSTANCE: 4]")
;~ 			ControlClick("[CLASS:TReportMainForm]","",$LoadReportBtn)
;~ 			; select the configuration file to load in the open file dialog
;~ 			WinWaitActive("Load Report File")
;~ 			$hWin=WinGetHandle("Load Report File")
;~ 			; if IsHWnd($hWin) then ConsoleWrite("Load Report File is an object."&@LF); test
;~ 			; $UIhwnd = ControlGetHandle($hWin,'',"[CLASS:DirectUIHWND; INSTANCE:2]")

;~ 			ControlSetText($hWin,"","[CLASS:Edit; INSTANCE:1]","")              ; clear the file path first.
;~ 			ControlSetText($hWin,"","[CLASS:Edit; INSTANCE:1]",$fconfig)      ; set the file path.
;~ 			;Send($fconfig) ;this can also work but differently
;~ 			ControlClick($hWin,"","[CLASS:Button;INSTANCE:1]")

;~ 			; ----------------------set up the output folder---------------------------------------------------------------------------------------

;~ 			$arrayofsubstring =  StringSplit($fout,"\") ; split the string into different subsections (each folder)
;~ 			For $iI = 1 to $arrayofsubstring[0] ; print the splitted string elements
;~ 			 	;UBound($arrayofsubstring) ; another method to get the array size
;~ 			 	ConsoleWrite($arrayofsubstring[$iI]&"."&@LF)
;~ 			Next 
;~ 			WinActivate("[CLASS:TReportMainForm]")		
;~ 			WinWaitActive("[CLASS:TReportMainForm]")
;~ 			WinMenuSelectItem("[CLASS:TReportMainForm]","","&Configure","&Text Output Folder...") 
;~ 			WinWaitActive("Text Report Output Folder")
;~ 			$hwnd_report = WinGetHandle("Text Report Output Folder")
;~ 			$Tree = ControlGetHandle($hwnd_report,"","[CLASS:TRzShellTree; INSTANCE:1]")
;~ 			;;;;ControlGetHandle($hwnd_report,"",$Tree)
;~ 			;;ConsoleWrite("The TRzShellTree is "&$Tree&@LF);; for test only
;~ 			; inspired by this post: https://www.autoitscript.com/forum/topic/151338-click-an-item-in-treeview-or-systreeview32/
;~ 			; tried  _GUICtrlTreeView xxx but I cannot get them to work but ControlTreeView can. The key is to specify an $Item which 
;~ 			; includes all of the parent folders up to the one of interest
;~ 			;;;ControlGetHandle($hwnd_report,"",$Tree)
;~ 			$Item = $arrayofsubstring[1]		
;~ 			$sExist = ControlTreeView($hwnd_report,"",$Tree,"Exists",$Item)
;~ 			ControlTreeView($hwnd_report,"",$Tree,"Check",$Item)
;~ 			ControlTreeView($hwnd_report,"",$Tree,"Select",$Item)
;~ 			ControlTreeView($hwnd_report,"",$Tree,"Expand",$Item)
;~ 			Sleep(5)
;~ 			;ConsoleWrite($Item&" " &$sExist&@LF)
;~ 			For $iI = 2 to $arrayofsubstring[0]
;~ 				$Item = $Item & "|" & $arrayofsubstring[$iI]				
;~ 				$sExistItem = ControlTreeView($hwnd_report,"",$Tree,"Exists",$Item)
;~ 				ConsoleWrite($Item&" Exists: "&$sExistItem&@LF)
;~ 				ControlTreeView($hwnd_report,"",$Tree,"Check",$Item)
;~ 				Sleep(5)
;~ 				ControlTreeView($hwnd_report,"",$Tree,"Select",$Item)
;~ 				Sleep(5)
;~ 				ControlTreeView($hwnd_report,"",$Tree,"Expand",$Item)
;~ 				Sleep(5)
;~ 			Next
;~ 			ControlClick("Text Report Output Folder","","[CLASS:TRzBitBtn; INSTANCE:2]")

;~ 			; ----------------------start the conversion-------------------------------------------------------------------------------------------
;~ 			;;;WinWaitActive("ROAMES System Reporter")
;~ 			$hWnd_Repoter = WinGetHandle("[CLASS:TReportMainForm]")
;~ 			Sleep(5)
;~ 			ControlClick($hWnd_Repoter,"","[CLASS:TButton; INSTANCE:6]")

;~ 			;---------------------wait until the conversion is done-------------------------------------------------------------------------------
;~ 			;While Not ControlCommand($hWnd_Repoter,"","[CLASS:TButton; INSTANCE:6]","IsEnabled")
;~                 ;ConsoleWrite("Conversion is still ongoing...")
;~ 				;Sleep(2)
;~ 			;WEnd
;~ 			Sleep(5)
;~      		WinActivate($hWnd_Repoter)
;~ 			Local $Status = ControlGetText($hWnd_Repoter,"","[CLASS:TPanel; INSTANCE:1]")
;~ 			While Not ($Status=="Done")
;~ 				$Status = ControlGetText($hWnd_Repoter,"","[CLASS:TPanel; INSTANCE:1]")			
;~ 				ConsoleWrite("The current conversion status is "&$Status&@LF)
;~ 				Sleep(10)
;~ 			WEnd
;~ 			WinMenuSelectItem("[CLASS:TReportMainForm]","","&File","E&xit") ;Exit the program
;~ 			Sleep(10)
;~ 			ConsoleWrite("Conversion is "&$Status&"."&@LF)
		EndIf
	Else
		;create an error Log
		ConsoleWrite("There is an error with a code " &@error &" in prtdb to csv conversion " &" for "&$vehicle&"." &@LF)		
		
	EndIf
EndFunc


