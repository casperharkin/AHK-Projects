#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#Include C:\Users\babb\Documents\Autohotkey\Lib\LV_InCellEdit.ahk
#SingleInstance force
#ClipboardTimeout 2000
OnExit("Exit",1)

Global Thread_Kill_Token := False, __OBJECT := {}, __Handles := {}, SuperDetails := {}, Client_Info
New Gui_Class()
SessionChange()
Exit ; End of Auto Exec Section




;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~ HOTKEYS ~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#q::Thread_Kill_Token := !Thread_Kill_Token 

^#P::Workbook_Class.Excel_Sheet_to_PDF()
#F12::run, E:\PortableApps\AHK Studio\AHK-Studio.ahk %A_ScriptName%,,,OutputVarPID
#w::WMPLayer.Toggle()
^#f::Workbook_Class.FinalEntsHotKey()



#x::
FormatTime, time, A_now, dd/MM/yyyy
send % time A_TAB "BABB"
return


#v::
FormatTime, Time,, dd/MM/yyyy
Send %Time%
Return

^RButton:: ; Set Control + Right Click as the hotkey to open the Cotnext Menu
Menu, ContextMenu, Show ; Show Context Menu
return

Exit() {
	Gui_Class.QMaster("Shutdown")
	Gui_Class.Workarea(A_ScreenWidth, A_ScreenHeight-27)
}

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~ MENU HANDLER ~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

MenuHandler() 
{
	Global
	Gui, submit, NoHide
	
	if (A_GuiControl = "Load Client Info") 	{	
		__OBJECT := New Workbook_Class(ServiceNo := StrReplace(StrReplace(ServiceNo, "-", A_Space), " ")) ; Fetching Client Info
		RefreshListView() 		
	}
	
	if (A_GuiControl = "Prepare Paid Due Diff") 
		Workbook_Class.PreparePaidDueDiff(__OBJECT)
	
	if (A_GuiControl = "Process Workbook") 
		Workbook_Class.Process_Workbook(__OBJECT)
	
	if (A_GuiControl = "Create ePOD Record") 
		Workbook_Class.CreateePODRecord(__OBJECT)
	
	if (A_GuiControl = "Adjust Gross Totals") 
		Workbook_Class.AdjustCumlativeTotals(__OBJECT)
	
	if (A_GuiControl = "Enter Recovery Action") 
		Workbook_Class.EnterRecoveryAction(__OBJECT)
	
	if (A_GuiControl = "Generate Subject Line") 
		Clipboard := "To Be Checked | BABB | Pay " . __OBJECT.Pay_Cycle . " | " . __OBJECT.ServiceNo . " | " . __OBJECT.Person_Last_Name . ", " . __OBJECT.Person_Given_Name . " (" . __OBJECT.Person_Pay_Centre . ") | " . __OBJECT.Date_Commenced . "-" . __OBJECT.Date_Ceased . " | $" . __OBJECT.TotalAmount
	
	if (A_GuiControl = "Clear") {
		__OBJECT :=  
		RefreshListView() 
	}
	
	if (A_GuiControl = "SG Recovery Action") 
		Workbook_Class.SGR(__OBJECT)

	if (A_GuiControl = "Refresh from Workbook") {
		Workbook_Class.RefreshFromWorkbook
		RefreshListView() 
		msgbox done



	}
}

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
;~~~~~~~~~ QUICK ACCESS MENU HANDLER ~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

QuickAccessHandler() 
{
	Global
	if (A_GuiControl = "Recent") {
		Gui_Class.OpenRecent()
		return	
	}
	
	if (A_GuiControl = "ePOD") {
		;MsgBox Place Holder
		return	
	}
	
	if (A_GuiControl = "PIPS") {		
		Gui_class.PIPS_Toolbar()
	}
	
	if (A_GuiControl = "F:\") {
		Gui_Class.Application_Control_Workaround("F:\Finance\Payroll Debt Recovery")
		return	
	}
	
	if (A_GuiControl = "Chrome") {
		Gui_Class.Application_Control_Workaround("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
		return	
	}
	
	if (A_GuiControl = "Phones") {
		run microsoft-edge:http://ntgcentral.nt.gov.au/phones?search=
		return	
	}
	
	if (A_GuiControl = "myHR") {
		run https://myhr.nt.gov.au/
		return	
	}
	
	if (A_GuiControl = "GovAcc") {
		Gui_Class.Application_Control_Workaround("iexplore.exe https://gas-web.nt.gov.au/")
		return	
	}
	
	if (A_GuiControl = "VarDB") {
		Gui_Class.Application_Control_Workaround("C:\Users\babb\Documents\Autohotkey\Lib\DebugVars.ahk")
		
		return	
	}
	
	if (A_GuiControl = "Q-Mast") {
		Gui_Class.QMaster()
		return	
	}
	
	if (A_GuiControl = "Calc") {
		Gui_Class.Application_Control_Workaround("C:\_Umbra Sector\_Projects\New Calc\MainCalc.ahk")
		return	
	}
	
	if (A_GuiControl = "FileNote") {
		Gui_Class.Application_Control_Workaround("C:\_Umbra Sector\_Tools\File Note\Create File Note.ahk")
		return	
	}
}

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~ PIPS MENU HANDLER ~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PIPS_Menu(){
	WinActivate Mochasoft - mainframe.nt.gov.au
	ControlGetText, OutputVar, Edit1,  % "ahk_id " __Handles.hComboBox
	
	if FindText(1538-150000, 154-150000, 1538+150000, 154+150000, 0, 0, Check_For_Time_Out := "|<>*48$91.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzsDzsDw7kDUDz00TllztnznzTzLzjjjnyztwzszjz9zrrrzzDxzTwDrzizvvvzzbwzbybvzrTxxxzznyTnzNxznbyyyzzvzDxzgyzvvzzTzztzbyzrDTxwzzjzz1znzTvrjwyTzrzzwTtzjxtryzjzvzzzbwzryyPz03zxzzznyTvzTBzDxzyzzzxzDtzjmzjyzzTzzyzbwzrwTbzDzjzzyTvyTvyDnzrzrzvzDwyTxzbvzvzvzwSDzCTs7nkDUT07zUTzkTzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk") {
		Send {Enter 3}
		Sleep 2000
	}
	
	if (A_GuiControl = "Load")
		SendInput, {F4}1.2.1{Enter}%OutputVar%{Enter}
	else
		SendInput, {F4}{Enter}{Enter}{Enter}%A_GuiControl%{Enter}{F8}	
}
	
	
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~ CONTEXT MENU HANDLERS ~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

ContextMenuHandler() 
{
	Global
	;msgbox % A_ThisMenuItem
	
	if (A_ThisMenuItem = "Chrome") {
		Gui_Class.Application_Control_Workaround("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
		return	
	}
	
	if (A_ThisMenuItem = "eReader") {
		Gui_Class.Application_Control_Workaround("C:\_Umbra Sector\_Tools\eReader Class\eReader Class.ahk")
	}
	if (A_ThisMenuItem = "Quick Code Tester") or (A_ThisMenuItem = "Quick Code") {
		Gui_Class.Application_Control_Workaround("C:\_Umbra Sector\_Tools\QuickTester\QuickTester_v2.7.ahk")
	}
	if (A_ThisMenuItem = "Create Manual Payee Sheet") {
		Workbook_Class.Create_Manual_Payee_Sheet()
	}
	if (A_ThisMenuItem = "Reset Work Area") or (A_GuiControl = "Edit This Script"){
		Gui_Class.Workarea(A_ScreenWidth-215, A_ScreenHeight-32)
	}
	if (A_ThisMenuItem = "AHK Studio") {
		run, E:\PortableApps\AHK Studio\AHK-Studio.ahk %A_ScriptName%,
	}
	if (A_ThisMenuItem = "Refresh from Workbook") {
		Workbook_Class.Refresh_from_Workbook()
		
	}
	if (A_ThisMenuItem = "Remove Symbols") 
	{
		ClipBoard := Workbook_Class.Change_String(ClipBoard, "Strip")
	}
	
	if (A_ThisMenuItem = "Remove Formating") 
	{
		ClipBoard := Workbook_Class.Change_String(ClipBoard, "Plain")
	}
	
	if (A_ThisMenuItem = "Uppercase Clipboard") 
	{
		ClipBoard := Workbook_Class.Change_String(ClipBoard, "Upper")
	}
	
	if (A_ThisMenuItem = "TitleCase Clipboard") 
	{
		ClipBoard := Workbook_Class.Change_String(ClipBoard, "Title")
	}
	
	if (A_ThisMenuItem = "Lowercase Clipboard") 
	{
		ClipBoard := Workbook_Class.Change_String(ClipBoard, "Lower")
	}
	
	if (A_ThisMenuItem = "StrReplace") 
	{
		Inputbox, ReplaceText, What to Replace:, What Char did you want to replace?,,240,130
		Clipboard := StrReplace(Clipboard, ReplaceText)
	}
	
	if (A_ThisMenuItem = "To Be Checked") 
	{
		
		ClipBoard = To Be Checked | Bryn | %TimeString% | 
	}
	
	if (A_ThisMenuItem = "Checked") 
	{
		ClipBoard = Checked | Bryn | %TimeString% | 
	}
	
	if (A_ThisMenuItem = "Invoice Raised") 
	{
		ClipBoard = Invoice Raised | Bryn | %TimeString% |{space}
	}
	
	if (A_ThisMenuItem = "Trimmed") 
	{
		ClipBoard = Trimmed | Bryn | %TimeString% | 
	}
	
	if (A_ThisMenuItem = "Sent to Client") 
	{
		ClipBoard = Sent to Client | Bryn | %TimeString% | 
	}
	
	if (A_ThisMenuItem = "Make Selection Negative Number") 
	{
		Positive := Clipboard
		Doubble := Positive * 2
		Negative := Positive - Doubble
		Clipboard := Negative
	}
	if (A_ThisMenuItem = "Mini Player Mode") {
		MenuItem := A_ThisMenuItem
		MenuName := A_ThisMenu
		Flag := !Flag ; Toggles the variable every time the function is called
		If (Flag) {
			Menu, %MenuName%, Check, %MenuItem%
			WMPLayer.Mode("mini") 
		}
		else {
			Menu, %MenuName%, UnCheck, %MenuItem%
			WMPLayer.Mode("none") 
		}
	}
	if (A_ThisMenuItem = "Select File") {
		WMPLayer.Select() 
		
	}
	
	if (A_ThisMenuItem = "GAS Notes")
	{
		Gui_Class.Application_Control_Workaround("C:\_Umbra Sector\_Projects\GAS Notes Wizard\GAS Notes Wizard.ahk")
	}
	
	if (A_ThisMenuItem = "Close WMP Instance")
	{
		WMPLayer.Close()
	}
	
	if (A_ThisMenuItem = "Build OVP Email") {
		
		MsgBox 0x24, Generate Super Email?, Do you wish to Generate a Super Email?
		
		IfMsgBox Yes	
			GenerateaSuperEmail := True
		Else IfMsgBox No
			GenerateaSuperEmail := False
		
		Gui_Class.emailbuild(GenerateaSuperEmail)
	}

	
	if (A_GuiControl = "Refresh from Workbook") {
		Workbook_Class.RefreshFromWorkbook
		RefreshListView() 
		msgbox done



	}
}


;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~ LIST VIEW HANDLER ~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~ AND REFRESH ~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;

SubLV1() 
{  
	If (A_GuiEvent == "F") {
		If (ICELV1["Changed"]) {
			Msg := ""
			For I, O In ICELV1.Changed
				Msg .= "Row " . O.Row . " - Column " . O.Col . " : " . O.Txt
			ICELV1.Remove("Changed")
		}
		Loop % LV_GetCount() {
			LV_GetText(VarName, A_Index)
			LV_GetText(VarData, A_Index, 2)
			__OBJECT[VarName] := TRIM(VarData)
		}
		RefreshListView()
	}
}

RefreshListView() 
{
	for key, value in __OBJECT {
		__OBJECT[key] := TRIM(value)
		If (__OBJECT[key] = "")
			__OBJECT.Delete(Key)
	}
	
	LV_Delete()
	GuiControl, -Redraw, MyListView 
	For Key, Value in __OBJECT
		LV_Add("", Key , Value )
	LV_ModifyCol(1,"")   
	LV_ModifyCol(2,"")
	GuiControl, +Redraw, MyListView
}



SessionChange(notify := true) {
    static WTS_CURRENT_SERVER := 0, NOTIFY_FOR_ALL_SESSIONS := 1

    if (notify)  ; http://msdn.com/library/bb530723(vs.85,en-us)
    {
        if !(DllCall("wtsapi32.dll\WTSRegisterSessionNotificationEx", "Ptr", WTS_CURRENT_SERVER, "Ptr", A_ScriptHwnd, "UInt", NOTIFY_FOR_ALL_SESSIONS))
            return false
        OnMessage(0x02B1, "WM_WTSSESSION_CHANGE")
    }
    else         ; http://msdn.com/library/bb530724(vs.85,en-us)
    {
        OnMessage(0x02B1, "")
        if !(DllCall("wtsapi32.dll\WTSUnRegisterSessionNotificationEx", "Ptr", WTS_CURRENT_SERVER, "Ptr", A_ScriptHwnd))
            return false
    }
    return true
}

WM_WTSSESSION_CHANGE(wParam, lParam) { ; http://msdn.com/library/aa383828(vs.85,en-us)
	static WTS_SESSION_UNLOCK := 0x8
	
	if (wParam = WTS_SESSION_UNLOCK) {
		Sleep 2000
		Gui_Class.Workarea(A_ScreenWidth-215, A_ScreenHeight-32)
	}
}


;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~ WORKBOOK CLASS ~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Class Workbook_Class
{
	__New(ServiceNo)
	{
		
		Local 
		
		Client_Info := {}
		Client_Info.ServiceNo := ServiceNo
		This.ServiceNo := Client_Info.ServiceNo 
		Client_Info.FBT_Date := This.FBT_Date()
		
		This.CheckListArray :=   {"TIMEOUTCHECK":"|<>*48$91.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzsDzsDw7kDUDz00TllztnznzTzLzjjjnyztwzszjz9zrrrzzDxzTwDrzizvvvzzbwzbybvzrTxxxzznyTnzNxznbyyyzzvzDxzgyzvvzzTzztzbyzrDTxwzzjzz1znzTvrjwyTzrzzwTtzjxtryzjzvzzzbwzryyPz03zxzzznyTvzTBzDxzyzzzxzDtzjmzjyzzTzzyzbwzrwTbzDzjzzyTvyTvyDnzrzrzvzDwyTxzbvzvzvzwSDzCTs7nkDUT07zUTzkTzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk"
							,"MAINMENUCHECK":"|<>*45$41.zzzzzzzzzzzzzztzw7znznznbzbzDzDbzbyTyzjzDsztzDzDnznyTyTbzbyzwSDzDxzwwzyTvzttzwzrznnztzjzbbznzTzDDzbwzySTzDtzwwTzTnztwzyTDzXtzyQzzDlzy3zyTnzzzztzbzzzznzbzzzzDzjzzzyTzzzzzzz"							
						     ,"SEPARATEDCHECK":"|<>*106$78.zzzzzzzbzzzzzzzzzzzzbzzzzzzDzzzzzDzzzyTzDkD1zzDUy1yTyDwzrzyDwTnzDyTyzbzyTwTnzDwTyTDzyTwDnzbwzzDDzwzxDnzbwzziTzwzxbnzXszzYzztzxbnznszzkzztzxnnznszztzznzxtnzntzztzznzxtnzntzztzzbzxwnznszztzzbzxwnznszztzzDzxyHznszztzzDzxz3znwzztzzDzxz3zXwzztzyTzxzXzbwTy0DyTzkDXzbyTzzzwzzzzzzDyDzzzwzzzzzzDzDzzzzzzzzzyTzjzzzzzzzzzyTzzzzzzzzzzzzzU"
					          ,"CASUALCHECK":"|<>*106$101.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzs7w2TzzzzzzzzzzzzzDXkzzzzzzzzzzzzzySTlzzzzzzzzzzzzzwxznzzzzzzzzzzzzztnzby1zy0DVsDy1zznbzzlszswTnyTlszzbTzzztzbwzbwzztzzCzzzzvzDtzDtzzvzyRzzzzrzDzyTnzzrzwvzzzzjz0Twzbzzjztrzzy0TzyTtzDy0TznbzzlwzzyTnyTlwzzbDzzDxzzyTbwzDxzzDDySTvzTwzDtyTvzyTTtwzbyTnzTXwzbzwzD7swDwTDyS7swDztz0zw33v0zy13w33k0Tzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
						  	,"TEMPCHECK":"|<>*47$101.zzzzzzwzzzzzzzzzzzzzzzztzzzzzzzzzzbzzzzzbzzzzDzzzzzDkT1zzDUy1yTzzzzwztzjzwztzjyTzzzztzvyTztzlzTwzzzzzXzntzzrzVyzwzzzzzDznrzzDzHxztzzzzyTzrDzyTynvzlzzzzszzYzztzxbrznzzzznzzXzznzvbjzbzzzzbzzbzzDzrjTzDzzzzDzzTzyTzjCzyTzzzyTzyzztzzTBzwzzzzwzzxzznzyyPztzzzztzzvzzDzxyLznzzzzlzzrzyTzvyDzbzzXznzzjzxzzrwTyDzy3zbzzTznzzjwzwzzw7z7zU3zjzw3tztzzwTzDzzzyTzzzzzbzzzzyTzzzwzzzzzzDzzzzyTzzzzzzzzzwzzzzzyzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzz"
						    	,"COFCHECK":"|<>*107$101.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzUHzUTs01zzzzzzzzwS7wSTwznzzzzzzzznyDnyTtzbzzzzzzzzjyT7yTnzDzzzzzzzyTwyTwzbyTzzzzzzzwzzwzwzDzzzzzzzzzvzznztyQzzzzzzzzzbzzbznwtzzzU0DzzzDzzDzbs3zzz00TzzyTzyTzDnbzzzzzzzzyzzwzyTbDzzzzzzzzwzzxzwzDzzzzzzzzztzztznyTzzzzzzzzztzntzbwzzzzzzzzzzvzDnyTtzzzzzzzzzztszltznzzzzzzzzzzs7zs7y0Dzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
							,"NTGPASSCHECK":"|<>*107$101.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzUy1s03zUHzzzzzzzzszbnnrwy7zzzzzzzzlzDbbjnzDzzzzzzzzVyTDDTjyTzzzzzzzzHwySSyTzzzzzzzzzyntwwxwzzzzzzzzzzxbnztzvzzzzzzzzzzvbbznzbzzzzU0DzzzrbDzbzDzzzz00TzzzjCTzDyTUDzzzzzzzzTAzyTyzwzzzzzzzzyyNzwzwztzzzzzzzzxyHztztznzzzzzzzzvy7znznzbzzzzzzzzrwDzbznzDzzzzzzzzjwTzDzlyTzzzzzzzw3szU3zs3zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
							,"SPR":"|<>*108$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzkC1w1zk0TUy1bzrzdztyTszbjzDyHznyTlzDTyTwrzbyzVySyQzvbzDwzHwwstzbDyTvynttlnzDTwzbxbnndbyyTtwTvbbaHDtwzk3zrbDAqznxzbbzjCTPhzU1zDbzTAyr/yTvyTbyyNxC7xznwzbxyHsQDnzbtzjvy7lwTbzjnzDrwDXszTzDbzTjwT7ls7kA1yA3szzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"}
		
		SuperDetails := {}
		SuperDetails["NEPAXG"] := {Name:"AXA Generations Personal Super- Carol Duke",Street:"750 Collins Street", State:"West Victoria  VIC  8007"}
		SuperDetails["NEPAXG"] := {Name:"SPNT PTY LTD",Street:"PO Box 1827", State:"New Farm  QLD  4005"}
		SuperDetails["NEPNV1"] := {Name:"SPNT NL - FM (SALARY PACKAGING AND NEW TECHNOLOGY PTY LTD)",Street:"PO BOX 1827", State:"NEW FARM QLD 4005"}
		SuperDetails["NEPNV2"] := {Name:"SPNT NL - ECM (SALARY PACKAGING AND NEW TECHNOLOGY PTY LTD)",Street:"PO BOX 1827", State:"NEW FARM QLD 4005"}
		SuperDetails["NEPNV3"] := {Name:"EPAC NL - FM (EPAC SALARY SOLUTIONS PTY LTD)",Street:"PO BOX 373", State:"RUNDLE MALL SA 5000"}
		SuperDetails["NEPNV4"] := {Name:"EPAC NL - ECM (EPAC SALARY SOLUTIONS PTY LTD)",Street:"PO BOX 373", State:"RUNDLE MALL SA 5000"}
		SuperDetails["NEPNV5"] := {Name:"FLEETNETWORK NL - FM (FLEETNETWORK)",Street:"PO Box 2461", State:"MALAGA WA 6090"}
		SuperDetails["NEPNV6"] := {Name:"FLEETNETWORK NL - ECM (FLEETNETWORK)",Street:"PO Box 2461", State:"MALAGA WA 6090"}
		SuperDetails["NEPNV7"] := {Name:"Vehicle Solutions Australia (Novated Lease Company)",Street:"PO Box 1625", State:"PALMERSTON NT 0831"}
		SuperDetails["NEPNV8"] := {Name:"Fleet Choice NT (Novated Lease Company)",Street:"GPO Box 3961", State:"DARWIN NT 0801"}
		SuperDetails["NEPPOW"] := {Name:"PowerWater In-House Benefits",Street:"Salary Packaging Unit & Executive Payroll Unit", State:""}
		SuperDetails["NEPX09"] := {Name:"The Portfolio Service (Executive)",Street:"Locked Bag 800", State:"Milsons Point NSW 1565"}
		SuperDetails["NEPX13"] := {Name:"Smartsave Members Choice - J Harrison",Street:"PO Box 20314", State:"World Square"}
		SuperDetails["NEPX14"] := {Name:"QSuper",Street:"GPO Box 200", State:"Brinband QLD 4002"}
		SuperDetails["NEPX24"] := {Name:"MLC WRAP SUPER",Street:"GPO BOX 2567", State:"MELBOURNE  VIC  3001"}
		SuperDetails["OTENJTK"] := {Name:"Central Australia Hospital Network CRA Fund",Street:"GPO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTENTC"] := {Name:"DCIS CRA Fund",Street:"GPO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTENTE"] := {Name:"NT Land Corp CRA Fund",Street:"PO Q106, QVB Post Shop", State:"Sydney NSW 1229"}
		SuperDetails["OTENTG"] := {Name:"Government Printing Office CRA Fund",Street:"GPO Box 1447", State:"Darwin NT 0801"}
		SuperDetails["OTENTJ"] := {Name:"Top End Hospital Network CRA Fund",Street:"GPO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTENTL"] := {Name:"Legal Aid Commission CRA Fund",Street:"GPO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTENTP"] := {Name:"PAWA CRA FUND",Street:"PO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTENTS"] := {Name:"Batchelor CRA fund",Street:"RTM Palm Court", State:"PO Box 2391"}
		SuperDetails["OTENTT"] := {Name:"Land Development Corporation CRA Fund",Street:"PO Box 2391", State:"Darwin  NT  0801"}
		SuperDetails["OTN000"] := {Name:"Commonwealth Life Personal Super",Street:"PO Box 3306", State:"Sydney  NSW  1030"}
		SuperDetails["OTN001"] := {Name:"Commonwealth Bank Superannuation Savings Account",Street:"GPO Box 3306", State:"Sydney NSW 2001"}
		SuperDetails["OTN003"] := {Name:"Health Super Fund",Street:"Locked Bag 2900 Collins Street", State:"West Melbourne  VIC  8007"}
		SuperDetails["OTN004"] := {Name:"AXA Australia",Street:"Super Department", State:"PO Box 14669"}
		SuperDetails["OTN005"] := {Name:"MLC Empl Retirement Plan",Street:"PO Box 200", State:"North Sydney  NSW 2059"}
		SuperDetails["OTN006"] := {Name:"MLC Universal Super",Street:"MLC Nominees Pty Ltd", State:"105-153 Miller Street"}
		SuperDetails["OTN007"] := {Name:"REST",Street:"PO Box 350", State:"Parramatta  NSW 2124"}
		SuperDetails["OTN008"] := {Name:"NAB All in One Super",Street:"PO Box 4341", State:"Melbourne  Vic  3001"}
		SuperDetails["OTN009"] := {Name:"LifeTrack Superannuation Fund",Street:"GPO Box 264C", State:"Melbourne VIC 3001"}
		SuperDetails["OTN012"] := {Name:"UNISUPER",Street:"Level 37 385 Burke", State:"Street Melbourne  Vic  3000"}
		SuperDetails["OTN013"] := {Name:"HESTA",Street:"PO Box 600", State:"CARLTON SOUTH VIC 3053"}
		SuperDetails["OTN015"] := {Name:"AMP Superleader Plan",Street:"Locked Bag 5095", State:"Parramatta, NSW 2124"}
		SuperDetails["OTN022"] := {Name:"Westpac Lifetime Super",Street:"GPO BOX 3960", State:"Sydney  NSW  2001"}
		SuperDetails["OTN024"] := {Name:"C+BUS",Street:"Locked Bag 999", State:"Carlton South  VIC 3053"}
		SuperDetails["OTN026"] := {Name:"Host Plus Superannuation Fund",Street:"Locked Bag 999", State:"Carlton South  Vic  3053"}
		SuperDetails["OTN027"] := {Name:"MTAA Superannuation Fund",Street:"Locked Bag 15", State:"HayMarket  NSW 1236"}
		SuperDetails["OTN028"] := {Name:"TAL Superannuation and Insurance Fund",Street:"PO Box 142", State:"Milsons Point NSW 1565"}
		SuperDetails["OTN030"] := {Name:"TWU Superannuation Fund",Street:"GPO Box 4689", State:"Melbourne  VIC  8060"}
		SuperDetails["OTN031"] := {Name:"Media Super",Street:"Locked Bag 1229", State:"Wollongong  NSW  2500"}
		SuperDetails["OTN032"] := {Name:"MLC Masterkey Superannuation Fund",Street:"PO Box 1315", State:"North Sydney  NSW  2059"}
		SuperDetails["OTN033"] := {Name:"ASGARD Independence Plan Division One",Street:"PO Box 7490", State:"Cloisters Square  WA  6850"}
		SuperDetails["OTN036"] := {Name:"Sunsuper Superannuation Fund",Street:"PO Box 2924", State:"Brisbane  QLD 4001"}
		SuperDetails["OTN037"] := {Name:"CARE Superannuation",Street:"Locked Bag 5087", State:"Parramatta NSW 2124"}
		SuperDetails["OTN039"] := {Name:"Asteron Life Ltd",Street:"GPO Box 1576", State:"Sydney  NSW 2001"}
		SuperDetails["OTN041"] := {Name:"SYNERGY Superannuation Fund",Street:"GPO Box 852", State:"Hobart  TAS 7001"}
		SuperDetails["OTN042"] := {Name:"The Portfolio Service",Street:"Locked Bag 800", State:"Milsons Point NSW 1565"}
		SuperDetails["OTN043"] := {Name:"VICSUPER",Street:"PO Box 89", State:"Melbourne  Vic  3001"}
		SuperDetails["OTN044"] := {Name:"BUSS(Q) Superannuation Fund",Street:"PO Box 902", State:"SPRING HILL QLD 4004"}
		SuperDetails["OTN046"] := {Name:"BT Life Super",Street:"BT Financial Group", State:"GPO Box 2675"}
		SuperDetails["OTN048"] := {Name:"ING Life Limited",Street:"GPO Box 5306", State:"Sydney  NSW  2001"}
		SuperDetails["OTN050"] := {Name:"Zurich Australia Superannuation",Street:"PO Box 994", State:"North Sydney  NSW  2059"}
		SuperDetails["OTN051"] := {Name:"ANZ Superannuation",Street:"GPO BOX 4028", State:"Sydney  NSW 2001"}
		SuperDetails["OTN052"] := {Name:"MAP Personal Pension Plan",Street:"PO Box 1130", State:"Brisbane  QLD 4001"}
		SuperDetails["OTN056"] := {Name:"APEX SUPER",Street:"GPO Box  398", State:"NORTH SYDNEY NSW 2059"}
		SuperDetails["OTN057"] := {Name:"Australian Ethical Retail Superannuation Fund (aka Australian Ethical Super Fund)",Street:"Attn: Emily Price", State:"GPO Box 529"}
		SuperDetails["OTN058"] := {Name:"Statewide Superannuation Trust",Street:"GPO Box 1749", State:"Adelaide  SA  5001"}
		SuperDetails["OTN059"] := {Name:"Advance Super",Street:"GPO Box B87", State:"Perth  WA  6838"}
		SuperDetails["OTN060"] := {Name:"Guild Super",Street:"GPO BOX 1088", State:"MELBOURNE  VIC  3001"}
		SuperDetails["OTN063"] := {Name:"Catholic Superannuation Fund",Street:"PO Box 2163", State:"Melbourne  Vic  3001"}
		SuperDetails["OTN064"] := {Name:"TASPLAN Super",Street:"PO Box 1547", State:"Hobart  TAS 7001"}
		SuperDetails["OTN065"] := {Name:"AUSTSAFE Superannuation Fund",Street:"GPO Box 3113", State:"Brisbane  QLD  4001"}
		SuperDetails["OTN066"] := {Name:"Colonial Mutual Super (aka Colonial Mutual Life Assurance)",Street:"PO Box 320", State:"Silverwater  NSW  2128"}
		SuperDetails["OTN067"] := {Name:"Australian Primary Superannuation Fund",Street:"Locked Bag 2229", State:"Wollongong DC NSW 2500"}
		SuperDetails["OTN068"] := {Name:"Summit Master Trust Personal Superannuation & Pension Plan",Street:"GPO Box 2754", State:"MELBOURNE  VIC  3001"}
		SuperDetails["OTN069"] := {Name:"Colonial Master Fund (aka Colonial Select Superannuation Fund)",Street:"Locked Bag 5075", State:"Parramatta  NSW 2124"}
		SuperDetails["OTN070"] := {Name:"AustralianSuper Corporate",Street:"GPO Box 4303", State:"Melbourne VIC 3001"}
		SuperDetails["OTN072"] := {Name:"ANZ Super Advantage",Street:"GPO BOX 4028", State:"Sydney  NSW 2001"}
		SuperDetails["OTN073"] := {Name:"Finium Super Master Plan",Street:"GPO Box 529", State:"Hobart  TAS  7001"}
		SuperDetails["OTN080"] := {Name:"SimpleWRAP",Street:"GPO Box 2945, Melbourne VIC 3001", State:""}
		SuperDetails["OTN081"] := {Name:"Health Industry Plan (HIP)",Street:"Locked Bag 20 Grosvenor Pl", State:"Sydney  NSW 1216"}
		SuperDetails["OTN083"] := {Name:"ASSET (Australian Superannuation Savings Employment Trust)",Street:"GPO Box 4030", State:"Sydney  NSW 2001"}
		SuperDetails["OTN090"] := {Name:"Symetry Lifetime Super",Street:"Locked Bag 3460", State:"Melbourne VIC 3001"}
		SuperDetails["OTN091"] := {Name:"Navigator Super Solutions",Street:"GPO Box 2567", State:"Melbourne Vic 3001"}
		SuperDetails["OTN092"] := {Name:"Plan B Superannuation Fund",Street:"Plan B Wealth Management", State:"PO Box 7008"}
		SuperDetails["OTN095"] := {Name:"Non Government Schools Superannuation Fund",Street:"NGS Super", State:"GPO Box 4303"}
		SuperDetails["OTN102"] := {Name:"Colonial Personal Super Fund",Street:"Locked Bag 5790", State:"Parramatta  NSW  2124"}
		SuperDetails["OTN103"] := {Name:"Colonial First State First Choice",Street:"GPO Box 3956", State:"Sydney NSW 2001"}
		SuperDetails["OTN107"] := {Name:"Retirement Portfolio Services",Street:"Locked Bag 50", State:"Australian Square NSW 1215"}
		SuperDetails["OTN109"] := {Name:"AustChoice Super Plan",Street:"Attn: Emily Price", State:"GPO Box 529"}
		SuperDetails["OTN117"] := {Name:"Christian Super",Street:"Locked Bag 5073", State:"Parramatta NSW 2124"}
		SuperDetails["OTN119"] := {Name:"AustralianSuper Industry",Street:"GPO Box 1901R", State:"Melbourne Vic  3000"}
		SuperDetails["OTN122"] := {Name:"New Name: MLC Navigator Retirement Plan",Street:"GPO Box 2567W", State:"Melbourne  Vic  3001"}
		SuperDetails["OTN123"] := {Name:"Strategy Retirement Fund",Street:"Locked Bag 1000", State:"Wollongong NSW 2500"}
		SuperDetails["OTN125"] := {Name:"QIEC Superannuation Trust",Street:"PO Box 2130", State:"MILTON QLD 4064"}
		SuperDetails["OTN127"] := {Name:"Recruitment Services Superannuation Fund",Street:"GPO Box 4839VV", State:"Melbourne  VIC  3001"}
		SuperDetails["OTN129"] := {Name:"WESTPAC Super Investment Fund",Street:"GPO Box 2362", State:"Adelaide  SA  5001"}
		SuperDetails["OTN130"] := {Name:"Suncorp Metway Easy Super",Street:"PO Box 1453", State:"Brisbane  QLD 4001"}
		SuperDetails["OTN131"] := {Name:"Perpetual Investor Choice Retirement Fund (aka Perpetual Wealth Focus Super Plan)",Street:"GPO Box 4171", State:"Sydney NSW 2001"}
		SuperDetails["OTN137"] := {Name:"Bendigo Super Plan",Street:"GPO Box 529", State:"Hobart  TAS  7001"}
		SuperDetails["OTN138"] := {Name:"AMP Retirement Savings Account",Street:"LOCKED BAG 5400", State:"PARAMATTA  NSW  1741"}
		SuperDetails["OTN139"] := {Name:"AMP Miscellaneous",Street:"LOCKED BAG 5400", State:"PARAMATTA NSW 1741"}
		SuperDetails["OTN140"] := {Name:"AMP Flexible Lifetime Super",Street:"LOCKED BAG 5400,", State:"PARAMATTA, NSW, 1741"}
		SuperDetails["OTN143"] := {Name:"AustralianSuper Industry",Street:"GPO Box 1901R", State:"Melbourne Vic  3000"}
		SuperDetails["OTN146"] := {Name:"Managed Australian Retirement Fund",Street:"PO Box 7074", State:"East Brisbane  QLD 4169"}
		SuperDetails["OTN148"] := {Name:"Colonial Super Retirement Fund",Street:"PO Box 320", State:"Silverwater  NSW  2128"}
		SuperDetails["OTN150"] := {Name:"Vision Super (aka Local Authorities Super)",Street:"PO Box 18041 Collins Street", State:"East Melbourne  VIC  8003"}
		SuperDetails["OTN153"] := {Name:"Australian Catholic Superannuation & Retirement Fund",Street:"PO Box 656", State:"Burwood NSW 1805"}
		SuperDetails["OTN161"] := {Name:"Mercer Super Trust",Street:"GPO Box 4303", State:"Melbourne  VIC   3001"}
		SuperDetails["OTN162"] := {Name:"Combined Fund",Street:"GPO Box 4559", State:"Melbourne  VIC  3001"}
		SuperDetails["OTN167"] := {Name:"IOOF Portfolio",Street:"GPO Box 264C", State:"Melbourne VIC 3001"}
		SuperDetails["OTN175"] := {Name:"Mentor Superannuation Master Trust",Street:"Matrix Superannuation Fund", State:"Locked Bag 1000"}
		SuperDetails["OTN180"] := {Name:"Club Super",Street:"PO Box 2239", State:"Milton QLD 4064"}
		SuperDetails["OTN190"] := {Name:"SuperWrap Superannuation",Street:"GPO Box 2337", State:"Adelaide SA 5001"}
		SuperDetails["OTN196"] := {Name:"Plum Superannuation Fund",Street:"GPO Box 63A", State:"Melbourne VIC 3000"}
		SuperDetails["OTN199"] := {Name:"Intrust Superannuation Fund",Street:"GPO Box 1416", State:"BRISBANE QLD 4001"}
		SuperDetails["OTN203"] := {Name:"Masterkey Custom Superannuation (aka Flexiplan Australia Ltd and aka HML Super fund - One source)",Street:"PO Box 7657", State:"Cloister Square WA 6850"}
		SuperDetails["OTN211"] := {Name:"FSP Superannuation Fund",Street:"FSP Super Pty Limited", State:"Locked Bag 3460"}
		SuperDetails["OTN223"] := {Name:"WA Local Government Superannuation Plan",Street:"PO Box Z5493", State:"St Georges Terrace"}
		SuperDetails["OTN224"] := {Name:"Maritime Super - Seafarers",Street:"Level 4, 6 Riverside Quay", State:"Southbank  VIC  3006"}
		SuperDetails["OTN230"] := {Name:"Medical & Associated Professions Superannuation Fund",Street:"c/- AMA", State:"14 Stirling Highway"}
		SuperDetails["OTN232"] := {Name:"AMP Custom Super",Street:"LOCKED BAG 5400,", State:"PARAMATTA, NSW, 1741"}
		SuperDetails["OTN242"] := {Name:"LUCRF (Labour Union Co-Operative Retirement Fund)",Street:"PO Box 211", State:"North Melbourne VIC 3051"}
		SuperDetails["OTN243"] := {Name:"Telstra Superannuation Scheme",Street:"PO Box 14309", State:"Melbourne VIC 8001"}
		SuperDetails["OTN249"] := {Name:"FuturePlus Super Fund",Street:"GPO Box 2617", State:"Sydney NSW 2001"}
		SuperDetails["OTN254"] := {Name:"IPAC Iaccess Personal Super Plan",Street:"C/- Investor Services", State:"GPO BOX 2754"}
		SuperDetails["OTN286"] := {Name:"Nationwide Superannuation Fund- NSF Super",Street:"PO Box 42", State:"Charlestown NSW 2290"}
		SuperDetails["OTN287"] := {Name:"Australian Enterprise Superannuation Fund (AESuper)",Street:"GPO Box 2258", State:"Melbourne  VIC  3001"}
		SuperDetails["OTN288"] := {Name:"Russell SuperSolutions",Street:"Locked Bag A4094", State:"Sydney South NSW 1235"}
		SuperDetails["OTN289"] := {Name:"IAccess Personal Superannuation",Street:"PO Box 471  Collins Street", State:"WEST MELBOUNRE VIC 8007"}
		SuperDetails["OTN291"] := {Name:"legalsuper",Street:"GPO Box 263C", State:"Melbourne  VIC  3001"}
		SuperDetails["OTN296"] := {Name:"The Duncan Superannuation Fund",Street:"PO Box 1427", State:"KATHERINE NT 0851"}
		SuperDetails["OTN312"] := {Name:"Virgin Superannuation Fund",Street:"Locked Bag 8", State:"Haymarket NSW 1236"}
		SuperDetails["OTN314"] := {Name:"Maritime Super - Stevedores",Street:"Level 4, 6 Riverside Quay", State:"Southbank  VIC  3006"}
		SuperDetails["OTN324"] := {Name:"LOCAL SUPER - Division of StatewideSuper)",Street:"PO Box 7035", State:"Hutt Street"}
		SuperDetails["OTN330"] := {Name:"The Executive Superannuation Fund",Street:"10 Shelly Street", State:"Sydney  NSW  2000"}
		SuperDetails["OTN331"] := {Name:"JR Childs PortfolioOne Superannuation Fund (only for AGS 77753756)",Street:"Locked Bag 50", State:"Australia Square NSW 1215"}
		SuperDetails["OTN334"] := {Name:"AON Master Trust",Street:"PO BOX 9819", State:"Sydney NSW 2001"}
		SuperDetails["OTN344"] := {Name:"Credit Suisse Asset Management Super MasterWrap",Street:"PO Box R240", State:"Royal Exchange NSW 1225"}
		SuperDetails["OTN346"] := {Name:"First State Super",Street:"PO Box 1229", State:"Wollongong  NSW  2500"}
		SuperDetails["OTN357"] := {Name:"REI Super Fund",Street:"GPO Box 4303", State:"Melbourne VIC 3001"}
		SuperDetails["OTN365"] := {Name:"Spectrum Super Fund",Street:"GPO BOX 529", State:"HOBART   TAS   7001"}
		SuperDetails["OTN366"] := {Name:"Lifetime Superannuation Fund",Street:"PO BOX 7008", State:"CLOISTERS SQUARE WA 6850"}
		SuperDetails["OTN371"] := {Name:"The Corporate Superannuation Master Trust - Portable Plan",Street:"PO Box 1647", State:"MILTON BC  QLD  4064"}
		SuperDetails["OTN396"] := {Name:"Mentor Superannuation Master Trust",Street:"Locked Bag 1000", State:"Wollongong NSW 2500"}
		SuperDetails["OTN402"] := {Name:"Energy Super formerly ESI Super",Street:"PO Box 1958", State:"Milton  QLD  4064"}
		SuperDetails["OTN421"] := {Name:"Netwealth Superannuation Master Trust Fund",Street:"Netwealth Investments Ltd", State:"Level 7/52 Collins Street"}
		SuperDetails["OTN431"] := {Name:"LG Super",Street:"GPO Box 264                                       Brisbane  QLD  4001", State:""}
		SuperDetails["OTN433"] := {Name:"Equipsuper Pty Ltd",Street:"Level 12, 114 William Street", State:"MELBOURNE VIC 3000"}
		SuperDetails["OTN435"] := {Name:"PortfolioCare eWRAP Super Account",Street:"PortfolioCare", State:"GPO BOX C113"}
		SuperDetails["OTN457"] := {Name:"AXA Retirement Bond",Street:"PO Box 14330", State:"Melbourne  VIC  8001"}
		SuperDetails["OTN482"] := {Name:"NAB Group Superannuation Fund A (previous NAB employees) (aka NAB Retained Benefit Account)",Street:"PO BOX 1321", State:"MELBOURNE VIC 3001"}
		SuperDetails["OTN507"] := {Name:"Garnet Superannuation Fund",Street:"GPO Box 4369", State:"Melbourne  VIC  3001"}
		SuperDetails["OTN512"] := {Name:"AvSuper",Street:"GPO Box 367 Canberra ACT 2601", State:""}
		SuperDetails["OTN520"] := {Name:"State Super Retirment Fund (aka State Super Personal Retirment Plan)",Street:"State Super Financial Services Pty Ltd", State:"GPO BOX 5336"}
		SuperDetails["OTN522"] := {Name:"ANZ OneAnswer Personal Super",Street:"ING Custodians Pty Limited", State:"GPO Box 4028"}
		SuperDetails["OTN531"] := {Name:"AMP SignatureSuper",Street:"C/- AMP Life Limited", State:"Locked Bag 5043"}
		SuperDetails["OTN540"] := {Name:"Tauondi Rhema Super Fund",Street:"PO BOX 40559", State:"Casuarina NT 0810"}
		SuperDetails["OTN544"] := {Name:"DIY Master Plan",Street:"PO BOX 361", State:"Collins Street West"}
		SuperDetails["OTN553"] := {Name:"Navigator Access Super and Pension",Street:"Navigator Australia Limited", State:"GPO BOX 2567W"}
		SuperDetails["OTN575"] := {Name:"ESS Super",Street:"GPO BOX 1974", State:"MELBOURNE VIC 3001"}
		SuperDetails["OTN576"] := {Name:"BHP Billiton Super Fund",Street:"GPO BOX 63", State:"MELBOURNE VIC 3001"}
		SuperDetails["OTN577"] := {Name:"AMG Universal Super",Street:"GPO BOX 330", State:"BRISBANE QLD 4000"}
		SuperDetails["OTN598"] := {Name:"Wealthtrac",Street:"Locked Bag 1000", State:"WOLLONGONG  NSW  2500"}
		SuperDetails["OTN604"] := {Name:"PLUM PERSONAL PLAN (previously Vanguard Personal Superannuation Plan)",Street:"PLUM FINANCIAL SERVICES LIMITED", State:"REPLY PAID 63"}
		SuperDetails["OTN621"] := {Name:"TIC Retirement Plan",Street:"TIC Retirement Plan", State:"PO Box 1282"}
		SuperDetails["OTN628"] := {Name:"Prime Super",Street:"PO Box 2229", State:"Wollongong  NSW  2500"}
		SuperDetails["OTN647"] := {Name:"MLC MasterKey Business Super",Street:"MLC", State:"PO Box 200"}
		SuperDetails["OTN653"] := {Name:"QSuper",Street:"GPO Box 200, Brisband QLD 4001", State:""}
		SuperDetails["OTN659"] := {Name:"AUST(Q)",Street:"PO Box 329", State:"SPRING HILL QLD 4004"}
		SuperDetails["OTN660"] := {Name:"North - Wealth Personal Superannuation + Pension Plan",Street:"North Service Centre", State:"GPO Box 2915 Melbourne, VIC 3001"}
		SuperDetails["OTN664"] := {Name:"AMIST Super (Australian Meat Industry Super Trust)",Street:"1A Homebush Bay Drive", State:"Rhodes  NSW  2138"}
		SuperDetails["OTN667"] := {Name:"Local Government Superannuation Scheme",Street:"PO Box N835", State:"Grosvenor Place NSW 1220"}
		SuperDetails["OTN674"] := {Name:"Club Plus Superannuation",Street:"Locked Bag 5007", State:"Parramatta NSW 2124"}
		SuperDetails["OTN714"] := {Name:"MLC Wrap Super",Street:"GPO BOX 2567", State:"MELBOURNE  VIC  3001"}
		SuperDetails["OTN723"] := {Name:"FirstWrap Super",Street:"Avanteos Superannuation Trust Fund", State:"FirstWrap"}
		SuperDetails["OTN749"] := {Name:"Water Corporation Superannuation Plan",Street:"PO Box 241", State:"West Perth WA 6872"}
		SuperDetails["OTN750"] := {Name:"Super Directions for Business",Street:"PO Box 14669", State:"Melbourne VIC 8001"}
		SuperDetails["OTN752"] := {Name:"Integra Super",Street:"GPO BOX 5306", State:"Sydney NSW 2001"}
		SuperDetails["OTN766"] := {Name:"Super Directions Personal Super Plan",Street:"AXA Australia", State:"Custormer Service"}
		SuperDetails["OTN785"] := {Name:"Emplus Superannuation",Street:"PO Box 3528", State:"TINGALPA DC, QLD 4173"}
		SuperDetails["OTN856"] := {Name:"Medical & Associated Professions Superannuation Fund",Street:"GPO Box 529", State:"Hobart TAS 7001"}
		SuperDetails["OTN865"] := {Name:"Voyage Superannution Master Trust",Street:"Locked Bag 1000", State:"Wollongong DC NSW 2500"}
		SuperDetails["OTN873"] := {Name:"ING Direct Living Superannuation Fund",Street:"Reply Paid 4307", State:"Sydney NSW 2001"}
		SuperDetails["OTN876"] := {Name:"BT Superannuation Investment Fund",Street:"GPO Box 2675 Sydney NSW 2001", State:""}
		SuperDetails["OTN878"] := {Name:"Compass Ewarp",Street:"PO Box 7241", State:"Cloisters Square Perth WA 6839"}
		SuperDetails["OTN891"] := {Name:"Quadrant Superannuation",Street:"GPO Box 863", State:"Hobart TAS 7001"}
		SuperDetails["OTN907"] := {Name:"Qantas Staff Credit Union Retirement Savings Account",Street:"420 Forest Road", State:"Hurstville NSW 2220"}
		SuperDetails["OTN912"] := {Name:"Police Credit SuperFuture RSA",Street:"PO Box 669", State:"CARTON SOUTH VIC 3053"}
		SuperDetails["OTN948"] := {Name:"Mercy Super",Street:"PO Box 8334", State:"Woolloongabba QLD 4102"}
		
		WinGet IDVar, ID, ahk_exe tn3270.exe  
		If !IDVar {
			msgbox % "Error: PIPS Not Found.`n`nPlease open PIPS then try again."
			Exit
		}
		
		
		This.ID := IDVar  
		This.Activate()
		sendinput, {F4}
		
		if (This.GetWindowState() = "TIMEOUTCHECK") 
		{
			This.Sleep(100)
			sendinput,{enter}
			This.Sleep(100)
			sendinput,{enter}
			This.Sleep(100)
			sendinput,{enter}
			This.Sleep(200)
		}
		
		This.Sleep(1000)
		
		if (This.GetWindowState() = "MAINMENUCHECK") 
		{				
			This.Sleep(100)
			sendinput, 1.2.1
			This.Sleep(100)
			sendinput,{enter}
			This.Sleep(100)
			sendinput % this.ServiceNo 
			This.Sleep(100)
			sendinput, {enter}
			This.Sleep(200)
		}
		
		if (This.GetWindowState() = "SEPARATEDCHECK") 
		{
			This.Sleep(100)
			sendinput, {enter}
			This.Sleep(200)
		}
		
		This.Sleep(1000)
		
		if (This.GetWindowState() = "CasualCheck") 
		{	
			Array := This.MainframeToTextArray()
			This.Sleep(1000)
			
			Client_Info.Person_Last_Name := substr(Array[6],22,23)
			Client_Info.Person_Given_Name := substr(Array[7],30,45)
			Client_Info.Person_Date_Of_Birth := substr(Array[12],30,12)
			Client_Info.Person_Pay_Centre := substr(Array[18],30,3)
			Client_Info.Person_Courtesy_Title := substr(Array[9],30,4)
		}
		
		else if (This.GetWindowState() = "TempCheck") 
		{
			Array := This.MainframeToTextArray()
			This.Sleep(1000)
			Client_Info.Person_Last_Name := substr(Array[6],22,23)
			Client_Info.Person_Given_Name := substr(Array[7],22,45)
			Client_Info.Person_Date_Of_Birth := substr(Array[11],22,12)
			Client_Info.Person_Pay_Centre := substr(Array[12],22,4)
			Client_Info.Person_Courtesy_Title := substr(Array[9],22,11)
		}	
		
		This.Sleep(200)						 		
		sendinput, {F4}
		sendinput, {TAB}1.2.3
		sendinput, {Enter}{F8}
		This.Sleep(1000)
		
		if (This.GetWindowState() = "SEPARATEDCHECK") 
		{
			This.Sleep(100)
			sendinput, {enter}
			This.Sleep(200)
		}
		This.Sleep(1000)
		
		Array := This.MainframeToTextArray()
		
		if (substr(Array[3],37,9) = "Addresses" and substr(Array[16],66,5) != "Other")
		{
			Client_Info.Address_Postal_Line_1 := substr(Array[12],18,21)
			Client_Info.Address_Postal_Line_2 := substr(Array[13],18,21) 
			Client_Info.Address_Postal_Line_3 := substr(Array[14],18,21) 
			Client_Info.Address_Postal_Suburb := substr(Array[15],18,21) 
			Client_Info.Address_Postal_State := substr(Array[15],46,3) 
			Client_Info.Address_Postal_PostCode := substr(Array[16],18,4) 
			Client_Info.Address_Personal_Mobile := substr(Array[10],61,20) 
			Client_Info.Address_Personal_Email := substr(Array[11],18,50)
			Client_Info.Address_Work_Email := substr(Array[22],18,50) 
		}
		This.Sleep(1000)
		
		if (substr(Array[3],37,9) = "Addresses" and substr(Array[16],66,5) = "Other")
		{
			Client_Info.Address_Personal_Mobile := substr(Array[10],18,20) 
			if (Client_Info.Address_Personal_Mobile = "")
				Client_Info.Address_Personal_Mobile := substr(Array[10],61,20) 
			Client_Info.Address_Personal_Email := substr(Array[11],18,50) 
			Client_Info.Address_Work_Email := substr(Array[22],18,50) 
			sendinput, {Enter}
			This.Sleep(1000)
			Array := This.MainframeToTextArray()
			Client_Info.Address_Postal_Line_1 := substr(Array[12],28,24) 
			Client_Info.Address_Postal_Line_2 := substr(Array[13],28,24) 
			Client_Info.Address_Postal_Line_3 := substr(Array[14],28,24)
			Client_Info.Address_Postal_Suburb := substr(Array[15],28,17) 
			Client_Info.Address_Postal_State := substr(Array[15],56,3) 
			Client_Info.Address_Postal_PostCode := substr(Array[16],28,4) 
		}
		
		Array := ""
		This.Sleep(1000)
		sendinput, {F4}
		sendinput, {TAB}1.2.7
		sendinput, {Enter}{F8}
		This.Sleep(1000)
		
		if (This.GetWindowState() = "SEPARATEDCHECK") 
		{
			This.Sleep(400)
			sendinput, {enter}
			This.Sleep(400)
		}
		
		This.Sleep(1000)
		Array := This.MainframeToTextArray()
		This.Sleep(1000)
		
		this.Activate()
		if (substr(Array[2],3,7) = "(1.2.7)")
		{
			Loop 6
			{
				Test := (InStr(substr(Array[10+A_Index],3,4), "A") ? A_Index : Return)
				if (Test = "")
					Break
				else
					Row := Test
			}
			
			sendinput, %Row% {enter}
			this.Sleep(1000)
		}
		
		Array := This.MainframeToTextArray()
		
		if (substr(Array[3],29,52) = "    Superannuation          - Choice   (Enquiry)    ") {
			this.Sleep(100)
			Client_Info.Super_Member_No := StrReplace(SubStr(Array[11], 55 , 20), "_") 
			Client_Info.Super_Fund_ID	:= SubStr(Array[11], 23 , 7) 
			
			Client_Info.Super_Fund_Name := SuperDetails[TRIM(Client_Info.Super_Fund_ID)].Name
			Client_Info.Super_Fund_Street := SuperDetails[TRIM(Client_Info.Super_Fund_ID)].Street
			Client_Info.Super_Fund_State := SuperDetails[TRIM(Client_Info.Super_Fund_ID)].State			
		}
		else {
			Client_Info.Super_Member_No := "Other - N/A"
			Client_Info.Super_Fund_ID	:= "Other - N/A"
		}
		
		ServiceNo := this.ServiceNo
		winactivate, ePOD
		This.sleep(200)
		WinGet, hWnd, ID, ePOD
		oAcc := Acc_Get("Object", "4.4.4", 0, "ahk_id " hWnd) 
		ControlHwnd := Acc_WindowFromObject(oAcc)
		This.sleep(200)
		ControlFocus,, ahk_id %ControlHwnd%
		This.sleep(200)
		ControlGetFocus, ControlName, ahk_id %ControlHwnd%
		This.sleep(200)
		dllcall("keybd_event", int, 0x25, int, 0x14B, int, 0, int, 0)
		This.sleep(200)
		send, {AppsKey}{Down}{Down}{Enter}
		This.sleep(200)
		Send, %ServiceNo%{ENTER}
		This.sleep(200)
		oAcc := Acc_Get("Object", "4.1.4.1.4.16.4", vChildID, "ahk_id " hWnd)
		Client_Info.Person_Cost_Code := oAcc.accValue(vChildID)
		oAcc := Acc_Get("Object", "4.1.4.1.4.18.4", vChildID, "ahk_id " hWnd)
		Client_Info.PayCentre := oAcc.accValue(vChildID)
		;msgbox % Client_Info.PayCentre
 

		
		Clipboard := Client_Info.ServiceNo
		
		for key, value in Client_Info {
			Client_Info[key] := TRIM(value)
			If (Client_Info[key] = "")
				Client_Info.Delete(Key)
		}
		return Client_Info
	}
	
	FinalEntsHotKey() {
		FinalEntsLine := StrSplit(clipboard,A_Tab)
		FormatTime, TimeString,, shortDate
		WinActivate Mochasoft - mainframe.nt.gov.au
		sleep 200
		SetKeyDelay, 90, 0,
		
		If (FinalEntsLine[2] = "ENT.PYM")
			send % "{F7}" FinalEntsLine[2] "{enter}" FinalEntsLine[3] TimeString "{tab}{tab}"  FinalEntsLine[4] := RegExReplace(FinalEntsLine[4], "[^0-9.]") "{tab}{+}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}Y{F9}"  FinalEntsLine[5] := FinalEntsLine[5]="" ? FinalEntsLine[1]:FinalEntsLine[1] " - " FinalEntsLine[5]
		If (FinalEntsLine[2] = "LVB.PYM") 
			send % "{F7}" FinalEntsLine[2] "{enter}" FinalEntsLine[3]FinalEntsLine[4] := RegExReplace(FinalEntsLine[4], "[^0-9.]") "{tab}{+}" StrSplit(FinalEntsLine[1],A_Space).1 "{tab}{tab}{tab}y{F9}"  FinalEntsLine[5] := FinalEntsLine[5]="" ? FinalEntsLine[1]:FinalEntsLine[1] " - " FinalEntsLine[5]
		If (FinalEntsLine[2] = "TAX.PYM")
			send % "{F7}" FinalEntsLine[2] "{enter}FTTAX1" FinalEntsLine[4] := RegExReplace(FinalEntsLine[4], "[^0-9.]") "{TAB}{TAB}Y{F9}" FinalEntsLine[5] := FinalEntsLine[5]="" ? FinalEntsLine[1]:FinalEntsLine[1] " - " FinalEntsLine[5]
		If (FinalEntsLine[1] = "RCY.PYM")
			send % "{F7}" FinalEntsLine[1] "{enter}rcy002n" TimeString "{TAB}" FinalEntsLine[3] := RegExReplace(FinalEntsLine[3], "[^0-9.]") "{TAB}{TAB}{TAB}Y{F9}Final Ents offset against salary overpayment."

		If (FinalEntsLine[2] = "EMP.PYM") {
			InputBox, Ref,,What is the Member Number?,,150,150		
			WinActivate Mochasoft - mainframe.nt.gov.au
			send % "{F7}" FinalEntsLine[2] "{enter}" FinalEntsLine[3] FinalEntsLine[4] := RegExReplace(FinalEntsLine[4], "[^0-9.]") "{TAB}{+}{TAB 5}" Ref "{TAB}{TAB}{TAB}y{F9}{F9}Super Due on Final Ents" ;{F9}Super Due on Final Ents"
			}
		}
		
		AdjustCumlativeTotals(Client_Info)
		{
			
			This.Sleep(200)
			
			Workbook_Reason_PIPS :=  Client_Info.Reason
			if winexist("Mochasoft - mainframe.nt.gov.au")
			{
				
				WinActivate Mochasoft
				SetKeyDelay, 80, 0,
				Datez := A_DD " " A_MMM " " A_YYYY
				This.Sleep(200)
				Client_Info.RPMENT := RegExReplace(Client_Info.RPMENT, "[^0-9.]")
				Client_Info.TAXPYM := RegExReplace(Client_Info.TAXPYM, "[^0-9.]")
				Client_Info.RDRENT := RegExReplace(Client_Info.RDRENT, "[^0-9.]")
				Amount_Field := "_________"
				Amount_Field_Larg := "________________"
				Date_Field := "___________"
				Receipt_Field := "________________"
				Receipt := "PAYROLL DRT ASP_"
				This.Sleep(500)
				Send {F7}RPM.ENT{enter}
				Clipboard := SubStr(Amount_Field,1,StrLen(Amount_Field) - StrLen(Client_Info.RPMENT)) . Client_Info.RPMENT . "+" . Datez . Receipt . "__" . SubStr(Amount_Field,1,StrLen(Amount_Field) - StrLen(Client_Info.RPMENT)) . Client_Info.RPMENT . "M" . "________________NY" . "_______________________" . This.Float(Client_Info.PayCentre, 3) . "2021/22" . Client_Info.Pay_Cycle
				This.Sleep(500)
				SendInput ^v
				This.Sleep(500)
				Workbook_Reason_PIPS := SubStr(Client_Info.Reason, 1, 160)
				SendInput {f9}%Workbook_Reason_PIPS%{Enter}{Enter}{Enter}
				This.Sleep(500)
				SendInput {F7}TAX.PYM{enter}
				This.Sleep(500)
				Clipboard := "ATTAX1" . SubStr(Amount_Field,1,StrLen(Amount_Field) - StrLen(Client_Info.TAXPYM)) . Client_Info.TAXPYM . "NY" . "_______________________" . This.Float(Client_Info.PayCentre, 3) . "2021/22" . Client_Info.Pay_Cycle
				SendInput ^v	
				This.Sleep(500)
				Send {F17}
				This.Sleep(500)
				Workbook_Reason_PIPS := SubStr(Client_Info.Reason, 1, 160)
				SendInput {f9}%Workbook_Reason_PIPS%{Enter}{Enter}{Enter}
				This.Sleep(500)
				Send {F7}RDR.ENT{enter}
				Clipboard  := "RPAY__" . SubStr(Amount_Field_Larg,1,StrLen(Amount_Field_Larg) - StrLen(Client_Info.TAXPYM)) . Client_Info.TAXPYM . "M" . Datez . "___________________________________" . "Y" . "_______________________" . This.Float(Client_Info.PayCentre, 3) . "2021/22" . Client_Info.Pay_Cycle
				This.Sleep(500)
				SendInput ^v
				This.Sleep(500)		
				Workbook_Reason_PIPS := SubStr(Client_Info.Reason, 1, 160)
				SendInput {f9}%Workbook_Reason_PIPS%{Enter}{Enter}{Enter}
				return
			}
			else, 
			{
				msgbox, % "Is PIPS Open to PTR Screen for AGS: " SubStr(Client_Info.Active_AGS, 1 , 8) "? `n`nYou Will Need to Run Cumulative Totals Adjustment Again."
				return
			}
			return
		}
		
		
		SGR(Client_Info){
			if winexist("Mochasoft - mainframe.nt.gov.au")
			{
				Amt := This.Float(RegExReplace((Client_Info.Super_Amount ? Client_Info.Super_Amount : 0) + (Client_Info.Previous_Fin_Year_Super ? Client_Info.Previous_Fin_Year_Super : 0), "[^0-9.]"))
				
				WinActivate Mochasoft
				SetKeyDelay, 80, 0,
				This.Sleep(500)
				Send {F7}SGR.COM{enter}
				
				Clipboard  :=  A_DD "/" A_MMM "/" A_YYYY . Amt . SubStr("_________",1,StrLen("_________")-StrLen(Amt)) . "Y_______________________" . This.Float(Client_Info.PayCentre,3) . "2021/22" . Client_Info.Pay_Cycle
				This.Sleep(500)
				SendInput ^v
				Workbook_Reason_PIPS := SubStr(Client_Info.Reason, 1, 160)
				SendInput {f9}%Workbook_Reason_PIPS%{Enter}{Enter}{Enter}
				
			}
			
		}
		EnterRecoveryAction(Client_Info)
		{
			if winexist("Mochasoft - mainframe.nt.gov.au")
			{
				Client_Info.RCYCOM := RegExReplace(Client_Info.RCYCOM, "[^0-9.]")
				
				IF (Client_Info.Percentage > Client_Info.RCYCOM)
					Client_Info.Percentage := Client_Info.RCYCOM
								
				WinActivate Mochasoft
				SetKeyDelay, 80, 0,
				Datez := A_DD " " A_MMM " " A_YYYY
				This.Sleep(200)
				formattime, datezz,a_now , dd MMM yyyy
				Amount_Field := "_________"
				Send {F7}RCY.COM{enter}
				This.Sleep(900)
				Clipboard  := "RCY002N" . datezz . SubStr(Amount_Field,1,StrLen(Amount_Field) - StrLen(Client_Info.Percentage)) . Client_Info.Percentage . "_________" . SubStr(Amount_Field,1,StrLen(Amount_Field) - StrLen(Client_Info.RCYCOM)) . Client_Info.RCYCOM . "______________" . "Y" . "_______________________" . This.Float(Client_Info.PayCentre,3) . "2021/22" . Client_Info.Pay_Cycle
				This.Sleep(900)
				SendInput ^v
				This.Sleep(900)
				Workbook_Reason_PIPS := SubStr(Client_Info.Reason, 1, 160)
				SendInput {f9}%Workbook_Reason_PIPS%
				This.Sleep(500)
				return
				
			}
			else, 
			{
				msgbox, % "Is PIPS Open to PTR Screen for AGS: " SubStr(Workbook_AGS, 1 , 8) "? `n`nYou Will Need to Run RCY.COM Again."
				return
			}
		}
		
		Create_Manual_Payee_Sheet(){
			winactivate, ePOD
			WinGet, hWnd, ID, ePOD
			
			oAcc := Acc_Get("Object", "4.1.4.1.4.8.4", vChildID, "ahk_id " hWnd)
			Last_Name := oAcc.accValue(vChildID)
			
			oAcc := Acc_Get("Object", "4.1.4.1.4.10.4", vChildID, "ahk_id " hWnd)
			First_Name := oAcc.accValue(vChildID)
			
			oAcc := Acc_Get("Object", "4.1.4.1.4.2.4", vChildID, "ahk_id " hWnd)
			AGS_No := oAcc.accValue(vChildID)
			
			oAcc := Acc_Get("Object", "4.1.4.1.4.12.4", vChildID, "ahk_id " hWnd)
			Amount_Due := oAcc.accValue(vChildID)
			
			Gui_Class.Application_Control_Workaround("C:\Users\babb\Documents\Custom Office Templates\Manual Payee Recovery.xls")
			WinWait, Manual Payee Recovery.xls  [Compatibility Mode] - Excel
			
			Workbook_Class.SetCellValue("Name:", First_Name " " Last_Name, 0, 1)
			Workbook_Class.SetCellValue("Name:", AGS_No, 1, 1)
			Workbook_Class.SetCellValue("Name:", Amount_Due, 2, 1)
			
			Xl := ComObjActive("Excel.Application")
			ComObjError(false)
			
			xl.ActiveSheet.SaveAs(A_Desktop . "\" AGS_No "_" This.Change_String(First_Name,"Title") "_" This.Change_String(Last_Name,"Title") ".xls")
			xl.ActiveWorkbook.Close
		}
		
		CreateePODRecord(Client_Info)
		{
			
			
			
			If (StrSplit(Client_Info.FBT_Date, "/").2 = "13") {
				InputBox, TempFBT, FBT Date Wrong:, What is the correct FBT Date?
				Client_Info.FBT_Date := TempFBT
			}	
			
			MsgBox 0x24, Create Combined Entry?, Create Combined Super and OVP?
			
			IfMsgBox Yes, {
				
				WinActivate ahk_exe Payback.exe
				WinGet,hWnd,id, Overpayment Details
				
				oAcc := Acc_Get("Object", "4.1.4.2.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, % Client_Info.Date_Detected, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, %  Client_Info.FBT_Date, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, RCY002, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.8.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Nett, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.10.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				;ControlSetText,, %  Client_Info.TotalAmount, ahk_id %ControlHwnd%
				
				Super_Amount := RegExReplace(Client_Info.Super_Amount, "[^0-9.]")
				Previous_Fin_Year_Super := RegExReplace(Client_Info.Previous_Fin_Year_Super, "[^0-9.]")
				
				EnvAdd,Super_Amount,0
				EnvAdd,Previous_Fin_Year_Super,0
				
				ControlSetText,, % Super_Amount+Previous_Fin_Year_Super+Client_Info.TotalAmount, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.16.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Source, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.19.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Type,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.22.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Cause,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.26.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Location, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Calculated/Reported, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				clipboard := "OVERPAID SUPER COMPONENT OF SALARY OVERPAYMENT ($" This.Float(Super_Amount+Previous_Fin_Year_Super) ") - " . Client_Info.Reason
				Send, ^v{Space}{BackSpace}
				
			} Else IfMsgBox No, {
				
				WinActivate ahk_exe Payback.exe
				WinGet,hWnd,id, Overpayment Details
				
				oAcc := Acc_Get("Object", "4.1.4.2.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, % Client_Info.Date_Detected, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, %  Client_Info.FBT_Date, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, RCY002, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.8.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Nett, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.10.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, %  Client_Info.TotalAmount, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.16.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Source, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.19.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Type,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.22.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Cause,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.26.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Location, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Calculated/Reported, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				clipboard := Client_Info.Reason
				Send, ^v{Space}{BackSpace}				
				
				
				
				
				
			}
			
			MsgBox 0x24, Create ePOD Super Entry?, Would you like to create an ePOD entry for Super?`n`nIf YES - Make Sure New ePOD record Window is Open.
			
			IfMsgBox Yes, {
				
				WinActivate ahk_exe Payback.exe
				WinGet,hWnd,id, Overpayment Details
				
				oAcc := Acc_Get("Object", "4.1.4.2.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, % Client_Info.Date_Detected, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, %  Client_Info.FBT_Date, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				ControlSetText,, %  Client_Info.Super_Fund_ID, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.8.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Nett, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.10.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				
				Super_Amount := RegExReplace(Client_Info.Super_Amount, "[^0-9.]")
				Previous_Fin_Year_Super := RegExReplace(Client_Info.Previous_Fin_Year_Super, "[^0-9.]")
				
				EnvAdd,Super_Amount,0
				EnvAdd,Previous_Fin_Year_Super,0
				
				ControlSetText,, % Super_Amount+Previous_Fin_Year_Super, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.16.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Source, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.19.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Type,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.22.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Error_Cause,, ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.1.4.26.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, % Client_Info.Location, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.6.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				Control, ChooseString, Calculated/Reported, , ahk_id %ControlHwnd%
				
				oAcc := Acc_Get("Object", "4.4.4", 0, "ahk_id " hWnd) 
				ControlHwnd := Acc_WindowFromObject(oAcc)
				ControlFocus,, ahk_id %ControlHwnd%
				clipboard := "OVERPAID SUPER COMPONENT OF SALARY OVERPAYMENT ($" This.Float(Super_Amount+Previous_Fin_Year_Super) ") - " . Client_Info.Reason
				Send, ^v{Space}{BackSpace}
				
			} Else IfMsgBox No, {
				return
			}
			Return
		}
		
		Excel_Sheet_to_PDF()
		{
			Xl := ComObjActive("Excel.Application")
			xl.ActiveSheet.ExportAsFixedFormat(0, Xl.activeworkbook.path . "\" . RegExReplace(XL.ActiveSheet.Name, "( Summary)", "_Paid Due Difference Sheet") . ".pdf" ,0,True,False,1,100,False)
			
		}
		
		Float( n, p:=6 ) 
		{ 
			Return SubStr(n:=Format("{:0." p "f}",n),1,-1-p) . ((n:=RTrim(SubStr(n,1-p),0) ) ? "." . n : "") 
		; By SKAN on D1BM @ goo.gl/Q7zQG9
		}
		
		Change_String(String, Options := "") 
		{
			
			if (Options = "Title") {
				StringUpper, string, string, T
				return % string
			}
			
			if (Options = "Upper") {
				StringUpper, string, string, 
				return % string
			}
			
			if (Options = "Lower") {
				StringLower, string, string, 
				return % string
			}
			
			if (Options = "Strip") {
				return RegExReplace(string, "[^0-9.]")
			}
			
			if (Options = "Clean") {
				string := regexreplace(string, "^\s+") 
				string := regexreplace(string, "\s+$") 
				return string
			}
			
			if (Options = "Plain") {
				string = %string%
				return string
			}
			MsgBox % "Option """ Options """ is not an Option. `n`nExiting Function." 
		}
		
		FBT_Date() 
		{
			Date := A_DD - 1 "/" A_MM + 1 "/" A_YYYY
			StringSplit, m, Date, `/
			v := (StrLen(m3)=2 ? "20" : "") m3 m2 m1		
			v += 0, Days
			return Date
		}
		
		
		
		PreparePaidDueDiff(Client_Info)
		{
			IniRead, Current_Financial_Year, % "C:\Users\babb\Desktop\Workbook Class Settings.ini", General, Current_Financial_Year
			Xl := ComObjActive("Excel.Application")
			ComObjError(false)
			
			For sheet in Xl.Worksheets
				if InStr(Sheet.Name, "Year Summary") 
					Xl.Sheets(Sheet.Name).activate
			
			PDD_Financial_Year := This.GetCellValue("Financial Year",0,1)
			
			if (PDD_Financial_Year != Current_Financial_Year)
			{
				Xl.Sheets("Payment Summary Amendment").Visible := True
				Xl.Sheets("Payment Summary Amendment").activate
				Xl.ActiveSheet.Unprotect
				This.Sleep(1000)
				This.SetCellValue("Pay Centre:",Client_Info.Person_Pay_Centre ,0,1)
				FormatTime, Time,, dd/MM/yyyy
				This.SetCellValue("Section 3: Requested By (Payroll)","Bryn Abbott",6,1)
				This.SetCellValue("Section 3: Requested By (Payroll)",Time,6,4)
				
				PDD_Financial_Year := Xl.Range("D4").Value
				PDD_Financial_Year := SubStr(PDD_Financial_Year,1,4)
				
				winactivate, Mochasoft
				sendinput, {F4}
				clipboard := 
				This.Sleep(1000)
				
				winactivate, Mochasoft
				sendinput, {Alt}ES
				sendinput, ^c
				
				ClipWait, 2
				if ErrorLevel
					MsgBox, The attempt to copy text onto the clipboard failed.
				This.Sleep(1000)
				Array := StrSplit(Clipboard, "`n")
				
				if (substr(Array[3],33,16) = "M A I N  M E N U" )
				{
					This.Sleep(1000)
					winactivate, Mochasoft
					This.Sleep(500)
					sendinput, pay
					sendinput,{enter}
					This.Sleep(500)
					sendinput,{TAB 3}
					sendinput, %PDD_Financial_Year%26{F8}
					This.Sleep(500)
					sendinput,s{ENTER}
					clipboard := 
					ClipWait, 2
					
					This.Sleep(500)
					sendinput, {Alt}ES
					sendinput, ^c
					ClipWait, 2
					if ErrorLevel
						MsgBox, The attempt to copy text onto the clipboard failed.
					
					Array := StrSplit(Clipboard, "`n")
					GrossSumTotal := substr(Array[18],14,15)
				}
				This.Sleep(1000)
				This.SetCellValue("Reduce gross total:",GrossSumTotal,0,3)
				This.Sleep(500)
				Xl.ActiveWorkbook.save()
			}
			
			For sheet in Xl.Worksheets
				if InStr(Sheet.Name, "Year Summary") 
					Xl.Sheets(Sheet.Name).activate
			
			Xl.ActiveSheet.Unprotect
			
			Cell := Xl.Range("A:J").Find("Total Overpayment").address
			Cell_No := SubStr(Cell, 4) + 1
			Cell_Letter := SubStr(Cell, 1,3)
			Non_Zero_Pays := ""
			
			loop, 27
			{
				Cell_No := Cell_No + 1
				Cell := Cell_Letter . Cell_No
				if (Xl.ActiveSheet.Range(Cell).Value = "0") 
					Xl.Sheets("Pay " Cell_No - 6).Select(False)
				
				if (Xl.ActiveSheet.Range(Cell).Value != "0") {
					Non_Zero_Pays .= "Pay " Cell_No - 6 "`n"
				}
			}
			
			Xl.ActiveWindow.SelectedSheets.Visible := False 
			Xl.Sheets("Previous Financial Year Summary").Visible := True
			Xl.Sheets("Previous Financial Year Summary").activate
			Xl.Sheets("Current Financial Year Summary").Visible := True
			Xl.Sheets("Current Financial Year Summary").activate
			PDD_Financial_Year := This.GetCellValue("Financial Year",0,1)
			Xl.ActiveSheet.Range("C1").Value := "Name:"
			Xl.ActiveSheet.Range("C2").Value := "AGS:"
			This.SetCellValue("Name:",TRIM(Client_Info.Person_Last_Name) . ", " . TRIM(Client_Info.Person_Given_Name),0,1)
			This.SetCellValue("AGS:",Client_Info.ServiceNo,0,1)
			
			Document_Title := Client_Info.ServiceNo "_Paid Due Diff_" . StrReplace(PDD_Financial_Year, "/", "-")
			Document_Location := Xl.activeworkbook.path . "\" . Document_Title . ".pdf"
			
			Loop, parse, Non_Zero_Pays, `n,
				Xl.Sheets(A_LoopField).Select(False)
			
			xl.ActiveSheet.ExportAsFixedFormat(0, Document_Location ,0,True,False,1,100,False)
			
			if (Financial_Year := This.GetCellValue("Current", 0, 1) = Current_Financial_Year) {
				Financial_Year := RegExReplace(This.GetCellValue("Current", 0, 1), "/", "_")
				Client_Info.Tax_Amount := This.GetCellValue("TOTALS", 0, 1)
				Client_Info.Nett_Amount := This.GetCellValue("TOTALS", 0, 2)
				Client_Info.SalarySac_Amount := This.GetCellValue("TOTALS", 0, 3)
				Client_Info.Deductions_Amount := This.GetCellValue("TOTALS", 0, 4)
				Client_Info.Super_Amount := This.GetCellValue("TOTALS", 0, 6)
				
				WinActivate, Excel
				Xl.DisplayAlerts := (False)
				Xl.ActiveSheet.Unprotect
				Xl.Workbook.Close(0) 
				Xl.quit()
			}
			
			if (Financial_Year := This.GetCellValue("Previous", 0, 1) != Current_Financial_Year)
			{
				Past_Loaction := Xl.activeworkbook.path . "\"
				
				if (Client_Info.Gross_Amount != "")
				{
					Client_Info.Gross_Amount := Client_Info.Gross_Amount + This.GetCellValue("TOTALS", 0, 2)
					Client_Info.SalarySac_Amount_Gross := Client_Info.SalarySac_Amount_Gross + This.GetCellValue("TOTALS", 0, 3)
					Client_Info.Previous_Fin_Year_Super := Client_Info.Previous_Fin_Year_Super + This.GetCellValue("TOTALS", 0, 5)
				}
				
				else if (Client_Info.Gross_Amount = "")
				{
					Client_Info.Gross_Amount := This.GetCellValue("TOTALS", 0, 2)
					Client_Info.SalarySac_Amount_Gross := This.GetCellValue("TOTALS", 0, 3)
					Client_Info.Previous_Fin_Year_Super := This.GetCellValue("TOTALS", 0, 5)
				}
				
				Xl := ComObjActive("Excel.Application")
				ComObjError(false)
				WinActivate, Excel
				Xl.DisplayAlerts := (False)
				Xl.ActiveSheet.Unprotect
				Xl.Workbook.Close(0) 
				Xl.quit()
				
			}
			
			
			
		}
		
		Process_Workbook(Client_Info)
		{
			NameArray := StrSplit(TRIM(Client_Info.Person_Given_Name), A_Space)
			Loop % NameArray.Length() 
				NameInit .= SubStr(NameArray[A_index],1,1) . ". "
			Client_Info.Initialise_Given_Names := TRIM(This.Change_String(Client_Info.Person_Courtesy_Title, "Title")) " " NameInit . This.Change_String(Client_Info.Person_Last_Name, "Title")
			
			
			Inputbox, Pay_Cycle, , What Pay Cycle will this be for?,,230, 125
			IniRead, Pay_Date, % "C:\Users\babb\Desktop\Workbook Class Settings.ini", Paydates, % Pay_Cycle
			
			Client_Info.Pay_Cycle := Pay_Cycle
			Client_Info.Pay_Date := Pay_Date
			
			
			Xl := ComObjActive("Excel.Application")
			ComObjError(false)
			Xl.ActiveSheet.Unprotect
			Xl.Sheets("OP Investigation Report").activate	
			Xl.ActiveSheet.Unprotect
			
			WinActivate, Mochasoft - mainframe.nt.gov.au
			This.Sleep(150)
			winactivate, Mochasoft
			This.Sleep(150)
			sendinput, {F4}Pay{enter}{F8}
			This.Sleep(150)
			;run, Calc
			InputBox, RecoveryRate , RCY.COM Percentage, Please Enter Ten Percent of the Gross Salary., , 200, 150
			
			if ErrorLevel
				MsgBox, CANCEL was pressed.
			WinClose, Calc
			
			Client_Info.RecoveryRate := RecoveryRate 
			This.SetCellValue("Pay Centre",Client_Info.Person_Pay_Centre,0,1)
			This.SetCellValue("10% of gross salary", Client_Info.RecoveryRate, 0, 1)
			This.SetCellValue("Date recovery will commence",Client_Info.Pay_Date, 0, 1)
			This.SetCellValue("Number and Street Address",TRIM(Client_Info.Address_Postal_Line_1), 0, 1)
			This.SetCellValue("Number and Street Address",TRIM(Client_Info.Address_Postal_Suburb) " " TRIM(Client_Info.Address_Postal_State) " " TRIM(Client_Info.Address_Postal_PostCode), 0, 3)
			
			This.SetCellValue("Name (Title, Initial, Surname)",TRIM(Client_Info.Initialise_Given_Names),0,1)
			This.SetCellValue("AGS",Client_Info.ServiceNo,0,1)
			This.SetCellValue("TAX",Client_Info.Tax_Amount,1,0)
			This.SetCellValue("Nett Diff",Client_Info.Nett_Amount,1,0)
			This.SetCellValue("Net Pay Diff",Client_Info.Nett_Amount,1,0) 
			This.SetCellValue("Gross Diff",Client_Info.Gross_Amount,1,0) 
			This.SetCellValue("Gross Diff", Client_Info.SalarySac_Amount_Gross,1,3) 
			This.SetCellValue("Salary Sacrifice Contribution",Client_Info.SalarySac_Amount,1,0)
			This.SetCellValue("Fund Code",(Client_Info.Super_Amount ? Client_Info.Super_Amount : 0) + (Client_Info.Previous_Fin_Year_Super ? Client_Info.Previous_Fin_Year_Super : 0),1,1)
			This.SetCellValue("Deductions",Client_Info.Deductions_Amount, 1, 0)
			
			if (Client_Info.Gross_Amount != "")
				This.SetCellValue("Gross",Client_Info.Gross_Amount,1,0)
			
			This.SetCellValue("Fund Name", Client_Info.Super_Fund_Name, 0, 1)
			This.SetCellValue("Fund Address", Client_Info.Super_Fund_Street, 0, 1)
			This.SetCellValue("Fund Name", Client_Info.Super_Fund_State, 2, 1)
			This.SetCellValue("Fund Code", Client_Info.Super_Fund_ID, 0, 1)
			
			;This.SetCellValue("Fund Name", "Previous Financial Year Super:", 1, 6)
			;This.SetCellValue("Fund Name", Client_Info.Previous_Fin_Year_Super, 2, 6)
			
			This.SetCellValue("Member Number",Client_Info.Super_Member_No , 0, 1)
			This.SetCellValue("Member DOB",Client_Info.Person_Date_Of_Birth, 0, 1)
			This.SetCellValue("FBT Date", Client_Info.FBT_Date, 0, 1)
			
			This.SetCellValue("Cost Code",Client_Info.Person_Cost_Code,0,1)
			
			Xl.Sheets("Recovery Authorisation").activate
			This.SetCellValue("Cost Code", Client_Info.Person_Cost_Code,0,1)
			Xl.Sheets("PFES Recovery Authorisation").activate
			This.SetCellValue("Cost Code", Client_Info.Person_Cost_Code,0,1)
			Xl.Sheets("OP Investigation Report").activate
			
			if (This.GetCellValue("Is Employee on Leave Without Pay?", 0, 1) = "Yes")
			{
				Xl.Sheets("OP Letter").Select(True)
				Xl.Sheets("PAWA OP Letter").Select(False)
				Xl.Sheets("Recovery Authorisation").Select(False)
				Xl.Sheets("PFES OP Letter").Select(False)
				Xl.Sheets("PAWA Recovery Authorisation").Select(False)
				Xl.Sheets("PFES Recovery Authorisation").Select(False)
				Xl.ActiveWindow.SelectedSheets.Visible := False 
				
				Xl.Sheets("OP Investigation Report").Visible := True
				Xl.Sheets("OP Investigation Report").activate
				Xl.Sheets("OP Investigation Report").Visible := True
			}
			if (This.GetCellValue("Is Employee on Leave Without Pay?", 0, 1) = "No")
			{
				Xl.Sheets("LWOP OP Letter").Select(True)
				Xl.Sheets("PFES LWOP OP Letter").Select(False)
				
				Xl.ActiveWindow.SelectedSheets.Visible := False 
				Xl.Sheets("OP Investigation Report").Visible := True
				Xl.Sheets("OP Investigation Report").activate
				Xl.Sheets("OP Investigation Report").Visible := True
			}
			
			if (This.GetCellValue("Overpayment Investigation Report", -1, 0) = "Previous Financial Year")
			{
				This.SetCellValue("Fund Name", Client_Info.Super_Fund_Name . " - N/A - Previous Financial Year", 0, 1)	
				
				Xl.Sheets("Checklist").activate
				
				This.SetCellValue("Overpayment Document placed in Future Action", "Pay Cycle: " Client_Info.Pay_Cycle, 0, 1)
				This.Merge_and_Center("C19", "D19", "N/A - Previous Financial Year")
				This.Merge_and_Center("C33", "D33", "N/A - Previous Financial Year")
				Xl.Sheets("OP Investigation Report").activate
			}
			
			if (This.GetCellValue("Overpayment Investigation Report", -1, 0) = "Current Financial Year")
			{
				Xl.Sheets("Checklist").activate
				This.SetCellValue("Overpayment Document placed in Future Action", "Pay Cycle: " Client_Info.Pay_Cycle, 0, 1)
				
				This.Merge_and_Center("C26", "D26", "N/A - Current Financial Year")
				This.Merge_and_Center("C18", "D18", "N/A - Current Financial Year")
				Xl.Sheets("OP Investigation Report").activate
			}
			
			if ((SubStr(This.GetCellValue("Pay Centre", 0, 1), 1,3), 0, 1) != "673")
			{
				Xl.Sheets("PAWA OP Letter").Select(True)
				Xl.Sheets("PAWA Recovery Authorisation").Select(False)
				Xl.Sheets("PFES OP Letter").Select(False)
				Xl.Sheets("PFES LWOP OP Letter").Select(False)
				Xl.Sheets("PAWA Recovery Authorisation").Select(False)
				Xl.Sheets("PFES Recovery Authorisation").Select(False)
				Xl.ActiveWindow.SelectedSheets.Visible := False 
				Xl.Sheets("OP Investigation Report").Visible := True
				Xl.Sheets("OP Investigation Report").activate
				Xl.Sheets("OP Investigation Report").Visible := True
			}	
			
			if ((SubStr(This.GetCellValue("Pay Centre", 0, 1), 1,3), 0, 1) = "673")
			{
				Xl.Sheets("PAWA OP Letter").Select(True)
				Xl.Sheets("PAWA Recovery Authorisation").Select(False)
				Xl.Sheets("OP Letter").Select(False)
				Xl.Sheets("Recovery Authorisation").Select(False)
				Xl.Sheets("PAWA Recovery Authorisation").Select(False)
				Xl.ActiveWindow.SelectedSheets.Visible := False 
				Xl.Sheets("OP Investigation Report").Visible := True
				Xl.Sheets("OP Investigation Report").activate
				Xl.Sheets("OP Investigation Report").Visible := True
			}	
			
	;Overpayment Investigation Report
			
			;Client_Info.PayCentre := This.GetCellValue("Pay Centre", 0, 1)
			Client_Info.Date_Detected := This.GetCellValue("Date Overpayment Detected", 0, 1)
			Client_Info.Date_Detected := StrReplace(Client_Info.Date_Detected, ".","/")
			This.SetCellValue("Date Overpayment Detected", Client_Info.Date_Detected, 0, 1)
			
			Client_Info.Date_Commenced := This.GetCellValue("Date O/P commenced", 1, 0)
			Client_Info.Date_Commenced := StrReplace(Client_Info.Date_Commenced, ".","/")
			This.SetCellValue("Date O/P commenced", Client_Info.Date_Commenced, 1, 0)
			
			Client_Info.Date_Ceased := This.GetCellValue("Date O/P ceased", 1, 0)
			Client_Info.Date_Ceased := StrReplace(Client_Info.Date_Ceased, ".","/")
			This.SetCellValue("Date O/P ceased", Client_Info.Date_Ceased, 1, 0)
			
			Client_Info.TotalAmount := This.GetCellValue("Total overpayment to be recovered", 1, 0)
			Client_Info.TotalAmount := RegExReplace(Client_Info.TotalAmount, "[^0-9.]")
			
			
			if Client_Info.PayCentre != "673"
				Client_Info.Reason := This.GetCellValue("Reason for Overpayment (This description will be shown on the letter to the employee)", 1, 0) " " . Client_Info.Date_Commenced . " to " . Client_Info.Date_Ceased . ". ($" . Client_Info.TotalAmount . ")"  	
			
	; ePOD Reporting
			Client_Info.Error_Source := This.GetCellValue("Error Source", 1, 0)
			Client_Info.Error_Type := This.GetCellValue("Error Type", 1, 0)
			Client_Info.Error_Cause := This.GetCellValue("Error Cause", 1, 0)
			Client_Info.Location := This.GetCellValue("Location", 1, 0)
			
	;Cumulative Totals and Recover Overpaid Tax
			Client_Info.RPMENT := This.Float(This.GetCellValue("RPM.ENT", 0, 2)) 
			Client_Info.TAXPYM := This.Float(This.GetCellValue("TAX.PYM", 0, 2)) 
			Client_Info.RDRENT := This.Float(This.GetCellValue("RDR.ENT", 0, 2)) 
			Client_Info.RCYCOM := This.Float(This.GetCellValue("RCY.COM", -1, 3))
			Client_Info.Percentage := This.Float(This.GetCellValue("10% of gross salary", 0, 1),2) 
			
	;Finalise and Fix Up stuff in the workbook
			Xl.Sheets("Checklist").activate
			This.SetCellValue("Overpayment Document placed in Future Action", "Pay Cycle: " Client_Info.Pay_Cycle, 0, 1)
			FormatTime, Time,, dd/MM/yyyy
			This.SetCellValue("Overpayment workbook (Blue Sections)", "Bryn Abbott | " time, 0, 1)
			
			Array := StrSplit(Client_Info.Date_Detected, "/")
			DetectedDate := Array.3 . Array.2 . Array.1
			FormatTime, TimeString,, shortDate
			ProcessedDateString := TimeString
			Array := StrSplit(ProcessedDateString, "/")
			ProcessedDate := Array.3 . Array.2 . Array.1
			datefrom := DetectedDate
			distance := dateto := ProcessedDate
			distance -= datefrom, days
			Client_Info.Distance := distance
			
			This.SetCellValue("Overpayment workbook (Blue Sections)", "Days Since Discovery:", 0, 4)
			This.SetCellValue("Overpayment workbook (Blue Sections)", Client_Info.Distance, 1, 4)
			
			Xl.Sheets("Checklist").activate
			Xl.ActiveSheet.Unprotect
			
			
			Pointer := StrReplace(Xl.ActiveSheet.Range("A:H").Find("Overpayment PTR's (RPM.ENT, TAX.PYM, RDR.ENT)").Offset(1, 0).address,"$")
			Letter := SubStr(Pointer,1,1)
			Number := SubStr(Pointer,2)
			
			
			
			
			Loop 24
				if (Letter = Chr(96+A_Index))
					OtherLetter := Chr(96+A_Index+3)
			
			Xl.ActiveSheet.Range(Letter Number ":" OtherLetter Number).select
			
			xlShiftDown := -4121
			xlShiftUp := -4162
			xlFormatFromLeftOrAbove := 0
			Xl.Selection.Insert(Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove)
			
			Xl.Selection.HorizontalAlignment := -4131
			Xl.Selection.VerticalAlignment := -4107
			Xl.Selection.WrapText := False
			Xl.Selection.Orientation := 0
			Xl.Selection.AddIndent := False
			Xl.Selection.IndentLevel := 0
			Xl.Selection.ShrinkToFit := False
			Xl.Selection.ReadingOrder := -5002
			Xl.Selection.MergeCells := False
			
			Xl.ActiveSheet.Range(Letter Number ":" "b" Number).select
			Xl.Selection.HorizontalAlignment := -4131
			Xl.Selection.VerticalAlignment := -4107
			Xl.Selection.HorizontalAlignment := -4131
			Xl.Selection.VerticalAlignment := -4107
			Xl.Selection.WrapText := False
			Xl.Selection.Orientation := 0
			Xl.Selection.AddIndent := False
			Xl.Selection.IndentLevel := 0
			Xl.Selection.ShrinkToFit := False
			Xl.Selection.ReadingOrder := -5002
			Xl.Selection.MergeCells := False
			
			Xl.Selection.Merge
			
			
			
			
			
			excelforula = =CONCATENATE("Super Recovery Actioned (",('OP Investigation Report'!B55),")")
			Xl.ActiveSheet.Range(Letter Number).value := (excelforula)
			
			
			Xl.Sheets("OP Investigation Report").activate
			
			Xl := ComObjActive("Excel.Application")
			ComObjError(false)
			
			Xl.ActiveSheet.Unprotect
			Pointer := StrReplace(Xl.ActiveSheet.Range("A:H").Find("Superannuation to be recovered from fund").Offset(5, 0).address,"$")
			Letter := SubStr(Pointer,1,1)
			Number := SubStr(Pointer,2)
			
			Loop 24
				if (Letter = Chr(96+A_Index))
					OtherLetter := Chr(96+A_Index+3)
			
			Xl.ActiveSheet.Range(Letter Number ":" Letter Number+1).select
			Xl.Selection.UnMerge
			
			Xl.ActiveSheet.Range(Letter Number).value := ("Actions Taken for Super Recovery")
			Selected_Cells := Letter Number . ":" . OtherLetter Number
			
			Xl.ActiveSheet.Range(Selected_Cells).select
			Xl.Selection.HorizontalAlignment:=-4108
			Xl.Selection.VerticalAlignment:=-4107
			Xl.Selection.WrapText := (False)
			Xl.Selection.Orientation := (0)
			Xl.Selection.AddIndent := (False)
			Xl.Selection.ShrinkToFit := (False)
			Xl.Selection.MergeCells := (False)
			Xl.Selection.Font.Bold := (True)
			Xl.Selection.Interior.Pattern := 1
			Xl.Selection.Interior.PatternColorIndex := -4105
			Xl.Selection.Interior.ThemeColor := 10.5
			
			Xl.Selection.Interior.TintAndShade := 0-0.249977111117893
			Xl.Selection.Interior.PatternTintAndShade := 0
			Xl.Selection.Merge
			
			Xl.Selection := 
			
			Xl.ActiveSheet.Range(Letter Number+1 ).value := ("Select Action Taken")
			Selected_Cells := Letter Number+1 . ":" . OtherLetter Number+1
			
			Xl.ActiveSheet.Range(Selected_Cells).select
			Xl.Selection.HorizontalAlignment:=-4108
			Xl.Selection.VerticalAlignment:=-4107
			Xl.Selection.WrapText := (False)
			Xl.Selection.Orientation := (0)
			Xl.Selection.AddIndent := (False)
			Xl.Selection.ShrinkToFit := (False)
			Xl.Selection.MergeCells := (False)
			Xl.Selection.Font.Bold := (True)
			
			Xl.Selection.Merge
			
			Xl.Rows(Number+1 ":" Number+1).EntireRow.AutoFit
			
			xlEdgeLeft = 7 ; http://techsupt.winbatch.com/ts/T000001033005F9.html
			xlEdgeTop = 8
			xlEdgeBottom = 9
			xlEdgeRight = 10
			xlContinuous = 1
			xlThick = 4
			xlAutomatic = -4105
			xlThin = 2
			
			Xl.Range(Letter Number ":" OtherLetter Number+1).Select
			Xl.Selection.Borders(xlDiagonalDown).LineStyle := xlNone
			Xl.Selection.Borders(xlDiagonalUp).LineStyle := xlNone
			Xl.Selection.Borders.LineStyle := xlContinuous 
			Xl.Selection.Borders.ColorIndex := 0
			Xl.Selection.Borders.TintAndShade := 0
			Xl.Selection.Borders.Weight := xlThin 
			
			Xl.ActiveSheet.Range(Letter Number+1).select
			xl.Selection.Validation.Add[Type:=3,AlertStyle:=1,Operator:=3, Formula1:="Created SGR.COM, Added to Total Overpayment,Applied to Super Fund for Refund, Mixture of Processes"]
			
			Xl.Sheets("OP Investigation Report").activate
			Xl.ActiveSheet.Unprotect
			
			Pointer := StrReplace(Xl.ActiveSheet.Range("A:H").Find("Date recovery will commence").Offset(0, 1).address,"$")
			Letter := SubStr(Pointer,1,1)
			Number := SubStr(Pointer,2)
			
			Xl.ActiveSheet.Range(Letter Number).select
			
			
			
			;xl.Selection.Validation.Add[Type:=3,AlertStyle:=1,Operator:=3, Formula1:="08/07/2020,22/07/2020,06/08/2020,20/08/2020,03/09/2020,17/09/2020,01/10/2020,15/10/2020,29/10/2020,12/11/2020,26/11/2020,10/12/2020,24/12/2020,07/01/2021,21/01/2021,04/02/2021,18/02/2021,04/03/2021,18/03/2021,01/04/2021,15/04/2021,29/04/2021,13/05/2021,27/05/2021,10/06/2021,24/06/2021"]
			
			
			
			Xl.Range("C19:F19").CheckSpelling 
			RefreshListView()
		}
		
		Merge_and_Center(Cell_1, Cell_2, String) 
		{
			Xl := ComObjActive("Excel.Application")
			Xl.Sheets("Checklist").activate	
			Xl.ActiveSheet.Unprotect
			
			if (String != " ")
				Xl.ActiveSheet.Range(Cell_1).value := (String)
			
			Selected_Cells := Cell_1 . ":" . Cell_2
			Xl.ActiveSheet.Range(Selected_Cells).select
			Xl.Selection.HorizontalAlignment:=-4108
			Xl.Selection.VerticalAlignment:=-4107
			Xl.Selection.WrapText := (False)
			Xl.Selection.Orientation := (0)
			Xl.Selection.AddIndent := (False)
			Xl.Selection.ShrinkToFit := (False)
			Xl.Selection.ReadingOrder.XlContext
			Xl.Selection.MergeCells := (False)
			Xl.Selection.Font.Bold := (True)
			Xl.Selection.Merge
			return
		}

		RefreshFromWorkbook(){
			msgbox
			Client_Info.Address_Postal_Line_1 := This.GetCellValue("Number and Street Address",0,1)
			Client_Info.Address_Postal_Suburb := This.GetCellValue("Suburb, State, Post Code",0,1)
			Client_Info.ServiceNo := This.Float(This.GetCellValue("AGS",0,1),8)
			Client_Info.Tax_Amount := This.GetCellValue("Tax Diff",0,1)
			Client_Info.Nett_Amount := This.GetCellValue("Net Pay Diff",0,1)
			Client_Info.Gross_Amount := This.GetCellValue("Gross Diff",0,1)
			Client_Info.SalarySac_Amount_Gross := This.GetCellValue("Salary Sacrifice Contribution",0,1)
			Client_Info.SalarySac_Amount := This.GetCellValue("Salary Sacrifice Contribution",0,1)
			Client_Info.Super_Amount := This.GetCellValue("Amount to be recovered",0,1)
			Client_Info.Deductions_Amount := This.GetCellValue("Deductions Diff",0,1)
			Client_Info.Gross_Amount := This.GetCellValue("Gross Diff",0,1)
			Client_Info.Super_Fund_Name := This.GetCellValue("SuperWrap Superannuation",0,1)
			Client_Info.Super_Fund_ID := This.GetCellValue("Fund Code",0,1)
			Client_Info.Super_Member_No := This.GetCellValue("Member Number",0,1)
			Client_Info.FBT_Date := This.GetCellValue("FBT Date",0,1)
			Client_Info.PayCentre := This.Float(This.GetCellValue("Pay Centre", 0, 1),3)
			Client_Info.Date_Detected := This.GetCellValue("Date Overpayment Detected", 0, 1)
			Client_Info.Date_Detected := StrReplace(Client_Info.Date_Detected, ".","/")
			Client_Info.Date_Commenced := This.GetCellValue("Date O/P commenced", 1, 0)
			Client_Info.Date_Commenced := StrReplace(Client_Info.Date_Commenced, ".","/")
			Client_Info.Date_Ceased := This.GetCellValue("Date O/P ceased", 1, 0)
			Client_Info.Date_Ceased := StrReplace(Client_Info.Date_Ceased, ".","/")
			Client_Info.TotalAmount := This.GetCellValue("Total overpayment to be recovered", 1, 0)
			Client_Info.TotalAmount := RegExReplace(Client_Info.TotalAmount, "[^0-9.]")
			
			
			if Client_Info.PayCentre != "673"
				Client_Info.Reason := This.GetCellValue("Reason for Overpayment (This description will be shown on the letter to the employee)", 1, 0) " " . Client_Info.Date_Commenced . " to " . Client_Info.Date_Ceased . ". ($" . Client_Info.TotalAmount . ")"  	
			
			; ePOD Reporting
			Client_Info.Error_Source := This.GetCellValue("Error Source", 1, 0)
			Client_Info.Error_Type := This.GetCellValue("Error Type", 1, 0)
			Client_Info.Error_Cause := This.GetCellValue("Error Cause", 1, 0)
			Client_Info.Location := This.GetCellValue("Location", 1, 0)
			
			;Cumulative Totals and Recover Overpaid Tax
			Client_Info.RPMENT := This.Float(This.GetCellValue("RPM.ENT", 0, 2)) 
			Client_Info.TAXPYM := This.Float(This.GetCellValue("TAX.PYM", 0, 2)) 
			Client_Info.RDRENT := This.Float(This.GetCellValue("RDR.ENT", 0, 2)) 
			Client_Info.RCYCOM := This.Float(This.GetCellValue("RCY.COM", -1, 3))
			Client_Info.Percentage := This.Float(This.GetCellValue("10% of gross salary", 0, 1),2) 

}

		Activate()  
		{ 
			WinActivate % "ahk_id " This.ID  
			Return This.ID
		}   
		
		GetWindowState()
		{
			This.Activate() 
			if FindText(1502-150000, 121-150000, 1502+150000, 121+150000, 0, 0, This.CheckListArray["MAINMENUCHECK"])
				Return This.WindowState:="MAINMENUCHECK"
			else if FindText(1538-150000, 154-150000, 1538+150000, 154+150000, 0, 0, This.CheckListArray["TIMEOUTCHECK"]) 
				Return This.WindowState:="TIMEOUTCHECK"
			else if FindText(2492-150000, 715-150000, 2492+150000, 715+150000, 0, 0, This.CheckListArray["SEPARATEDCHECK"])
				Return This.WindowState:="SEPARATEDCHECK"
			else if FindText(1510-150000, 184-150000, 1510+150000, 184+150000, 0, 0, This.CheckListArray["CASUALCHECK"])
				Return This.WindowState:="CASUALCHECK"
			else if FindText(1750-150000, 550-150000, 1750+150000, 550+150000, 0, 0, This.CheckListArray["TEMPCHECK"])
				Return This.WindowState:="TEMPCHECK"
			else if FindText(2493-150000, 153-150000, 2493+150000, 153+150000, 0, 0, This.CheckListArray["COFCHECK"])
				Return This.WindowState:="COFCHECK"
			else if FindText(1494-150000, 87-150000, 1494+150000, 87+150000, 0, 0, This.CheckListArray["SPR"])
				Return This.WindowState:="SPR"
			else if FindText(2493-150000, 153-150000, 2493+150000, 153+150000, 0, 0, This.CheckListArray["NTGPASSCHECK"])
				Return This.WindowState:="NTGPASSCHECK"
			else
				Return This.WindowState:="ERROR"
		}
		
		Refresh_from_Workbook() {
			
		;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		;~~~~~~~~ Need to do this sometime soon! ~~~~
		;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
		}
		
		GetCellValue(Find,offset_X,offset_Y)
		{
			Temp := ComObjActive("Excel.Application")
			ComObjError(false)
			Pointer := Temp.ActiveSheet.Range("A:H").Find(Find)
			return Temp.Range(Pointer.Offset(offset_X, offset_Y).address).Value
		}
		
		SetCellValue(Find,Value,offset_X,offset_Y)
		{
			Temp := ComObjActive("Excel.Application")
			ComObjError(false)
			Pointer := Temp.ActiveSheet.Range("A:H").Find(Find)
			Temp.ActiveSheet.Range(Pointer.Offset(offset_X, offset_Y).address).Value := Value
		}
		
		MainframeToTextArray()
		{
			This.Sleep(100)	
			winactivate, Mochasoft
			This.Sleep(100)
			sendinput, {Alt}ES
			This.Sleep(500)
			sendinput, ^c
			This.Sleep(500)
			This.Array := StrSplit(Clipboard, "`n")
			This.Sleep(200)
			Return This.Array
		}
		
		Sleep(Duration)
		{
			if (Thread_Kill_Token = True)
			{
				Thread_Kill_Token := !Thread_Kill_Token
				Exit
			}
			Sleep % Duration
		}
		
		__Delete()
		{
		;MsgBox % "Delete Actions."
		}
	}
	
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~ GUI CLASS ~~~~~~~~~~~~~~~~
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	Class Gui_Class
	{
		__New()
		{
		;Global
			Static
			
		; Building File, Edit and Tools Menus
			Menu, Tray, Icon, % "C:\_Umbra Sector\_Resources\QuickCodeIcon.ico"
			
			Menu FileMenu, Add, Refresh List View, ContextMenuHandler
			Menu FileMenu, Add, Hide Descktop Icons, ContextMenuHandler
			Menu FileMenu, Add, Refresh from Workbook, ContextMenuHandler
			Menu FileMenu, Add
			Menu FileMenu, Add, Close Toolbar, ContextMenuHandler
			Menu MenuBar, Add, File, :FileMenu
			Menu Edit, Add, Remove Symbols, ContextMenuHandler
			Menu Edit, Add, Remove Formating, ContextMenuHandler
			Menu Edit, Add, Uppercase Clipboard, ContextMenuHandler
			Menu Edit, Add, TitleCase Clipboard, ContextMenuHandler
			Menu Edit, Add, Lowercase Clipboard, ContextMenuHandler
			Menu Edit, Add, Make Selection Negative Number, ContextMenuHandler
			Menu Edit, Add, StrReplace, ContextMenuHandler
			Menu EditMenu, Add, Text Manipulation, :Edit
			Menu EditMenu2, Add, To Be Checked, ContextMenuHandler
			Menu EditMenu2, Add, Checked, ContextMenuHandler
			Menu EditMenu2, Add, Invoice Raised, ContextMenuHandler
			Menu EditMenu2, Add, Trimmed, ContextMenuHandler
			Menu EditMenu2, Add, Sent to Client, ContextMenuHandler
			Menu EditMenu, Add, Notes, :EditMenu2
			Menu MenuBar, Add, Edit, :EditMenu
			Menu ToolsMenu, Add, Build OVP Email, ContextMenuHandler
			Menu ToolsMenu, Add, Quick Code Tester, ContextMenuHandler
			Menu ToolsMenu, Add, Reset Work Area, ContextMenuHandler
			Menu ToolsMenu, Add, Edit This Script, ContextMenuHandler
			Menu ToolsMenu, Add, Close IE Windows, ContextMenuHandler
			Menu ToolsMenu, Add, AutoGUI, ContextMenuHandler
			Menu ToolsMenu, Add, AHK Studio, ContextMenuHandler
			Menu ToolsMenu, Add, eReader, ContextMenuHandler
			Menu ToolsMenu, Add, Create Manual Payee Sheet, ContextMenuHandler
			
			
			Menu ContextMenu, Add, Chrome, ContextMenuHandler
	;Menu ContextMenu, Icon, C:\Users\babb\AppData\Local\Google\Chrome\Application\chrome.exe,, 1
	;Menu, ContextMenu, Icon, Chrome, C:\Users\babb\AppData\Local\Google\Chrome\Application\chrome.exe,,32
			Menu ContextMenu, Add, GAS Notes, ContextMenuHandler
	;Menu, ContextMenu, Icon, GAS Notes, C:\_Umbra Sector\Icons\1485347325_EditDocument.ico,,32
			Menu ContextMenu, Add, Quick Code, ContextMenuHandler
	;Menu ContextMenu, Icon, Quick Code,C:\_Umbra Sector\Icons\QuickCodeIcon.ico,,32
			Menu ContextMenu, Add, Webnovel Reader, ContextMenuHandler
	;Menu ContextMenu, Icon, Webnovel Reader,C:\_Umbra Sector\Icons\reaad_kzP_icon.ico,,32
			Menu ContextMenu, Add
			Menu Notes, Add, To Be Checked, ContextMenuHandler
			Menu Notes, ADD, Checked, ContextMenuHandler
			Menu Notes, ADD, Invoice Raised , ContextMenuHandler
			Menu Notes, ADD, Trimmed  , ContextMenuHandler
			Menu Notes, ADD, Sent to Client  , ContextMenuHandler
			Menu ContextMenu, Add, Notes, :Notes
			Menu ScriptsMenu, Add, BLANK	 AHK, ContextMenuHandler
			Menu ContextMenu, Add, Scripts, :ScriptsMenu
			Menu Navigate, Add, GAS Reports 	Explorer, ContextMenuHandler
			Menu Navigate, Add, Script Dir		Explorer, ContextMenuHandler
			Menu Navigate, Add, Templates	Explorer, ContextMenuHandler
			Menu ContextMenu, Add, Navigate, :Navigate
			Menu Programs, Add, AutoGUI, MenuHandler
			Menu ContextMenu, Add, Programs, :Programs
			Menu TextManipulation, Add, Remove Symbols, ContextMenuHandler
			Menu TextManipulation, Add, Remove Formating, ContextMenuHandler
			Menu TextManipulation, Add, Uppercase Clipboard, ContextMenuHandler
			Menu TextManipulation, Add, TitleCase Clipboard, ContextMenuHandler
			Menu TextManipulation, Add, Lowercase Clipboard, ContextMenuHandler
			Menu TextManipulation, Add, Make Selection Negative Number, ContextMenuHandler
			Menu TextManipulation, Add, StrReplace, ContextMenuHandler
			Menu ContextMenu, Add, Text Manipulation, :TextManipulation
			Menu ContextMenu, Add
			Menu ContextMenu, Add, Reload Context Menu, ContextMenuHandler
			
			Menu ContextMenuMedia, Add, Mini Player Mode, ContextMenuHandler
			
			Menu ContextMenuMedia, Add, Select File, ContextMenuHandler
			Menu ContextMenuMedia, Add, Close WMP Instance, ContextMenuHandler
			
			
			Gui, +LastFound -Resize -Caption -Border +HWNDhGui1 +ToolWindow  -SysMenu +AlwaysOnTop 
		;Gui, +LastFound  +HWNDhGui1 +ToolWindow  +AlwaysOnTop ; Movable
			
			Gui Color, FFFFFF
			Gui ,Font, s9 cFFFFFF, Segoe UI ; Set font options
			
		;Outlook Blue Menu BG
			Gui, Add, Text, % "x" 1 " y" 0 " w545"  vHELLOWORLD " h" 59 " +0x4E +HWNDhTitleHeader " 
			DllCall("SendMessage", "Ptr", hTitleHeader, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0173C7", 1, 1))
			
		;Blue Border
			Gui, Add, Text, % " x" 0 " y" 0 " w" 1 " h" 876 " +0x4E +HWNDhBorderLeft "
			DllCall("SendMessage", "Ptr", hBorderLeft, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0072C6", 1, 1))
			Gui, Add, Text, % "x" 1 " y" 870-1 " w" 199-2 " h" 1 " +0x4E +HWNDhBorderBottom"
			DllCall("SendMessage", "Ptr", hBorderBottom, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0072C6", 1, 1))
			
		; File Menu
			Gui , Add, Picture, % " x" 15 " y" 30 " w" 60 " h" 24 " +0x4E +HWNDhButtonMenuFileN Hidden0"
			Gui ,Add, Picture, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +0x4E +HWNDhButtonMenuFileH Hidden1"
			Gui ,Add, Text, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +HWNDhButtonMenuFileText +BackgroundTrans +0x201", % "File"
			DllCall("SendMessage", "Ptr", hButtonMenuFileN, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0173C7", 1, 1))
			DllCall("SendMessage", "Ptr", hButtonMenuFileH, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("2A8AD4", 1, 1))
			
		; Edit Menu
			Gui , Add, Picture, % " x+" 2 " yp" 0 " w" 60 " h" 24 " +0x4E +HWNDhButtonMenuEditN Hidden0"
			Gui , Add, Picture, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +0x4E +HWNDhButtonMenuEditH Hidden1"
			Gui , Add, Text, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +HWNDhButtonMenuEditText +BackgroundTrans +0x201", % "Edit"
			DllCall("SendMessage", "Ptr", hButtonMenuEditN, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0173C7", 1, 1))
			DllCall("SendMessage", "Ptr", hButtonMenuEditH, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("2A8AD4", 1, 1))
			
		; Tools Menu
			Gui , Add, Picture, % " x+" 2 " yp" 0 " w" 60 " h" 24 " +0x4E +HWNDhButtonMenuToolsN Hidden0"
			Gui , Add, Picture, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +0x4E +HWNDhButtonMenuToolsH Hidden1"
			Gui , Add, Text, % " xp" 0 " yp" 0 " wp" 0 " hp" 0 " +HWNDhButtonMenuToolsText +BackgroundTrans +0x201", % "Tools"
			DllCall("SendMessage", "Ptr", hButtonMenuToolsN, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("0173C7", 1, 1))
			DllCall("SendMessage", "Ptr", hButtonMenuToolsH, "UInt", 0x172, "Ptr", 0, "Ptr", This.CreateDIB("2A8AD4", 1, 1))
			
		; Title Text
			Gui Font, Bold  s10
			Gui , Add, Text, x0 y6 w215 h24 +BackgroundTrans +0x201, % "Payroll Debt Recovery"
			Gui Font,
			
			Gui Add, Edit, x11 y71 w155 h21 +Center vServiceNo,  
			Gui Add, Button, x+ y70  h23 gMenuHandler, Clear
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler, Load Client Info
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler , Prepare Paid Due Diff
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler , Process Workbook
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler, Create ePOD Record
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler , Adjust Gross Totals
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler , Enter Recovery Action
			
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler  , SG Recovery Action
			
			
			Gui Add, Button, x11 y+ w191 h23 gMenuHandler  , Generate Subject Line
			
			
			;Gui Add, GroupBox, x3 y267 w207 h274, Variables:
			Gui Add, ListView, x11 y283 w191 h250 -readonly gSubLV1 hwndHLV1 r10 hwndHLV1 AltSubmit -E0x200  vLV1 , Variable:|Contents:
			
			Gui Add, ActiveX, x11 y556 w191 h117 vWMP +hwndhWMPLayer +Hidden, WMPLayer.OCX
			Gui Add, GroupBox, x3 y541 w207 h140, Quick Notes
			
			Gui Add, Edit, x11 y556 w191 h117 -E0x200 +Multi vQuickNotes +HwndhQuickNotes
			
			
			
			Gui Add, GroupBox, x3 y682 w207 h112, Quick Access
			Gui Add, Button, x15 y703 w45 h23 gQuickAccessHandler, Recent
			Gui Add, Button, x61 y703 w45 h23 gQuickAccessHandler, ePOD
			Gui Add, Button, x107 y703 w45 h23 gQuickAccessHandler, PIPS
			Gui Add, Button, x153 y703 w45 h23 gQuickAccessHandler, F:\
			Gui Add, Button, x14 y731 w45 h23 gQuickAccessHandler, Chrome
			Gui Add, Button, x60 y731 w45 h23 gQuickAccessHandler, Phones
			Gui Add, Button, x106 y731 w45 h23 gQuickAccessHandler, myHR
			Gui Add, Button, x154 y731 w45 h23 gQuickAccessHandler, GovAcc
			Gui Add, Button, x14 y758 w45 h23 gQuickAccessHandler, VarDB
			Gui Add, Button, x60 y758 w45 h23 gQuickAccessHandler, Q-Mast
			Gui Add, Button, x106 y758 w45 h23 gQuickAccessHandler, Calc
			Gui Add, Button, x154 y758 w45 h23 gQuickAccessHandler, FileNote
			
			
			Gui, Show, x1225 y0 h870 w215, PDR Toolbar
			
			ICELV1 := New LV_InCellEdit(HLV1)
			ICELV1.OnMessage()
			OnMessage(0x200, ObjBindMethod(this,"WM_MOUSEMOVE"))
			OnMessage(0x202, ObjBindMethod(this,"WM_LBUTTONUP"))
			OnMessage(0x112, ObjBindMethod(this, "WM_SYSCOMMAND")) 
			
			VarSetCapacity(TME, 16, 0), NumPut(16, TME, 0), NumPut(2, TME, 4), NumPut(hGui1, TME, 8)
			This.Workarea(A_ScreenWidth-215, A_ScreenHeight-31)
			
			__Handles.hGui := hGui1
			__Handles.hWMPLayer := hWMPLayer
			__Handles.WMP := WMP
			__Handles.hQuickNotes := hQuickNotes
			__Handles.hButtonMenuFileN := hButtonMenuFileN
			__Handles.hButtonMenuFileH := hButtonMenuFileH
			__Handles.hButtonMenuFileText := hButtonMenuFileText
			__Handles.hButtonMenuEditN := hButtonMenuEditN
			__Handles.hButtonMenuEditH := hButtonMenuEditH
			__Handles.hButtonMenuEditText := hButtonMenuEditText
			__Handles.hButtonMenuToolsN := hButtonMenuToolsN
			__Handles.hButtonMenuToolsH := hButtonMenuToolsH
			__Handles.hButtonMenuToolsText := hButtonMenuToolsText
			
			This.PIPS_Toolbar()
		}
		
	;WMP() {
		;guicontrol,Hide, % __Handles.hQuickNotes
		;guicontrol,Show, % __Handles.hWMPLayer
		;WMP := __Handles.WMP
		;FileSelectFile, SelectedFile, 3, , Open a file
		;WMP.Url := SelectedFile 
		;WMP.uiMode := "None"
		;WMP.stretchToFit := 1  
	;}
		
		OpenRecent()
		{
			_Most_Recent := ""
			
			Loop, Files, % "C:\Users\" A_Username "\Documents\Offline Records (34)\*.*", R
			{
				FileGetTime, FileTime, % A_LoopFileLongPath
				If (FileTime > LatestFileTime) OR (A_Index = 1)
					LatestFileTime := FileTime, LatestFilePath := A_LoopFileLongPath
			}
			SplitPath, LatestFilePath,, dir
			Run % dir 
			
			Path := "C:\Users\" . A_UserName . "\Documents\Offline Records (34)"
			
			Loop Files, C:\Users\babb\Documents\Offline Records (34)\*.dat, R
			{
				FileMove, % A_LoopFileFullPath, % "C:\Users\babb\Documents\Offline Records (34)\TRIM .DAT FILE DUMP\" . A_TickCount . "_" . A_LoopFileShortName
				If (A_Index > 100)
					Break
				
			}
			
			Loop, C:\Users\babb\Documents\Offline Records (34)\*, 2, 1
				FL .= ((FL<>"") ? "`n" : "" ) A_LoopFileFullPath
			Sort, FL, R D`n ; Arrange folder-paths inside-out
			Loop, Parse, FL, `n
			{
				
				FileRemoveDir, %A_LoopField% ; Do not remove the folder unless is  empty
				If ! ErrorLevel
					Del := Del+1,  RFL .= ((RFL<>"") ? "`n" : "" ) A_LoopField
				
				If (A_Index > 100)
					Break
			}
			
		}
		
		PIPS_Toolbar() {
			Static
			Gui, 2:New, +LastFound -Caption +AlwaysOnTop +HWNDhPIPS_Bar
			
			gui, Margin, 0, 0
			Gui, Color, 0xF3F3F3
			Gui, Add, Button,  y3 h20 vpay gPIPS_Menu,  Payslip
			Gui, Add, Button,  y3 h20 vptrc gPIPS_Menu, PTR's
			Gui, Add, Button,  y3 h20 vESUP gPIPS_Menu, Super
			Gui, Add, Button,  y3 h20 vEADD gPIPS_Menu, Address 
			Gui, Add, Button,  y3 h20 vEPER gPIPS_Menu, Payee
			Gui, Add, Button,  y3  h20 vall gPIPS_Menu, Allowances
			Gui, Add, Button,  y3  h20 vLEV gPIPS_Menu, Leave
			Gui, font, Bold
			Gui, Add, Text, x+5 y6 h15, Load AGS No.: 
			Gui, font, 
			Gui, Add, ComboBox, x+5 y3 w150 h15 +hwndhComboBox vComboBox,
			Gui, Add, Button, x+5  y3  h20 Default vLoad gPIPS_Menu, Load
			
			WinGetClass, OutputVar , Mochasoft - mainframe.nt.gov.au
			
			This.SetParentByClass(OutputVar, 2)
			;This.SetParentByClass("WindowsForms10.Window.8.app.0.141b42a_r33_ad1", 2)
			
			
			Gui 2:Show, x0 y22, Pips toolbar
			
			__Handles.hPIPS_Bar := hPIPS_Bar
			__Handles.hComboBox := hComboBox
		}
		
		SetParentByClass(Window_Class, Gui_Number) { 
			Parent_Handle := DllCall( "FindWindowEx", "uint",0, "uint",0, "str", Window_Class, "uint",0) 
			Gui, %Gui_Number%: +LastFound 
			Return DllCall( "SetParent", "uint", WinExist(), "uint", Parent_Handle )
		}
		
		WM_SYSCOMMAND(wParam, lParam, msg, hwnd){
			
			static SC_CLOSE := 0xF060
			if (wParam = SC_CLOSE && __Handles.hGui = hwnd) { ; fired when closing the Gui
				This.Workarea(A_ScreenWidth, A_ScreenHeight-27)
				eReader.Cleanup(__Handles.wb)
			}
			return
		}
		
		WM_MOUSEMOVE(wParam, lParam, Msg, Hwnd) {
			DllCall("TrackMouseEvent", "UInt", &TME)
			MouseGetPos,,,, MouseCtrl, 2
			GuiControl, % (MouseCtrl = __Handles.hButtonMenuFileText) ? "Show" : "Hide", % __Handles.hButtonMenuFileH
			GuiControl, % (MouseCtrl = __Handles.hButtonMenuEditText) ? "Show" : "Hide", % __Handles.hButtonMenuEditH
			GuiControl, % (MouseCtrl = __Handles.hButtonMenuToolsText) ? "Show" : "Hide", % __Handles.hButtonMenuToolsH
			
		}
		
		WM_LBUTTONUP(wParam, lParam, Msg, Hwnd) {
			DllCall("TrackMouseEvent", "UInt", &TME)
			MouseGetPos,,,, MouseCtrl, 2
			If (MouseCtrl = __Handles.hButtonMenuFileText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " __Handles.hButtonMenuFileText
				Menu, FileMenu, Show, %ctlX%, % ctlY + ctlH
			} Else If (MouseCtrl = __Handles.hButtonMenuEditText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " __Handles.hButtonMenuEditText
				Menu, EditMenu, Show, %ctlX%, % ctlY + ctlH
			} Else If (MouseCtrl = __Handles.hButtonMenuToolsText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " __Handles.hButtonMenuToolsText
				Menu, ToolsMenu, Show, %ctlX%, % ctlY + ctlH
			}
			
			WinGetClass, OutputVar, AHK_ID %MouseCtrl%
			
			if (OutputVar = "EVRVideoHandler") or (OutputVar = "AtlAxWin1") or (OutputVar = "AtlAxWin") or (OutputVar = "EVRVideoHandler1") {
				MouseGetPos, ctlX, ctlY 
				Menu, ContextMenuMedia, Show, %ctlX%, %ctlY%
				
			}
			
		}
		
		EncodeInteger( p_value, p_size, p_address, p_offset ) {
			loop, %p_size%
				DllCall( "RtlFillMemory"
			, "uint", p_address+p_offset+A_Index-1
			, "uint", 1
			, "uchar", ( p_value >> ( 8*( A_Index-1 ) ) ) & 0xFF )
		}
		
		Workarea(Width, Height) {
			VarSetCapacity( area, 16 )
			This.EncodeInteger( Width, 4, &area, 8 )
			This.EncodeInteger( Height, 4, &area, 12 )
			success := DllCall( "SystemParametersInfo", "uint", 0x2F, "uint", 0, "uint", &area, "uint", 0 )
			if ( ErrorLevel or ! success )
			{
				MsgBox, [1] failed: EL = %ErrorLevel%
				;ExitApp
			}
			This.Fix_Maximized_Windows()
		}
		
		Fix_Maximized_Windows() {
			WinGet, OutputVar, List
			Loop % OutputVar
			{
				WinGet, Name, ProcessName, % "AHK_ID " OutputVar%A_Index%
				WinGet, MinMax, MinMax, % "AHK_ID " OutputVar%A_Index%
				If (MinMax = 1) and (Name != "ApplicationFrameHost.exe") {
					WinRestore, % "AHK_ID " OutputVar%A_Index%
					WinMaximize, % "AHK_ID " OutputVar%A_Index%
				}
			}
		}
		
		CreateDIB(Input, W, H, ResizeW := 0, ResizeH := 0, Gradient := 1 ) {
			WB := Ceil((W * 3) / 2) * 2, VarSetCapacity(BMBITS, (WB * H) + 1, 0), P := &BMBITS
			Loop, Parse, Input, |
			{
				P := Numput("0x" . A_LoopField, P + 0, 0, "UInt") - (W & 1 && Mod(A_Index * 3, W * 3) = 0 ? 0 : 1)
			}
			hBM := DllCall("CreateBitmap", "Int", W, "Int", H, "UInt", 1, "UInt", 24, "Ptr", 0, "Ptr")
			hBM := DllCall("CopyImage", "Ptr", hBM, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x2008, "Ptr")
			DllCall("SetBitmapBits", "Ptr", hBM, "UInt", WB * H, "Ptr", &BMBITS)
			If (Gradient != 1) {
				hBM := DllCall("CopyImage", "Ptr", hBM, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x0008, "Ptr")
			}
			return DllCall("CopyImage", "Ptr", hBM, "Int", 0, "Int", ResizeW, "Int", ResizeH, "Int", 0x200C, "UPtr")
		}
		
		Application_Control_Workaround(_Path) {
			Check := FALSE
			Sendinput, {LWin down}{r}{LWin up} 
			WinWait, Run
			ControlSetText, Edit1, %_Path%, Run
			WinWaitClose, Run,, % 1/4
			ControlClick, Button2, Run
		}
		
		QMaster(Check := "") {
			If (Check = "Shutdown") {
				WinShow ahk_class TFormVterm 
				WinMaximize, UC for Business Desktop - Bryn Abbot
				ExitApp
			}
			
			WinGet, Style, Style, UC for Business Desktop - Bryn Abbot
			If (Style & 0x10000000)  {
				WinHide ahk_class TFormVterm
			}
			else { 
				WinShow ahk_class TFormVterm 
				WinMaximize, UC for Business Desktop - Bryn Abbot
			}
		}
		
		emailbuild(GenerateaSuperEmail) {
			
			Client := {}
			
			winactivate, Mochasoft
			
			clipboard := 
			ClipWait, 2
			
			winactivate, Mochasoft
			sendinput, {Alt}ES
			sendinput, ^c
			
			ClipWait, 3
			if ErrorLevel
				MsgBox, The attempt to copy text onto the clipboard failed.
			
			Array := StrSplit(Clipboard, "`n")
			
			if (substr(Array[3],37,9) = "Addresses") {
				Client._AGS := StrReplace(substr(Array[4],22,21),A_Space)
				Client._NameLast := Workbook_Class.Change_String(substr(Array[5],22,21),"Clean")
				Client._NameFirst := Workbook_Class.Change_String(substr(Array[6],22,21),"Clean")
				Client._Email := Workbook_Class.Change_String(substr(Array[11],18,25),"Clean")
				Client._Title := substr(Array[7],22,4)
				;msgbox % Client._Email
				If !InStr(Client._Email, "@")
					Client._Email := Workbook_Class.Change_String(substr(Array[22],18,25),"Clean")
			}
			
			Xl := ComObjActive("Excel.Application")
			ComObjError(false)
			
			For sheet in Xl.Worksheets
				if InStr(Sheet.Name, "OP Investigation Report") 
					Xl.Sheets(Sheet.Name).activate
			
			Pointer := Xl.ActiveSheet.Range("A:H").Find("Fund Code")
			
			if (Xl.Range(Pointer.Offset(0, 1).address).Value = "OTN119")
				SuperProvider := "admincare@australiansuper.com"
			
			if (Xl.Range(Pointer.Offset(0, 1).address).Value = "OTN058")
				SuperProvider := "info@statewide.com.au"
			
			Pointer := Xl.ActiveSheet.Range("A:H").Find("Reason for overpayment")
			Val4 := A_tab . Xl.Range(Pointer.Offset(1, 0).address).Value
			
			Pointer := Xl.ActiveSheet.Range("A:H").Find("Pay Centre")
			
			;msgbox % Floor(Xl.Range(Pointer.Offset(0, 1).address).Value)
			
			if (Floor(Xl.Range(Pointer.Offset(0, 1).address).Value) = "673")
				RecRate := "5%"
			else
				RecRate := "10%"
			
			Pointer := Xl.ActiveSheet.Range("A:H").Find("10% of gross salary")
			RecAmount := Xl.Range(Pointer.Offset(0, 1).address).Value
			
			Pointer := Xl.ActiveSheet.Range("A:H").Find("Tax Diff")
			TAXAMOUNT := RegExReplace(Xl.Range(Pointer.Offset(1, 0).address).Value, "[^0-9.]")
			
			
			
			default_Text =
(

I wish to advise that a salary overpayment has occurred with your pay. 

Please refer to the attached letter for further information. Full details of the overpayment are provided in the attached letter and ‘Paid/Due/Difference’ explanation.

To advise your preferred repayment option, please complete and sign the attached Recovery Authorisation for Salary Overpayment and return to payrolldebtrecovery.asp@nt.gov.au within 10 working days from the date of this email.  

If you wish to repay at less than the scheduled %RecRate% ($%RecAmount%) please complete Section 5 of the form and return for submission to your Agency for their consideration. 

If you have any queries regarding the calculations supplied please contact Payroll Debt Recovery Services.

Overpaid tax of $%TAXAMOUNT% will be recovered by reducing tax payable for the fortnight.  This adjustment will appear on your payslip but will not affect nett pay. 

Overpaid superannuation contributions will be recovered by reducing fortnightly employer contributions until the overpaid amount is recovered. Adjustment will appear on your payslip.

We apologise for any inconvenience this may cause.

)
			
			m :=	ComObjActive("Outlook.Application").CreateItem(0)
			m.Subject := "Salary Overpayment - " . Client._AGS . " | " . Client._NameLast " , " Client._NameFirst
			m.To :=	Client._Email,m.Display
			myInspector :=	m.GetInspector, myInspector.Activate
			wdDoc :=	myInspector.WordEditor
			wdRange :=	wdDoc.Range(0, wdDoc.Characters.Count)
			wdRange.InsertBefore("Dear " . Workbook_Class.Change_String(Client._Title,"Clean") . " " . Workbook_Class.Change_String(Client._NameLast,"Title") . "`n`r" . default_Text)
			m.Attachments.Add("C:\Users\babb\Documents\Blank.txt")
			m.SentOnBehalfOfName := "Payrolldebtrecovery.asp@nt.gov.au"
			m.SendUsingAccount := m.Application.Session.Accounts.Item(1)
			
			
			Val := Client._NameLast
			Val2 := Client._NameFirst
			Val3 := Client._AGS
			
			If (GenerateaSuperEmail = True) {
				SuperText =
(

Please find attached request for reimbursement of overpaid superannuation contributions.

Employee Name: 		%Val%, %Val2%
Employee AGS Number: 	%Val3%
Reason for Overpayment: %Val4%	

Calculations of overpayment are attached for your information.


Super refund to DCIS
A/C Name: DCIS Payroll Super
BSB:	  085 933 
A/C:	  944 008 093
Ref:	  %Val3%

)
				
				m :=	ComObjActive("Outlook.Application").CreateItem(0)
				m.Subject := "Overpaid Super - " . Client._AGS . " | " . Client._NameLast " , " Client._NameFirst
				m.To :=	SuperProvider ,m.Display
				m.cc := "Mary.Lei@nt.gov.au"
				
				myInspector :=	m.GetInspector, myInspector.Activate
				wdDoc :=	myInspector.WordEditor
				wdRange :=	wdDoc.Range(0, wdDoc.Characters.Count)
				wdRange.InsertBefore(SuperText)
				m.Attachments.Add("C:\Users\babb\Documents\Blank.txt")
				m.SentOnBehalfOfName := "Payrolldebtrecovery.asp@nt.gov.au"
				m.SendUsingAccount := m.Application.Session.Accounts.Item(1)
				
				Return
			}
		}
		
	}
		
		
		
		
		Class WMPLayer {
			
			__New() {
				FileSelectFile, SelectedFile, 3, , Open a file
		;SelectedFile := "C:\_Umbra Sector\_LoFi Hiphop\Lofi hip hop mix - Beats to RelaxStudy to [2019].mp4"
				__Handles.WMP.Url := SelectedFile 
				__Handles.WMP.uiMode := "None"
				__Handles.WMP.stretchToFit := 1  
				
			}
			
			Toggle() {
				Static ToggleCheck := 0
				
				If (__Handles.hWMP_STATE != "Active") {
					New WMPLayer
					__Handles.hWMP_STATE := "Active"
				}
				
				If (ToggleCheck = 0) {
					ToggleCheck := !ToggleCheck
					guicontrol, hide, % __Handles.hQuickNotes
					guicontrol, Show, % __Handles.hWMPLayer	
				}
				Else {
					ToggleCheck := !ToggleCheck
					guicontrol, Show, % __Handles.hQuickNotes
					guicontrol, hide, % __Handles.hWMPLayer	
				}
			}
			
			Mode(Selected) {
				__Handles.WMP.uiMode := Selected
			}
			
			Select() {
				FileSelectFile, SelectedFile, 3, , Open a file
				__Handles.WMP.Url := SelectedFile 
			}
			
			Close() {
				__Handles.WMP := ""
			}
		}
		