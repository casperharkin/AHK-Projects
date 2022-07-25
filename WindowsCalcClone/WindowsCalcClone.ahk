	#NoEnv
	#MaxHotkeysPerInterval 99000000
	#HotkeyInterval 99000000
	#KeyHistory 0
	#SingleInstance Force

	ListLines Off
	Process, Priority, , A
	SetBatchLines, -1
	SetKeyDelay, -1, -1
	SetMouseDelay, -1
	SetDefaultMouseSpeed, 0
	SetWinDelay, -1
	SetControlDelay, -1
	SendMode Input
	
;	Menu, Tray, Icon, % "C:\_Umbra Sector\_Projects\New Calc\Windows-10-Calculator-Fluent-Icon-Big-256.ico"

	Global TME, GUI := {} , BottomPanelSize 
	BuildGUI() 
	

	Exit ; End of AES


	
	#IfWinActive Calculator ahk_class AutoHotkeyGUI
	Numpad0::
	Numpad1::
	Numpad2::
	Numpad3::
	Numpad4::
	Numpad5::
	Numpad6::
	Numpad7::
	Numpad8::
	Numpad9::
	NumpadAdd::
	NumpadSub::
	NumpadMult::
	NumpadDiv::
	NumpadDot::
	NumpadEnter::
	Backspace::
	Delete::

		IF (A_ThisHotkey = "Backspace")
			DeleteLastFig()
	
		IF (A_ThisHotkey = "Delete")
		{
			ControlSetText,,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
			ControlSetText,,, % "AHK_ID " GUI["EditTOP"]["Hwnd"].U
		}
	
		GuiControl, Show, % GUI["BTN_" temp := StrReplace(A_ThisHotkey,"Numpad")]["Hwnd"].H
		ControlClick,, % "AHK_ID " GUI["BTN_" temp := StrReplace(A_ThisHotkey,"Numpad")]["Hwnd"].U
		GuiControl, Hide, % GUI["BTN_" temp := StrReplace(A_ThisHotkey,"Numpad")]["Hwnd"].H

	Return

	^v::
		ControlGetText, int,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		PastedFig := RegExReplace(Clipboard, "[^0-9.]")
		ControlSetText,,% int . PastedFig, % "AHK_ID " GUI["Edit"]["Hwnd"].U
	Return
	
	#If
	
	ResizeFont(){
		Count++
		If (Count = 1)
			Size := 20
	
		While(Size>1){
			GuiControlGet,Poss,Pos, % GUI["Edit"]["Hwnd"].U

			GuiControlGet, OutputVar,, % GUI["Edit"]["Hwnd"].U
			Gui,Fake:Font,s%Size%,Segoe UI
			Gui,Fake:Add,Text, -Wrap +hwndhDummy,% OutputVar
			GUI["MainGUI"]["DummyHWND"] := hDummy
			GuiControlGet,Pos,Fake:Pos, % GUI["MainGUI"]["DummyHWND"]
			Gui,Fake:Destroy
		
			;If(posw>252) Old Version
			
			If(posw>270){
				Size := Size - 3
				BottomPanelSize := Size
				Gui,Font, % "s" Size 
				GuiControl,Font, % GUI["EditTOP"]["Hwnd"].U
				GuiControl,Font, % GUI["Edit"]["Hwnd"].U
				
			}else	{
				break
			}
		}
	}
	
	AddControl(ControlType, Name_Control, Options := "", Value := "", DIB := "") {


		Gui, Font,, % StrSplit(Options,"|").3
	
		If (StrSplit(Options,"|").2 != "")
			Gui Font, % StrSplit(Options,"|").2
	
		Options := StrSplit(Options,"|").1
	
		Gui Add, Picture, % Options " +BackgroundTrans +0x4E +HWNDh" Name_Control "N Hidden0" 
		Gui Add, Picture, % Options " +BackgroundTrans +0x4E +HWNDh" Name_Control "H Hidden1" 
		Gui Add, % ControlType, % Options " BackgroundTrans 0x200 +HWNDh" Name_Control "U Hidden0", % Value 
	
		If (DIB != "") {
			DllCall("SendMessage", "Ptr", h%Name_Control%N, "UInt", 0x172, "Ptr", 0, "Ptr", CreateDIB(StrSplit(DIB,"|").1, 1, 1))
		 	DllCall("SendMessage", "Ptr", h%Name_Control%H, "UInt", 0x172, "Ptr", 0, "Ptr", CreateDIB(StrSplit(DIB,"|").2, 1, 1))
	
		}
	
		GUI[Name_Control] := {"Hwnd":{"N":h%Name_Control%N,"H":h%Name_Control%H,"T":h%Name_Control%T,"U":h%Name_Control%U}
						 	 ,"Options":Options
						 	 ,"Value":Value}
		Gui Font,
	
		If (ControlType = "Text") {
		ControlHandler := Func("Process_String").Bind(h%Name_Control%U)
		GuiControl +g, % h%Name_Control%U, % ControlHandler
		}
	}
	
	Process_String(){
	
		DllCall("TrackMouseEvent", "UInt", &TME)
		MouseGetPos, , , , controlUnderMouse, 2
	
		If (controlUnderMouse = GUI["StaticTitle"]["Hwnd"].u) 
		{
			PostMessage, 0xA1, 2,,, A ; WM_NCLBUTTONDOWN
			exit
		}
	
		ControlGetText, int,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		ResizeFont()
	
		If (Int = "0") 
		{
			ControlSetText,,  , % "AHK_ID " GUI["Edit"]["Hwnd"].U
			ControlGetText, int,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		}
	
		If A_GuiControl is number
			ControlSetText,,% int . A_GuiControl, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = Chr(0xE948)) ;+
			ControlSetText,,% int "+", % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = Chr(0xE947)) ;x
			ControlSetText,,% int "*", % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = Chr(0xE94A)) ;/
			ControlSetText,,% int " / ", % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = ".") ;.
			ControlSetText,,% int ".", % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = Chr(0xE949)) ;.
			ControlSetText,,% int "-", % "AHK_ID " GUI["Edit"]["Hwnd"].U
		Else If (A_GuiControl = Chr(0xE94F)) ; <=  
				DeleteLastFig()
		Else If (A_GuiControl = "C") or (A_GuiControl = "CE") ; C or CE
		{
			ControlSetText,,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
			ControlSetText,,, % "AHK_ID " GUI["EditTOP"]["Hwnd"].U
		}
		Else If (A_GuiControl = Chr(0xE94E)) ;=
		{
			ControlSetText,,% int . " =", % "AHK_ID " GUI["EditTOP"]["Hwnd"].U
			Result := Eval(int)
			Rounded := Round(Eval(int),4)
			Cleaned := StrReplace(Rounded , ".0000")
			ControlSetText,,% Cleaned, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		}
		Else If (Chr(0xEF2C) = A_GuiControl)
			ExitApp
			obj :=
		}
		
		DeleteLastFig(){
			ControlGetText, int,, % "AHK_ID " GUI["Edit"]["Hwnd"].U
			Array := StrSplit(int)
			Loop % Array.MaxIndex()-1
				r.= Array[A_Index]
			ControlSetText,, % r, % "AHK_ID " GUI["Edit"]["Hwnd"].U
		}
		
	
	
		WM_MOUSEMOVE(wParam, lParam, Msg, Hwnd) 
		{
	
		DllCall("TrackMouseEvent", "UInt", &TME)
			MouseGetPos,,,, MouseCtrl, 2
			
			GuiControl, % (MouseCtrl = GUI["menu"]["Hwnd"].U) ? "Show" : "Hide", % GUI["menu"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["Close"]["Hwnd"].U) ? "Show" : "Hide", % GUI["Close"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["History"]["Hwnd"].U) ? "Show" : "Hide", % GUI["History"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_M_Plus"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_M_Plus"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_MC"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_MC"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_MR"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_MR"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_M_Minus"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_M_Minus"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_MS"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_MS"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_M_Dot"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_M_Dot"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Delete"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_Delete"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_C"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_C"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_CE"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_CE"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Perc"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_Perc"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_1byx"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_1byx"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_2x"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_2x"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_sqr"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_sqr"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Div"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_Div"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_7"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_7"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_8"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_8"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_9"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_9"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Mult"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_Mult"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_4"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_4"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_5"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_5"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_6"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_6"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Sub"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_Sub"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_1"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_1"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_2"]["Hwnd"].U) ? "Show" : "Hide", % GUI["BTN_2"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_3"]["Hwnd"].U) 	? "Show" : "Hide", % GUI["BTN_3"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Add"]["Hwnd"].U)	? "Show" : "Hide", % GUI["BTN_Add"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_AddMinus"]["Hwnd"].U)	? "Show" : "Hide", % GUI["BTN_AddMinus"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_0"]["Hwnd"].U)	? "Show" : "Hide", % GUI["BTN_0"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Dot"]["Hwnd"].U)	? "Show" : "Hide", % GUI["BTN_Dot"]["Hwnd"].H
			GuiControl, % (MouseCtrl = GUI["BTN_Enter"]["Hwnd"].U)	? "Show" : "Hide", % GUI["BTN_Enter"]["Hwnd"].H
		}
	
	ON_EN_SETFOCUS(wParam, lParam) {
		DllCall("user32\HideCaret", "ptr", "AHK_ID " GUI["Edit"]["Hwnd"].U)
	}
	
	BuildGUI() {
		
		Gui +LastFound -Resize -Caption +Border +hwndhWnd -ToolWindow +hwndhWnd -SysMenu
		Gui color, E6E6E6
	
		GUI["MainGUI"] := {"hWnd":hWnd,"FontSize":50,"DummyHWND":"NA"}
		
		AddControl("Text", "StaticTitle", " x12 y7 w250 h16| s13","Calculator")
	
		AddControl("Text", "Menu", " x8 y35 w32   +Center| s25|Segoe MDL2 Assets", Chr(0xE8C4),"F1F1F1|C7C7C7")
		AddControl("Text", "History", " x300 y45 w32 +Center |s18|Calculator MDL2 Assets",Chr(0xE81C),"E6E6E6|E81123")
		AddControl("Text", "Close", " x312 y0 w28 h28  +Center|s18|Segoe MDL2 Assets",Chr(0xEF2C),"E6E6E6|E81123")

		AddControl("Edit", "EditTOP", " x14 y90 w310 h35  12 -border +readonly +0x4E 0x802 -E0x200 +Cener -VScroll|Bold s20|Calculator MDL2 Assets","")
		AddControl("Edit", "Edit", " x14 y120 w310 h45  12 -border +readonly +0x4E 0x802 -E0x200 +Cener -VScroll|Bold s20|Calculator MDL2 Assets","0")
		
		AddControl("Text", "BTN_MC", "x16 y175 w48 h20 +Disabled +Center|s13|Segoe MDL2 Assets","MC","E6E6E6|C7C7C7")
		AddControl("Text", "BTN_MR", "x68 y175 w48 h20 +Disabled +Center|s13 Bold|Segoe MDL2 Assets","MR","E6E6E6|C7C7C7")
		AddControl("Text", "BTN_M_Plus", "x120 y175 w49 h20 +Center|s13|Segoe MDL2 Assets","M+","E6E6E6|C7C7C7")
		AddControl("Text", "BTN_M_Minus", " x173 y175 w48 h20 +Center|s13|Segoe MDL2 Assets","M-","E6E6E6|C7C7C7")
		AddControl("Text", "BTN_MS", "x223 y175 w48 h20 +Center|s13|Segoe MDL2 Assets","Ms","E6E6E6|C7C7C7")
		AddControl("Text", "BTN_M_Dot", "x274 y175 w48 h20 +Disabled +Center|s13|Segoe MDL2 Assets","M.","E6E6E6|C7C7C7")
		
		AddControl("Text", "BTN_Perc","x14 y208 w75 h49 +Center|  s13|Segoe MDL2 Assets", Chr(0xE94C),"F1F1F1|C7C7C7")
		AddControl("Text", "BTN_CE","x92 y208 w75 h49 +Center|  s16", "CE","F1F1F1|C7C7C7")
		AddControl("Text", "BTN_C","x170 y208 w75 h49 +Center |  s16", "C","F1F1F1|C7C7C7")
		AddControl("Text", "BTN_Delete"," x250 y208 w74 h49 +Center|  s13|Calculator MDL2 Assets", Chr(0xE94F),"F1F1F1|C7C7C7")
		
		AddControl("Text", "BTN_1byx", "x13 y261 w75 h49 +Center|  s13|Calculator MDL2 Assets", Chr(0xF7C9),"F1F1F1|C7C7C7")
		AddControl("Text", "BTN_2x", "x92 y261 w75 h49 +Center|  s13|Calculator MDL2 Assets", "x2","F1F1F1|C7C7C7")
		AddControl("Text", "BTN_sqr", "x171 y261 w75 h49 +Center|  s13|Calculator MDL2 Assets", Chr(0x221A),"F1F1F1|C7C7C7")
		AddControl("Text", "BTN_Div", " x250 y261 w74 h49 +Center|  s13|Calculator MDL2 Assets", Chr(0xE94A),"F1F1F1|C7C7C7")
		
		AddControl("Text", "BTN_7"," x13 y314 w75 h49 +Center |Bold s16", "7","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_8"," x92 y314 w75 h49 +Center |Bold s16", "8","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_9"," x171 y314 w75 h49 +Center |Bold s16", "9","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_Mult"," x249 y314 w75 h49 +Center |  s13|Calculator MDL2 Assets", Chr(0xE947),"F1F1F1|C7C7C7")
		
		AddControl("Text", "BTN_4","x12 y367 w75 h49 +Center |Bold s16", "4","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_5"," x92 y367 w75 h49 +Center |Bold s16", "5","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_6"," x171 y367 w75 h49 +Center |Bold s16", "6","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_Sub"," x250 y367 w74 h49 +Center |  s13|Calculator MDL2 Assets", Chr(0xE949),"F1F1F1|C7C7C7")
		
		AddControl("Text", "BTN_1", "x13 y420 w75 h49 +Center |Bold s16", "1","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_2", "x92 y420 w75 h49 +Center |Bold s16", "2","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_3"," x171 y420 w75 h49 +Center |Bold s16", "3","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_Add"," x250 y420 w75 h49 +Center |  s13|Calculator MDL2 Assets", Chr(0xE948),"F1F1F1|C7C7C7")
		
		AddControl("Text", "BTN_AddMinus", " x13 y472 w75 h49 +Center |  s13|Segoe MDL2 Assets", Chr(0xE94D),"F1F1F1|C7C7C7")
		AddControl("Text", "BTN_0", "x92 y473 w75 h48 +Center |Bold s16", "0","F7F7F7|C7C7C7")
		AddControl("Text", "BTN_Dot", "x171 y473 w75 h48 +Center |  s13", ".","F1F1F1|C7C7C7")
		AddControl("Text", "BTN_Enter", "x250 y472 w75 h49 +Center | Default s13|Calculator MDL2 Assets", Chr(0xE94E),"8ABAE0|4599DB")
	
		GuiControl, Focus, Static3
		Gui Show, w339 h538 , Calculator 
	
		VarSetCapacity(TME,16,0)
		NumPut(16, TME, 0)						; cbSize
		NumPut(1, TME, 4)						; dwFlags = TME_HOVER (1) | TME_LEAVE (2)
		NumPut(GUI["MainGUI"].Hwnd, TME, 8)		; hwndTrack
		OnMessage(0x200, "WM_MOUSEMOVE")			; 0x200 = WM_MOUSEMOVE
		OnMessage(0x0111, "ON_EN_SETFOCUS")		; 0x0111 = ON_EN_SETFOCUS - Sent when an edit control receives the input focus.
	}
	
	CreateDIB(Input, W, H, ResizeW := 0, ResizeH := 0, Gradient := 1 ) {
		_WB := Ceil((W * 3) / 2) * 2, VarSetCapacity(BMBITS, (_WB * H) + 1, 0), _P := &BMBITS
		Loop, Parse, Input, |
		{
			_P := Numput("0x" . A_LoopField, _P + 0, 0, "UInt") - (W & 1 && Mod(A_Index * 3, W * 3) = 0 ? 0 : 1)
		}
		hBM := DllCall("CreateBitmap", "Int", W, "Int", H, "UInt", 1, "UInt", 24, "Ptr", 0, "Ptr")
		hBM := DllCall("CopyImage", "Ptr", hBM, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x2008, "Ptr")
		DllCall("SetBitmapBits", "Ptr", hBM, "UInt", _WB * H, "Ptr", &BMBITS)
		If (Gradient != 1) {
			hBM := DllCall("CopyImage", "Ptr", hBM, "UInt", 0, "Int", 0, "Int", 0, "UInt", 0x0008, "Ptr")
		}
		return DllCall("CopyImage", "Ptr", hBM, "Int", 0, "Int", ResizeW, "Int", ResizeH, "Int", 0x200C, "UPtr")
	}
	
	
	Eval(Expr, Format := FALSE) {
	    static obj := ""    
	    If ( !obj )
	        obj := ComObjCreate("HTMLfile")
	
	    Expr := StrReplace( RegExReplace(Expr, "\s") , ",", ".")
	    Expr := RegExReplace(StrReplace(Expr, "**", "^"), "(\w+(\.*\d+)?)\^(\w+(\.*\d+)?)", "pow($1,$3)")    ; 2**3 -> 2^3 -> pow(2,3)
	    Expr := RegExReplace(Expr, "=+", "==")    ; = -> ==  |  === -> ==  |  ==== -> ==  |  ..
	    Expr := RegExReplace(Expr, "\b(E|LN2|LN10|LOG2E|LOG10E|PI|SQRT1_2|SQRT2)\b", "Math.$1")
	    Expr := RegExReplace(Expr, "\b(abs|acos|asin|atan|atan2|ceil|cos|exp|floor|log|max|min|pow|random|round|sin|sqrt|tan)\b\(", "Math.$1(")
	
	    obj.write("<body><script>document.body.innerText=eval('" . Expr . "');</script>")
	    Expr := obj.body.innerText
	
	    return InStr(Expr, "d") ? "" : InStr(Expr, "false") ? FALSE    ; d = body | undefined
	                                 : InStr(Expr, "true")  ? TRUE
	                                 : ( Format && InStr(Expr, "e") ? Format("{:f}",Expr) : Expr )
	} ; CREDITS (tidbit) - https://autohotkey.com/boards/viewtopic.php?f=6&t=15389
	