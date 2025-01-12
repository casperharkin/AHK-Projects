#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; Script:    ClassGUI.ahk
; License:   MIT License
; Author:    Casper Harkin
; Github:    
; Date:      12/01/2025
; Version:   1

/*
    ClassGUI.ahk is an AutoHotkey script designed to provide an easy-to-use GUI framework.

    Features include:
    - Creation of custom GUI controls (buttons, pictures, text).
    - Mouse event handling for hover and click interactions.
    - Dynamic window resizing with automatic adjustment of control positions.
    - Status bar updates to display real-time data such as mouse position and window size.
    - Bitmap handling for button hover effects using DIB (Device-Independent Bitmap) sections.
    
    License:
    This script is released under the MIT License. You are free to use, modify, and distribute 
    the script as long as you include the original license and copyright notice.

    Dependencies:
    - No external libraries are required. This script uses standard AutoHotkey commands and functions.
    
    Author's Note:
    If you encounter any issues or would like to contribute improvements, please feel free to 
    reach out to me through my GitHub repository.

    Example Usage:
    To use this class, instantiate an object of ClassGUI with a specific window size and title.
    For example:

		Gui := New GUI(800, 500, "Simple Text Viewer")
		Gui.AddMenu("FileMenu", "Open Text File", "OpenTextFile")
		Gui.AddMenu("EditMenu", "Upper Case Clipboard", "FormatStr")
		Gui.AddMenu("EditMenu", "Title Case Clipboard", "FormatStr")
		Gui.AddMenu("EditMenu", "Lower Case Clipboard", "FormatStr")
		Gui.AddMenu("ToolsMenu", "Toggle Bluetooth", "ToggleBluetooth")
		
		Gui, Font, cWhite +bold, Segoe UI , s12
		gui, add, text, x230 y15 +BackgroundTrans +0x4E, Simple Text Viewer
		
		Gui, Font,  
		gui, add, edit, x10  y80 w780 h350 +hwndhEditControl ,
		Exit ; End of AES
		
		OpenTextFile(){
			FileSelectFile, SelectedFile, 3, , Open a file, Text Documents (*.txt)
			if SelectedFile {
				FileRead, File, % SelectedFile
				GuiControl,, Edit1, % File
			}
		}	
		
		FormatStr(){
			Clipboard := (Format("{:" Substr(A_ThisMenuItem, 1, 1) . "}", Clipboard))
		}
		
		ToggleBluetooth(){ ; Script Needs to be run as Admin
		    static t := True
		    t := !t 
		    Run % "PowerShell -Command ""(Get-PnpDevice -Class Bluetooth).Where({$_.Status " (t ? "-ne 'OK'}) | Enable" : "-eq 'OK'}) | Disable") "-PnpDevice -Confirm:$false""", , Hide
		}

*/


Class GUI {

	id := {} ;track hwnds
	Bitmaps := {Normal: This.CreateDIB(0x0173C7, 70, 50), Hover: This.CreateDIB(0x2A8AD4, 70, 50)} ;Set BITMAPs ;ARGB Colour, w, h

__New(w,h, title, Options := ""){	

	Gui, % "+HwndhGuiWindow " Options
	This.id.hGuiWindow := hGuiWindow

	Gui, Color, White 
	Gui, Font, cWhite +bold, Segoe UI , s12

	This.AddPicture("x0 y0 w500 h70 +BackgroundTrans +0x4E Hidden0 +HWND", "hBackgroundN")

	This.AddButton(10, 10, "File")
	This.AddButton(80, 10, "Edit")
	This.AddButton(150, 10, "Tools")

	OnMessage(0x200, ObjBindMethod(This, "WM_MOUSEMOVE"))
	OnMessage(0x202, ObjBindMethod(This, "WM_LBUTTONUP"))
	OnMessage(0x5, ObjBindMethod(this, "WM_SIZE"))

	Gui, Font, 

	Gui, Add, StatusBar,,
	SB_SetParts(200)
	Gui, Show, % "w" w " h" h, % title
}


AddButton(X, Y, Text) {
	BaseName := "hButtonMenu" Text
	This.AddPicture("x" X " y" Y " w70 h50 +BackgroundTrans +0x4E Hidden0 +HWND", BaseName "N")
	This.AddPicture("x" X " y" Y " w70 h50 +BackgroundTrans +0x4E Hidden1 +HWND", BaseName "H")
	This.AddText("x" X " y" Y " w70 h50 +BackgroundTrans +0x201 +Center +HWND", BaseName "Text", "`n" Text)
}

AddText(Options, Name, Text){
	Gui, Add, Text, % Options . Name, % Text
	This.id[Name]  := %Name%
}

AddPicture(Options, Name){
	Gui, Add, Picture, % Options . Name, % "HBITMAP:*" (Substr(Name, -0) = "N" ? This.Bitmaps.Normal : This.Bitmaps.Hover)
	This.id[Name]  := %Name% ;Add hwnd of control under its name to the id object
}

WM_LBUTTONUP(wParam, lParam, Msg, Hwnd){
	MouseGetPos,,,, MouseCtrl, 2

	try {
			If (MouseCtrl = This.id.hButtonMenuFileText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " This.id.hButtonMenuFileText
				Menu, FileMenu, Show, %ctlX%, % ctlY + ctlH
			} Else If (MouseCtrl = This.id.hButtonMenuEditText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " This.id.hButtonMenuEditText
				Menu, EditMenu, Show, %ctlX%, % ctlY + ctlH
			} Else If (MouseCtrl = This.id.hButtonMenuToolsText) {
				ControlGetPos, ctlX, ctlY, ctlW, ctlH, , % "ahk_id " This.id.hButtonMenuToolsText
				Menu, ToolsMenu, Show, %ctlX%, % ctlY + ctlH
			}
	}
}

WM_MOUSEMOVE(wParam, lParam, Msg, Hwnd){
	MouseGetPos,x,y,, MouseCtrl, 2 ;Track mouse as it moves, Mode 2 - save hwnd of the control it is over
	GuiControl, % (MouseCtrl = This.id["hButtonMenuFileText"]) ? "Show" : "Hide", % This.id["hButtonMenuFileH"]
	GuiControl, % (MouseCtrl = This.id["hButtonMenuEditText"]) ? "Show" : "Hide", % This.id["hButtonMenuEditH"]   
	GuiControl, % (MouseCtrl = This.id["hButtonMenuToolsText"]) ? "Show" : "Hide", % This.id["hButtonMenuToolsH"] 
	SB_SetText("Mouse X: " x " Y:" y, 2) 
}

WM_SIZE(wParam, lParam, Msg, Hwnd){
	Width := lParam & 0xFFFF ; lower 16 bits of lParam represent the new width.
	Height := lParam >> 16 ;upper 16 bits of lParam represent the new height.
	GuiControl, Move, % This.id["hBackgroundN"], % "w" . Width
	Tooltip % "Width: " Width "`nHeight: " Height
	SB_SetText("Width: " Width "`nHeight: " Height, 1)
}

CreateDIB(Colour, W, H) {  ; https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader | Numputs ; Size ; Width ; Height (Negative so (0, 0) is top-left) ; Planes ; BitCount / BitsPerPixel
    bi := VarSetCapacity(bi, 40, 0), NumPut(40, bi, 0, "UInt"), NumPut(W, bi, 4, "Int"), NumPut(-H, bi, 8, "Int"), NumPut(1, bi, 12, "UShort"), NumPut(32, bi, 14, "UShort")
    hbm := DllCall("CreateDIBSection", "Ptr", 0, "Ptr", &bi, "UInt", 0, "PtrP", pBits := 0, "Ptr", 0, "UInt", 0, "Ptr") ; Create a DIBSection and get a pointer (pBits) to pixel data
    Loop, % W * H ; Fill the pixel data with the specified Colour
        NumPut(Colour, pBits + (A_Index - 1) * 4, "UInt") ; ARGB Colour
    Return hbm
}

AddMenu(Menu, Text, Handler){
	Menu % Menu, Add, % Text, % Handler
}

}


