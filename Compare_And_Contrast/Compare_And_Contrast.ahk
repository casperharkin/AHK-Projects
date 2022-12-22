#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Examples of checking dif types of data. 
Report := " "

If !Check(Obj := {Review:"     873 95008", Against:"873-95008", Subject:"AGS", DataType: "AGS"})
	Report .= "`n" Obj.Subject " is Wrong"

If !Check(Obj := {Review:"    1,337.56", Against:"$1337.5560000 ", Subject:"Total Amount", DataType:"Figure"})
	Report .= "`n" Obj.Subject " is Wrong"

If !Check(Obj := {Review:"8/1/2023 ", Against:"08/01/2023", Subject:"FBT Date", DataType:"Date"})
	Report .= "`n" Obj.Subject " is Wrong"

If !Check(Obj := {Review:"10/O1/2023 ", Against:"10/01/2023", Subject:"Detected Date", DataType:"Date"})
	Report .= "`n" Obj.Subject " is Wrong"

If !Check(Obj := {Review:"  po box 12345", Against:"PO BOX 12345", Subject:"Postal Address", DataType:"Text"})
	Report .= "`n" Obj.Subject " is Wrong"

If (Report != " ")
	MsgBox % Report 

ExitApp ;EOAES   


Check(Obj){ ; This is a helper function to go along with Compare_And_Contrast.Compare
	Return Compare_And_Contrast.Compare(Obj["Review"], Obj["Against"], Obj["Subject"], Obj["DataType"])
}

Class Compare_And_Contrast { 


/* 

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ####   Compare and Contrast   ####
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Aim:
I have a need to compare data against different sources. 

Issue:
The issue is that each source can have different formatting, rounding, spacing etc. for the data. Simple comparisons can return false negatives. 

Examples of False Negatives I Wish to Avoid: 
•	Source A displays the data as $1,337.00, Source B displays the data as $1337 and Source C displays the data as 1337.00 (Formatting)
•	Source A displays the data as "Handcock Street ", Source B displays the data as "Handcock  Street" (Spacing)
•	Source A displays the data as $1337.1001, Source B displays the data as $1337.10 (Rounding)
•	Source A displays the data as 01/01/2021, Source B displays the data as 1/1/2021 (Leading Zeros)

*/

	Compare(Compare, Against, Context := "", DataType := "", Format := ""){

		;------------------------------------------
		;----------------[Settings]----------------
		;------------------------------------------

		This.ShowGuiOnSuccess := 1 ;Show the Data Compare GUI even on successful matches. 

		;------------------------------------------
		;------------------------------------------

		if !Compare or !Against {
			MsgBox % "Error - Review or Against Empty`n`nReview: " Compare "`nAgainst: " Against
			Return 0
		}

		This.Properties := {}

		This.Properties["OG_Compare"] := This.Properties["Compare"] := Trim(Compare), This.Properties["Context"] := Context 
		This.Properties["OG_Against"] := This.Properties["Against"] := Trim(Against), This.Properties["Format"] := Format

		if (DataType = "AGS"){
			If InStr(This.Properties["Compare"], "-")
				This.Properties["Compare"] := StrReplace(Compare, "-")
			else If InStr(This.Properties["Compare"], " ")
				This.Properties["Compare"] := StrReplace(Compare, A_Space)

			If InStr(This.Properties["Against"], "-")
				This.Properties["Against"] := StrReplace(Against, "-")
			else If InStr(This.Properties["Against"], " ")
				This.Properties["Against"] := StrReplace(Against, A_Space)
		}

		if (DataType = "Figure"){
				This.Properties["Compare"] := Round(StrReplace(StrReplace(Compare, "$"), ","), 2)
				This.Properties["Against"] := Round(StrReplace(StrReplace(Against, "$"), ","), 2)
		}

		if (DataType = "Date"){
			CompareDate := StrSplit(This.Properties["Compare"], "/")
			AgainstDate := StrSplit(This.Properties["Against"], "/")

			If (This.CompareObj(CompareDate, AgainstDate) = 1)
				This.Properties["Compare"] := This.Properties["Against"] := This.Properties["Against"]
		}

		if (DataType = "Text"){
			This.Properties["Compare"] := Format("{:U}",This.Properties["Compare"])
			This.Properties["Against"] := Format("{:U}",This.Properties["Against"])
		}

		This.DirectCompare()
		Return This.Properties["Response"]
	}

	CompareObj(Obj1, Obj2){
		for e, i in Obj1 {
			if (Obj1[e] != Obj2[e])
				Return 0
		}
		return 1
	}

	DirectCompare(){
		if (This.Properties["Compare"] = This.Properties["Against"]){
			This.Properties["Response"] := 1

			if (This.ShowGuiOnSuccess = 1)
				This.GuiCompare("Sucsess")
		}
		Else
			This.GuiCompare()
	}

	CnCMenuHandler(){
		if (A_ThisMenuItem = "lowercase"){
			LV_Modify(1,,, t := Format("{:L}",This.Properties["Compare"]))
			LV_Modify(2,,, t := Format("{:L}",This.Properties["Against"]))
		}
		if (A_ThisMenuItem = "UPPERCASE"){
			LV_Modify(1,,, t := Format("{:U}",This.Properties["Compare"]))
			LV_Modify(2,,, t := Format("{:U}",This.Properties["Against"]))
		}
		if (A_ThisMenuItem = "TitleCase"){
			LV_Modify(1,,, t := Format("{:T}",This.Properties["Compare"]))
			LV_Modify(2,,, t := Format("{:T}",This.Properties["Against"]))
		}
		if (A_ThisMenuItem = "Original Formating"){
			LV_Modify(1,,, This.Properties["OG_Compare"])
			LV_Modify(2,,, This.Properties["OG_Against"])	
		}
		if (A_ThisMenuItem = "E&xit"){
			ExitApp
		}
		if (A_ThisMenuItem = "&About"){
			This.About()
		}		
	}

	GuiCompare(Sucsess := ""){

		CnCMenuHandler := ObjBindMethod(This, "CnCMenuHandler")
			
		Gui, New

		Menu, FileMenu, Add, E&xit,  % CnCMenuHandler
		Menu, HelpMenu, Add, &About,  % CnCMenuHandler
		Menu, FormatMenu, Add, Original Formating,  % CnCMenuHandler
		Menu, FormatMenu, Add, UPPERCASE,  % CnCMenuHandler
		Menu, FormatMenu, Add, lowercase,  % CnCMenuHandler
		Menu, FormatMenu, Add, TitleCase,  % CnCMenuHandler

		Menu, MyMenuBar, Add, &File, :FileMenu  
		Menu, MyMenuBar, Add, &Format, :FormatMenu  
		Menu, MyMenuBar, Add, &Help, :HelpMenu

		Gui, Menu, MyMenuBar

		If !This.Properties["Context"]
			This.Properties["Context"] := "Compare and Contrast:"

		Gui ,Font, s12 +Bold, Segoe UI ; Set font options
		Gui, Add, Text, x0 y9 w275 h30 +Center r2 +HwndhContextText cBold, % This.Properties["Context"] ":"
		Gui, Font, 

		Gui, Add, ListView, x12 y39 w250 h100 +HWNDhListView cRed -ReadOnly +AltSubmit -Multi , Source:|Formatted:|Original:

		LV_Add("", "Review: ", This.Properties["Compare"], This.Properties["OG_Compare"])
		LV_Add("", "Against: ", This.Properties["Against"], This.Properties["OG_Against"])
			LV_ModifyCol() 
			LV_ModifyCol(1, "AutoHdr") 
			LV_ModifyCol(2, "AutoHdr") 

		Gui, Add, Button, x12 y145 w100 h30 +HwndhYes +Default, Yes
			This.Bind(hYes, "GuiResponse", 1)

		Gui, Add, Button, x162 y145 w100 h30 +HwndhNo, No
			This.Bind(hNo, "GuiResponse", 0)

		if Sucsess {
			Gui, Font, cGreen ,
			GuiControl, Font, % hListView 
			LV_ModifyCol() 
			LV_ModifyCol(1, "Left") 
			LV_ModifyCol(2, "Left") 
		}		

		Gui, Show, w275 , Review Data

		WinWaitClose, Review Data
	}

	GuiResponse(Response){
		This.Properties["Response"] := Response
		This.GuiDestroy()
	}

	GuiDestroy(){
		Gui, +LastFound
		Gui, Destroy
	}

	CheckInString(String, ArrayOfStrings){
		for e, i in ArrayOfStrings
			If InStr(String, i)
				Return 1
		Return 0
	}

	Bind(Hwnd, Method, Params*){
		BoundFunc := ObjBindMethod(This, Method, Params*)
		GuiControl +g, % Hwnd, % BoundFunc
   	}

	About(){
		s := "iVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAMAAABrrFhUAAAACXBIWXMAAC4jAAAuIwF4pT92AAAANlBMVEXm5uYAAAAyMjINDQ3W" 
		s .= "1tYlJSUAAAAAAABBQUEAAAB+fn66urqenp4AAAAAAABgYGAAAABHcEzqFsqYAAAAEnRSTlP////z+e8UsuuG7vHsWdHuMgAER7Qq" 
		s .= "AAANQklEQVR42u2diZK1qA6A40FREJd+/5cdF1A8emRH8B/mVt2tutt8ZiMJCH/xF8ZjR+q67nuEigKhvp/+C+lGjB94GIgt+yT6" 
		s .= "KvhpoX7GEJtCTACL8FeiHylMEN4IAHdK4SUI8RhArHffF0arj6UHMQCM5Ordo2lVy5r/E7rQAzK+AsBI0LfkVTNQ1rblttqW0aGp" 
		s .= "vjGgCAgCA8Bf4iPUUDZJDBdr+p8ZbVBkBGEBjLUszyo83K4TBFSP2QLA8ttHFW0Vwm8QWlrJP0lwlgBwJzn+amCa0nMGbKikkBAw" 
		s .= "IkAM7a+okfScAa1i2AGE134r8U8IQtlBGAC43k2fWYq/WoJAgIpA8SAIgN36rd/+hRb0XS4ANvVHQwvOqx3Qlh7jHADs6l8x8LEk" 
		s .= "O6hx+gDGenv9JXha5aYE/qMBhJLf0+vna1MC7wQ8AxiF+2ta8LraRjiCMWUAwv17VP+TGXgOBhBEfgoBFg1CAALoP2IQZDEUwArA" 
		s .= "v/xVIPnDEADv/r9qIdhqK++xwBsAHEF+iQBODQAmYe3/2wq8ZcW+AHRx5N8JdGkBGFEk+TcCaEwJAOYBgEKERXkowOkAEA7Af/53" 
		s .= "nRP6dANeAHSrVlZR5J8INB7dgA8APANCLURaLfKXD4E3A0AUoi2+LSBpAOAGMJTxAMDqBlCXAgCeAlYx5Yey8pUQugPgEYBB1MVj" 
		s .= "IXkeAPeATRkXAI8E7n4Q/ChAvAjwFQnI0wC4AgwQfVE/KgBeFKAq4wPgfpA8C4ArAAXIVQUgigKUZWmoIxo/4UcF3ABgHQUo2dBU" 
		s .= "VdVQbUfZ0uUHVEMV1Meu0A1Ap66CSe1dpNctafcZofvmMq+Pdc8B4Ekg1eno8P2C0hJKepyrapUq4JYOOgFY60B3OUBbfY0/qnYM" 
		s .= "YrO/z5e0qlzArTYE7i6wMZBf1TQ7yX9PoHF3g+DuApnKTy8DkptAVCPFXwZpNwK/kTF3GwBnC7h5vmGbjf18KjH6h5Q6PU8Wfaaf" 
		s .= "EAwGBWGnXTE4W8CgrN/O0syLV05/75tErUv8AGd2g2xwtgFwjQE3lfDhKP8kkEIeUerafwApvMzK2MUGIJwFcHF2+YU8w71PQ5+P" 
		s .= "NjJuA+MjADqFBdCTOEKeH8zKEzGBjCq0rHsEALl/Nm7QsjifD7qzmlWfD8Q4st9ug7o6AXAMgjfaeSEOl4fevc0jsVUFkMLOHPYD" 
		s .= "ENgFXAJo7lzAEcAHBXYCEMwFXCn053Oj0dxmPh99o/HgBCCUC+BZGrqS51pt+Mv8XOkMg1BOwB6AKgugVwpdGQNAiq4TzwTiA+Bp" 
		s .= "UGmoAYUKQGGoAaVjKgRuQeBuo3LjBEvtuKn0AVxv4gNYg8BNP+QmCgw3/gyZRQGOzT4MQKggIF7NRSL0y3NepI6c2F3VdXDbEEKo" 
		s .= "IHCZ11S375PrTKWtMj7CgCOAu5YoK07y3L9P7gUv9gLqP/MQgNupMC4P0t/a0BMytQVwZ5siAFHfqr7k/13k4yXEnUClUUV7CkCt" 
		s .= "bgqLkmBfyfLfiUMLmYD4gfvGU+uWCYUEsE11LhcFiCLn3SSBqIktVwuIkqBi+vJZAKXOKI92lVuuo+9lZEXrvXwGANYCcC7zqw4T" 
		s .= "sNNVEspWilsuHBaAdOBN9zAFq0waKckDkE496p4kbWVmGiePEwcwdzvXbi9Cg+YkGRsW/zdft6Hz+1MHMN8IwSilzGCOqlx/QO+3" 
		s .= "pw8g8KzQ/wD+B/A/gP8BJA+gtPrnJQCmRMByaV1Akn4iNBTWa3hDItQiewA6Q+jJA2AOAHSOYaRvApW9/NUr9gL76Ntc5dD8l/5J" 
		s .= "1AyigKhzVdoLFfrncDLIAw7TfxpLZ6Iwr0SIGhHY5KfvyQQ3N6AmsJeP6YtS4T0ZUinBfpfmUL5qL7CfArhBsL99nZMFDwPodeP0" 
		s .= "OR9C1wzka4UNrqIo3QblogE4nB1YL5Ve5W6WV3+8VNnkKpo10coAwNdhmDktQn0/CT5fto8MD9ZkCeBwHOpmA2R4Ed1TAJAFgKXk" 
		s .= "rxB/ML2E1XFKyhaAckbqd+On+skAVRaXkDpOSUUHsHQ9mvN1+uu16zZFwWcArPdm2B4aL8V1+utar9u3LbAOGQJYKSwfFlg/M+Dy" 
		s .= "e9zGpa0BPHdq/Hqj8T+AyABI7KuDQk1KOgGIfXVMiElJNw1IAwB6AkCtX7MKvdzm5CwB4PQA4LgA+lSaw64FAVsAdnuhMACcdkOW" 
		s .= "ABy2Av4BOG0GLAF4yIT9LaczI7YAitQAFHEBkHQyYcdUELLPgxxTQXDIg4o2EQAumRBknwY4JgKQfRrgmAhA9mmAYyIA+UdBtzgI" 
		s .= "+UdBtzjoAoClAoDFBpDSZth1Qwz5R0G3OAj5BwG3MAAvCAJOYQBeEAScwoAVgDqlrZDjbgBe4AOdvCBY+8AqJQD2F2rZAEiqHuZa" 
		s .= "FYM3+EAXLwjWeSBLCYD91ZrwBh/o4gXB1gU0aQFobJ0AvMIFODgBsEyDCpYWAH7HdAwAKboABycAr3ABDhtCsHQBAyS2bC/ZBjsX" 
		s .= "gFhqAGydgDEA9UcVHnUCY3AAXZGiC9hOJ3bBASSZBThkAmAXBNv0AFhesg1WQbAq0wNg+bUJeEcQtA+EYGUBLEUAzMoG4C0WYGsD" 
		s .= "8BoLsIwD8JIYYB0HwCILStMCtss6unAAND6uloANmFUGjQCkug842oDZnhhe4wLnZfHxNTB2gQVLF4BFKgDGCpCqC5TcIAkDAPdp" 
		s .= "u8DdDZqoABjHwDZlAK1xNgimMXCApNdgGgnBUAFQmzYAfmlT5x9AHgpgrgLwLgUwVwEwCwHJK4BQAe1AAEY5QPoKsKkA8QtgROnn" 
		s .= "AMdcQHdHACYeMOUk8DsX0PWDYOIBs1CA7e66zh+AsU+0HXTbJNJrk4G+AeTgAQ9+UMsIQDsCFBSyWdwIiB8AHcrKACQj0NkUga4D" 
		s .= "yMcAJCPQcAOgKz+FrBY3AjUB0HOA2jd8JpYRqx0h6Mnf5Cb/dp2zisA9gLEuMigD3SaERT3aAxDyIwYZLnGh9T2BOwBdn7P8+0W2" 
		s .= "fWcFAHcob/ml7x122BzAWGcvv3SZ8W8zgF+vv9f+QF7SBMRlxv0vJYBr8WubK45TjgWTElwjgEvxtytPm8zln+8y3q5tvUTwDQCP" 
		s .= "pLe84jnVjEi60LonI/4NAOOxIz0y+T5mZo5gfqc96UaMjwBw1xFC6rp3uOE7GyVYINT1JHC3WATM+z10dcV1Cy9a7dWV3mjeK4I4" 
		s .= "A/Rm8WcluERAZgC4P93wTd8m/qoF9HSreY9PAFA1sBJeukr2dbP7CqDePnzRDLR9rfTCFNZLzbdiAWyzX63jFd8ZQSgnj7D7gHTP" 
		s .= "gISvGa4AukyLXh4KZt2aB6B8Gp/eAGxXTsBW+GT/EgC2FUxha301/xKAfaYWMpiBDpAT7VPVkMEUfKgYUGOxHSavKP6YKsDaPAap" 
		s .= "/zf8KwAGqW8I8gzYPxIImDxHBjnOwHhJgnjjGA5DIMM/QEB8/ZIcaoJiDOgfiAT02DWG4yTk+wnw8uA2RwnHUcACvdsKSlEY605l" 
		s .= "cUxe0wu5SQBEl4Tgc19AVIaKir5UCUoqGgTS2Aicx2EK1LyxLFiyrUlW4+vWGN4L5I2qMFyKtXwt7xFe0p9XP0JLt++dosPY0KE3" 
		s .= "iIn05d+Jwfm3zn+SUTo0TbV9LhFVywcTy7ivc65tHh5heob5q41XzyxJPwlG8E1ztDvUyBfBGJu/ithO/77IPf+56++FNrHaCZM8" 
		s .= "P75buny9Unrolq1ftpQr4d19d3gk6PxL16XzxdzwejCZss6T/HhoREbVfAAea40/8BQCPfF/Pt3FpIxiQMLij4RsK7aDk/g6AxLC" 
		s .= "G94zQPNXs/t6Xf2XpoUbK5Ab/etD8Ceo+/7eRhGqyWg0JTYPS8yiFec/OXfWR7yuv+mfebCik4kFmiyQu/yTPPwh/viTLA9Bzq9j" 
		s .= "/sJ5PQ9FWAxK8l9KFsrLRMF4GK74nq3ZGYTYUIht7JzIkJ/yTM+8PjTZHxpjv3eK3oxWbgj8E9jkR3WncxTG98FJXQQijxhCyd93" 
		s .= "2Osj+wXwt89Xey4s0EJr9vt5APu4gddYwFChfxDsWQBbYcFnm0HMexLv8gcAsBHwV2MWpz8CyB8CwGYF3tzAEEr/AwHwfdJO/wxc" 
		s .= "KgD8nrU0OAWZDACvp235fQDkLycAHs9bm5yETgeAaLS4G4EwgPEvLwDejMDgIHhaALCfSCAiAM4OgJ9IEDYChAXgo+X+1crODIAY" 
		s .= "Q7cnIOQPZwBBAQgjKGwrZEL+gAYQFsBGwE4HNvnJX64Atpa7Tan8opWdHQCp5W5cHmEXrez8AEgtdzNHsE1yhJY/NACp5V7pN83K" 
		s .= "vQMSWv7gAKSWu27fUJpkQMHlDw9AHjpAGh10uZePSHD5IwA4DB0gdHsqr2Ry97PvIjxcDACHyZPJGVwfzpsPtFWKXnauAOQbGcTx" 
		s .= "TMrYNt8zT918D3347gA9C0BumknmUE1rnfT5/v96giM9WCwAJnMXqO/GaI8VD8DSQUe2gxyvADCXCru6R4pBDhz1iSIDWEzhPHrC" 
		s .= "73UgEVX/OQB//LaSWmBY524UgxzB1n8YHLKEmainPgAAAABJRU5ErkJggg==" 

	 	GUI, New
		GUI, 2: +AlwaysOnTop


		GUI, 2: Add, Text, x8 y280 w408 h16, Email:  Casper.Harkin@gmail.com
		GUI, 2: Add, Edit, x288 y144 w128 h20 +ReadOnly, % 0.1
		GUI, 2: Add, Edit, x288 y176 w128 h20 +ReadOnly, % "22/2/20221"
		GUI, 2: Add, Text, x184 y144 w88 h16, Build Version:
		GUI, 2: Font, Bold s20

		Gui, 2: Add, Text, x22 y10 w440  +Center, Compare and Contrast
		Gui, 2: Add, Text, x22 y40 w440  +Center, by Casper Harkin

		GUI, 2: Font
		GUI, 2: Add, Button, x424 y264 w64 h28 +HwndhClose, Close
			This.Bind(hClose, "GuiDestroy")

		GUI, 2: Add, Text, x184 y176 w88 h16, Created:
		GUI, 2: Add, Text, x8 y264 w408 h16, Phone: (00) 0000 0000
		GUI, 2: Add, Text, x8 y248 w408 h16, For assistance please contact :
		GUI, 2: Add, Picture, x35 y121 w96 h96, % "HBITMAP:" h := This.GDI_Startup(s)
		GUI, 2: Show, w496 h303, About Compare And Contrast:
	}

	GDI_Startup(Base64) {
		hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll") ; Load module
		VarSetCapacity(GdiplusStartupInput, (A_PtrSize = 8 ? 24 : 16), 0) ; GdiplusStartupInput structure
		NumPut(1, GdiplusStartupInput, 0, "UInt") ; GdiplusVersion
		VarSetCapacity(pToken, 0) ;Make var big enough to fit contents
		DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &GdiplusStartupInput, "Ptr", 0) ; Initialize GDI+
		BMP := This.GdipCreateFromBase(Base64) ;Turn the raw data into an image. 
		DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip) ; Free GDI+ module from memory
		Return BMP 
	}

	GdipCreateFromBase(B64, IsIcon := 0) {
		VarSetCapacity(B64Len, 0)
		DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", StrLen(B64), "UInt", 0x01, "Ptr", 0, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
		VarSetCapacity(B64Dec, B64Len, 0) ; pbBinary size
		DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", StrLen(B64), "UInt", 0x01, "Ptr", &B64Dec, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
		pStream := DllCall("Shlwapi.dll\SHCreateMemStream", "Ptr", &B64Dec, "UInt", B64Len, "UPtr")
		VarSetCapacity(pBitmap, 0)
		DllCall("Gdiplus.dll\GdipCreateBitmapFromStreamICM", "Ptr", pStream, "PtrP", pBitmap)
		VarSetCapacity(hBitmap, 0)
		DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "UInt", pBitmap, "UInt*", hBitmap, "Int", 0XFFFFFFFF)
		If (IsIcon) 
		    DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hIcon, "UInt", 0)
		ObjRelease(pStream)
		return (IsIcon ? hIcon : hBitmap)
	}
}
