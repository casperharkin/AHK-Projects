#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


Report = 
(
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                          Errors            
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
)

If !Check(Obj := {Review:"     873 95008", Against:"87395008", Subject:"AGS", DataType : "AGS"}){
	Report .= "`n" Obj.Subject " is Wrong"
}

If !Check(Obj := {Review:"    $1,337.56", Against:"$1337.5560000 ", Subject:"Total Amount", DataType:"Figure", Format:"", Float:""}){
	Report .= "`n" Obj.Subject " is Wrong"
}


If !Check(Obj := {Review:"8/1/2023 ", Against:"08/01/2023", Subject:"FBT Date", DataType:"Date"}){
	Report .= "`n" Obj.Subject " is Wrong"
}

If !Check(Obj := {Review:"po box 42378 ", Against:"PO BOX 42378", Subject:"Postal Address", DataType:"Text"}){
	Report .= "`n" Obj.Subject " is Wrong"
}

If Report
	MsgBox % Report 


Exit ; EOAES



Check(Obj){
	Return CnC.Compare(Obj["Review"], Obj["Against"], Obj["Subject"], Obj["DataType"] , Obj["Format"] , Obj["Float"])
}

Class CnC { 


/* 

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ####   Compare_And_Contrast - by Casper Harkin   ####
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Aim:
I have a need to compare / check data against different sources. 

Issue:
The issue is that each source can have different formatting, rounding, spacing etc for the data and simple comparisons can return false negatives. 

Examples of False Negatives I Wish to Avoid: 
•	Source A displays the data as $1,337.00, Source B displays the data as $1337 and Source C displays the data as 1337.00 (Formatting)
•	Source A displays the data as "Handcock Street  ", Source B displays the data as " Handcock  Street" (Spacing)
•	Source A displays the data as $1337.1001, Source B displays the data as $1337.10 (Rounding)
•	Source A displays the data as 01/01/2021, Source B displays the data as 1/1/2021 (Leading Zeros)


The solution was to try and write a class that automatically sanitise the data and then compares it.

A GUI will be created displaying the data with options to compare with different formatting, with the ability
to update the data, or manually confirm result. If the data looks correct, paint the text Green else mark it Red for Review.  



	TODO:

	Compare Objects
	Build Out Regex Patterns  
	Build system for editing input via List View <- Needed?  


	MAYBE FEATURES:

	
	



*/

	Compare(Compare, Against, Context := "", DataType := "", Format := "", Float := ""){

		if !Compare or !Against {
			MsgBox % "Error - Review or Against Empty`n`nReview: " Compare "`nAgainst: " Against
			Return False
		}


		This.Properties := {}

		This.Properties["OG_Compare"] := This.Properties["Compare"] := Compare, This.Properties["Context"] := Context 
		This.Properties["OG_Against"] := This.Properties["Against"] := Against, This.Properties["Format"] := Format


		if (DataType = "AGS"){
			If InStr(This.Properties["Compare"], "-")
				This.Properties["Compare"] := StrReplace(Compare, "-")
			else If InStr(This.Properties["Compare"], " ")
				This.Properties["Compare"] := StrReplace(Compare, A_Space)
		}

		if (DataType = "Figure"){
			If InStr(This.Properties["Compare"], "$"){

				This.Properties["Compare"] := Round(StrReplace(StrReplace(Compare, "$"), ","), 2)
				This.Properties["Against"] := Round(StrReplace(StrReplace(Against, "$"), ","), 2)

			}
		}

		if (DataType = "Date"){

			CompareDate := StrSplit(This.Properties["Compare"], "/")
			AgainstDate := StrSplit(This.Properties["Against"], "/")

			for each, part in CompareDate
				if (part = AgainstDate[each])
					This.Properties["Compare"] := This.Properties["Against"] := Against
		}

		if (DataType = "Text"){
			This.Properties["Compare"] := Format("{:U}",This.Properties["Compare"])
			This.Properties["Against"] := Format("{:U}",This.Properties["Against"])
		}

		if This.Properties["Format"]["RegExReplace"] {
			This.Properties["Compare"] := RegExReplace(This.Properties["Compare"], This.Properties["Format"]["RegExReplace"])
			This.Properties["Against"] := RegExReplace(This.Properties["Against"], This.Properties["Format"]["RegExReplace"])
		}

		if (This.Properties["Format"]["Trim"] != "False") {
			This.Properties["Compare"] := Trim(This.Properties["Compare"])
			This.Properties["Against"] := Trim(This.Properties["Against"])
		}

		This.DirectCompare()

		Return This.Properties["Response"]
	}

	DirectCompare(){

		if (This.Properties["Compare"] = This.Properties["Against"]){
			This.Properties["Response"] := "True"
			This.GuiCompare("Looks Good")
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
	}

	GuiCompare(Passed := ""){

		CnCMenuHandler := ObjBindMethod(This, "CnCMenuHandler")
			
		Gui, New

		Menu, FileMenu, Add, &Open`tCtrl+O, % CnCMenuHandler  
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
			This.Properties["Context"] := "Compare and Contrast"


		Gui, Add, Text, x0 y9 w275 h30 +Center r2 +HwndhContextText cBold, % This.Properties["Context"]

		Gui, Add, ListView, x12 y39 w250 h100 +HWNDhListView cRed -ReadOnly +AltSubmit -Multi , Source|Formatted|Original

		ListView := New LV_InCellEdit(hListView)

		LV_Add("", "Review: ", This.Properties["Compare"], This.Properties["OG_Compare"])
		LV_Add("", "Against: ", This.Properties["Against"], This.Properties["OG_Against"])
			LV_ModifyCol() 
			LV_ModifyCol(1, "AutoHdr") 
			LV_ModifyCol(2, "AutoHdr") 

		Gui, Add, Button, x12 y145 w90 h30 +HwndhYes, Yes
			This.Bind(hYes, "GuiResponse", "True")

		Gui, Add, Button, x162 y145 w90 h30 +HwndhNo, No
			This.Bind(hNo, "GuiResponse", False)

		if Passed {
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
		Gui, Destroy
	}

	CheckInString(String, ArrayOfStrings){
		for e, i in ArrayOfStrings
			If InStr(String, i)
				Return "True"
		Return False
	}

	Bind(Hwnd, Method, Params*){
		BoundFunc := ObjBindMethod(This, Method, Params*)
		GuiControl +g, % Hwnd, % BoundFunc
   	}

	Float(n, p := 6 ){ 	; By SKAN on D1BM @ goo.gl/Q7zQG9
		Return SubStr(n:=Format("{:0." p "f}",n),1,-1-p) . ((n:=RTrim(SubStr(n,1-p),0) ) ? "." . n : "") 
	}

}
