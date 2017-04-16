#SingleInstance, Force
#NoEnv
AutoTrim, On
SetWorkingDir %A_ScriptDir%\
iniRead, Client_List, %A_ScriptDir%\Settings.ini, Clients
iniRead, Supervisor_List, %A_ScriptDir%\Settings.ini, Supervisors
iniRead, Operative_List, %A_ScriptDir%\Settings.ini, Operatives
iniRead, Sender_List, %A_ScriptDir%\Settings.ini, Administrators
iniRead, Vehicle_List, %A_ScriptDir%\Settings.ini, Vehicles

xlsx := ".xlsx"
docx := ".docx"

	
;GUI

Gui, Add, Text, x25 y15 , Blitzd_GUI
Gui, Show, h700 w475, Blitzd_GUI 0.5
Gui, Add, Text, x25 y50 , Client:
Gui, Add, DropDownList, x25 w200 vClient gChooseClient AltSubmit, %Client_List%
Gui, Add, Text, x250 y50 , Sender:
Gui, Add, DropDownList, w200 Choose1 vSender gChooseSender AltSubmit, %Sender_List%
GuiControlGet, Sender,,, Text

	iniRead, ChosenSender, %A_ScriptDir%\Settings.ini, %Sender%
;	MsgBox, %ChosenSender%
	iniRead, ChosenSender_Name, %A_ScriptDir%\Settings.ini, %Sender%, Name
;	MsgBox, %ChosenSender%
	iniRead, ChosenSender_Email, %A_ScriptDir%\Settings.ini, %Sender%, Email
;	MsgBox, %ChosenSender_Email%
	iniRead, ChosenSender_Phone, %A_ScriptDir%\Settings.ini, %Sender%, Phone
;	MsgBox, %ChosenSender_Phone%
	Gui, Submit, NoHide

Gui, Add, Text, x25 y100 , Location:
Gui, Add, DropDownList, w0 vLocation gChooseLocation AltSubmit, %Client_%job_data%_List%
Gui, Add, Text, x250 y100 , Service:
Gui, Add, DropDownList, w0 vService gChooseService AltSubmit
Gui, Add, Text, x25 y150 , Contact:
GUI, Add, Edit, w0 vContact, %contact_check%
Gui, Add, Text, x250 y150 , Email:
GUI, Add, Edit, w0 vEmail, %email_check%
Gui, Add, Text, x25 y200 , Price:
GUI, Add, Edit, w0 vPrice, %price_check%
Gui, Add, Text, x250 y200 , Quote:
GUI, Add, Edit, w0 vQuote, %quote_check%
Gui, Add, Text, x25 y275 , Start Date:
Gui, Add, DateTime, w0 vChooseStartDate
;Gui, Add, Button,, StartDate
Gui, Show
Gui, Add, Text, x250 y275 , End Date:
Gui, Add, DateTime, w0 vChooseEndDate
;Gui, Add, Button,, EndDate
Gui, Show
Gui, Add, Text, x25 y325, StartTime:
GUI, Add, Edit, w0 vStartTime, %timestart_check%
Gui, Add, Text, x250 y325, EndTime:
GUI, Add, Edit, w0 vEndTime, %timeend_check%
Gui, Add, Text, x25 y375 , Supervisor:
Gui, Add, DropDownList, w0 Choose1 vSupervisor gChooseSupervisor AltSubmit, %Supervisor_List%

    GuiControlGet, Supervisor,,, Text
	iniRead, ChosenSupervisor, %A_ScriptDir%\Settings.ini, %Supervisor%
;	MsgBox, %ChosenSupervisor%
	iniRead, ChosenSupervisor_Name, %A_ScriptDir%\Settings.ini, %Supervisor%, Name
;	MsgBox, %ChosenSupervisor_Name%
	iniRead, ChosenSupervisor_Email, %A_ScriptDir%\Settings.ini, %Supervisor%, Email
;	MsgBox, %ChosenSupervisor_Email%
	iniRead, ChosenSupervisor_Phone, %A_ScriptDir%\Settings.ini, %Supervisor%, Phone
;	MsgBox, %ChosenSupervisor_Phone%
	Gui, Submit, NoHide

Gui, Add, Text, x250 y375 , Vehicle:
Gui, Add, DropDownList, w0 vExtra gChooseExtra AltSubmit, %Vehicle_List%
Gui, Add, Text, x25 y425 , Op1:
Gui, Add, DropDownList, w0 vOp1 gChooseOp1 AltSubmit, %Operative_List%
Gui, Add, Text, x250 y425 , Op2:
Gui, Add, DropDownList, w0 vOp2 gChooseOp2 AltSubmit, %Operative_List%
Gui, Add, Text, x25 y475 , Op3:
Gui, Add, DropDownList, w0 vOp3 gChooseOp3 AltSubmit, %Operative_List%
Gui, Add, Text, x250 y475 , Op4:
Gui, Add, DropDownList, w0 vOp4 gChooseOp4 AltSubmit, %Operative_List%
Gui, Add, Text, x25 y525 , Op5:
Gui, Add, DropDownList, w0 vOp5 gChooseOp5 AltSubmit, %Operative_List%
Gui, Add, Text, x250 y525 , Op6:
Gui, Add, DropDownList, w0 vOp6 gChooseOp6 AltSubmit, %Operative_List%
Gui, Add, Text, x25 y575 , Op7:
Gui, Add, DropDownList, w0 vOp7 gChooseOp7 AltSubmit, %Operative_List%
Gui, Add, Text, x250 y575 , Op8:
Gui, Add, DropDownList, w0 vOp8 gChooseOp8 AltSubmit, %Operative_List%


Gui, Add, Button,, Submit



Return


;GUI_Choose

;What happens when you:
ChooseClient:
    GuiControlGet, Client,,, Text
	ClientBook := ComObjGet(A_ScriptDir "\Clients\" Client ".xlsx")
	
    GuiControlGet, Client
	Client%Client%_List := ClientBook.Sheets("Info").Cells(2, 2).Value
	
	%Site%_List := ClientBook.Sheets("Info").Cells(2, 2).Value
	
	GuiControl,, Location, % "|" %Site%_List
	GuiControl, Move, Location, x25 w200
	GuiControl, Move, Service, w0
	GuiControl, Move, Contact, w0
	GuiControl, Move, Email, w0
	GuiControl, Move, Price, w0
	GuiControl, Move, Quote, w0
	GuiControl, Move, ChooseStartDate, w0
	GuiControl, Move, ChooseEndDate, w0
	GuiControl, Move, StartTime, w0
	GuiControl, Move, Endtime, w0
	GuiControl, Move, Supervisor, w0	
	GuiControl, Move, Op1, w0
	GuiControl, Move, Op2, w0
	GuiControl, Move, Op3, w0
	GuiControl, Move, Op4, w0
	GuiControl, Move, Op5, w0
	GuiControl, Move, Op6, w0
	GuiControl, Move, Op7, w0
	GuiControl, Move, Op8, w0

    GuiControlGet, Client,,, Text
	Return	
	
	
;What happens when you:	
ChooseLocation:	

	GuiControlGet, Location,,, Text
	GuiControl, Move, Service, w0
	GuiControl, Move, Contact, w0
	GuiControl, Move, Email, w0
	GuiControl, Move, Price, w0
	GuiControl, Move, ChooseStartDate, w0
	GuiControl, Move, ChooseEndDate, w0
	GuiControl, Move, StartTime, w0
	GuiControl, Move, Endtime, w0
	GuiControl, Move, Supervisor, w0	
	GuiControl, Move, Op1, w0
	GuiControl, Move, Op2, w0
	GuiControl, Move, Op3, w0
	GuiControl, Move, Op4, w0
	GuiControl, Move, Op5, w0
	GuiControl, Move, Op6, w0
	GuiControl, Move, Op7, w0
	GuiControl, Move, Op8, w0

	Location_Name := Location
	LookForServices := Location
;	MsgBox, %Location%
	client_check := ClientBook.Sheets("Info").Range("B5:B9").Find(LookForServices).Offset(0,1).Text
;	MsgBox, %client_check%
	%Site%_List := % client_check
;	Client_%Client_1%2_List := % client_check

;	%Location%_List := ClientBook.Sheets("Info").Cells(12, 2).Value	

	GuiControlGet, Location
	GuiControl,, Service, % "|" %Site%_List
	GuiControl, Move, Service, x250 w200
	Return


;What happens when you:	
ChooseService:	

    GuiControlGet, Service ,,, text
	Location_Choice := Service	
	
	StringReplace , Location_Choice, Location_Choice, %A_Space%,,All
	
	Location_Number := Location_Choice
	
;	MsgBox, %Location_Number%
	
	JobSheet := Location_Name "-" Location_Number
	
	StringReplace , JobSheet, JobSheet, %A_Space%,,All
	
	LookForContact := "Contact:"
	contact_check := ClientBook.Sheets(JobSheet).Range("A1:A10").Find(LookForContact).Offset(0,1).Text

	LookForEmail := "Email:"
	email_check := ClientBook.Sheets(JobSheet).Range("A1:A10").Find(LookForEmail).Offset(0,1).Text
	
	LookForPrice := "Price:"
	price_check := ClientBook.Sheets(JobSheet).Range("A1:A10").Find(LookForPrice).Offset(0,1).Text

	LookForStart := "Start:"
	timestart_check := ClientBook.Sheets(JobSheet).Range("A1:A10").Find(LookForStart).Offset(0,1).Text

	LookForEnd := "End:"
	timeend_check := ClientBook.Sheets(JobSheet).Range("A1:A10").Find(LookForEnd).Offset(0,1).Text

	ChooseHours := timestart_check " – " timeend_check
	
;	MsgBox, %price_check%
;	MsgBox, %JobSheet%
;	MsgBox, %Service%
;	MsgBox, %Job_Number%


	GuiControl, Move, Contact, x25 w200
	GuiControl,, Contact, %contact_check% 
	GuiControl, Move, Email, x250 w200
	GuiControl,, Email, %email_check% 
	GuiControl, Move, Price, x25 w200
	GuiControl,, Price, %price_check%
	GuiControl, Move, Quote, x250 w200
	GuiControl,, Quote, %quote_check% 
	GuiControl, Move, ChooseStartDate, x25 w200
	GuiControl, Move, ChooseEndDate, x250 w200
	GuiControl,, StartTime, %timestart_check% 
	GuiControl,, EndTime, %timeend_check% 
	GuiControl, Move, StartTime, x25 w200
	GuiControl, Move, EndTime, x250 w200
	GuiControl, Move, Supervisor, x25 w200
	GuiControl, Move, Extra, x250 w200
	GuiControl, Move, Op1, x25 w200
	GuiControl, Move, Op2, x250 w200
	GuiControl, Move, Op3, x25 w200
	GuiControl, Move, Op4, x250 w200
	GuiControl, Move, Op5, x25 w200
	GuiControl, Move, Op6, x250 w200
	GuiControl, Move, Op7, x25 w200
	GuiControl, Move, Op8, x250 w200


;	MsgBox, %ChooseHours%
	Return
	
ChooseSupervisor:
    GuiControlGet, Supervisor,,, Text
	iniRead, ChosenSupervisor, %A_ScriptDir%\Settings.ini, %Supervisor%
;	MsgBox, %ChosenSupervisor%
	iniRead, ChosenSupervisor_Name, %A_ScriptDir%\Settings.ini, %Supervisor%, Name
;	MsgBox, %ChosenSupervisor_Name%
	iniRead, ChosenSupervisor_Email, %A_ScriptDir%\Settings.ini, %Supervisor%, Email
;	MsgBox, %ChosenSupervisor_Email%
	iniRead, ChosenSupervisor_Phone, %A_ScriptDir%\Settings.ini, %Supervisor%, Phone
;	MsgBox, %ChosenSupervisor_Phone%
	Gui, Submit, NoHide
Return

ChooseSender:
    GuiControlGet, Sender,,, Text
	iniRead, ChosenSender, %A_ScriptDir%\Settings.ini, %Sender%
;	MsgBox, %ChosenSender%
	iniRead, ChosenSender_Name, %A_ScriptDir%\Settings.ini, %Sender%, Name
;	MsgBox, %ChosenSender%
	iniRead, ChosenSender_Email, %A_ScriptDir%\Settings.ini, %Sender%, Email
;	MsgBox, %ChosenSender_Email%
	iniRead, ChosenSender_Phone, %A_ScriptDir%\Settings.ini, %Sender%, Phone
;	MsgBox, %ChosenSender_Phone%
	Gui, Submit, NoHide
Return

ChooseExtra:
    GuiControlGet, Extra,,, Text
	iniRead, ChosenExtra, %A_ScriptDir%\Settings.ini, %Extra%
	iniRead, ChosenExtra_Vehicle, %A_ScriptDir%\Settings.ini, %Extra%, Vehicle
;	MsgBox, %ChosenExtra_Vehicle%
	iniRead, ChosenExtra_Reg, %A_ScriptDir%\Settings.ini, %Extra%, Reg
;	MsgBox, %ChosenExtra_Reg%
;	MsgBox, %ChosenExtra%
	Gui, Submit, NoHide
Return

StartDate:
	GuiControlGet, StartDate
	Gui, Submit, NoHide	
Return	


EndDate:
	GuiControlGet, EndDate
	Gui, Submit, NoHide
;	MsgBox, % EndDate

Return	
	

ButtonStartDate:
	GuiControlGet, StartDate
	Gui, Submit, NoHide
;	MsgBox, % ChooseStartDate
	
Return	
	
ButtonEndDate:
	GuiControlGet, EndDate
	Gui, Submit, NoHide
;	MsgBox, % ChooseEndDate
	
Return	

ChooseOp1:
	GuiControlGet, Op1
	Gui, Submit, NoHide
Return		

ChooseOp2:
	GuiControlGet, Op2
	Gui, Submit, NoHide
Return		

ChooseOp3:
	GuiControlGet, Op3
	Gui, Submit, NoHide
Return		

ChooseOp4:
	GuiControlGet, Op4
	Gui, Submit, NoHide
Return		

ChooseOp5:
	GuiControlGet, Op5
	Gui, Submit, NoHide
Return		

ChooseOp6:
	GuiControlGet, Op6
	Gui, Submit, NoHide
Return		

ChooseOp7:
	GuiControlGet, Op7
	Gui, Submit, NoHide
Return		

ChooseOp8:
	GuiControlGet, Op8
	Gui, Submit, NoHide
Return		


	
	
;SubmitButton	

ButtonSubmit:
Gui, Submit, NoHide 	

 



; .dotx is a Word template, but .doc and .docx will also work
TemplateFilePath := A_ScriptDir "\Templates\Template.docx"
wordApp := ComObjCreate("Word.Application") ; Create an instance of Word

; You can remove this after testing so that Word stays invisible. 
; Confirm that the script closes Word at the end so you don't get a bunch of
; invisible Word applications open in the background.
WordApp.Visible := true

NewDoc := WordApp.Documents.Add(TemplateFilePath) ; Open the template

	FormatTime, ChooseStartDate,%ChooseStartDate%, dddd dd MMMM yyyy
	FormatTime, ChooseEndDate,%ChooseEndDate%, dddd dd MMMM yyyy
    GuiControlGet, Client,,, Text
    GuiControlGet, Location,,, Text
	GuiControlGet, Service,,, Text	
	GuiControlGet, Supervisor,,, Text
	GuiControlGet, Sender,,, Text
	GuiControlGet, Op1,,, Text			
	GuiControlGet, Op2,,, Text			
	GuiControlGet, Op3,,, Text			
	GuiControlGet, Op4,,, Text			
	GuiControlGet, Op5,,, Text			
	GuiControlGet, Op6,,, Text			
	GuiControlGet, Op7,,, Text			
	GuiControlGet, Op8,,, Text
	GuiControlGet, Extra,,, Text		
	
QuoteNumber := "AM 0000"	
	
;QuoteData := ComObjGet(A_ScriptDir "\Clients\Quotes\" QuoteNumber docx)	
;MsgBox, % QuoteData.Sections(1).Range.Text
;Method := % QuoteData.Sections(1).Range.FormattedText
;MsgBox, % QuoteData.Sections(1).Range.Text
;MsgBox, % Word.Sections(1).Footers(1).Range.Text
;QuoteData.Close()	
	
	
	
if !ChosenExtra 
{ChosenExtra=/
}

if !Op1 
{Op1=-
}

if !Op2 
{Op2=-
}

if !Op3 
{Op3=-
}

if !Op4 
{Op4=-
}

if !Op5 
{Op5=-
}

if !Op6 
{Op6=-
}

if !Op7 
{Op7=-
}

if !Op8 
{Op8=-
}


ChosenExtra := ChosenExtra_Vehicle "`n" ChosenExtra_Reg

	
; Put the text into the bookmarks
NewDoc.Bookmarks("FooterName").Range.Text := Location
NewDoc.Bookmarks("FooterNameRAMS").Range.Text := Location
NewDoc.Bookmarks("Contact").Range.Text := contact_check
NewDoc.Bookmarks("FromName").Range.Text := ChosenSender_Name
NewDoc.Bookmarks("FromEmail").Range.Text := ChosenSender_Email
NewDoc.Bookmarks("FromPhone").Range.Text := ChosenSender_Phone
NewDoc.Bookmarks("Email").Range.Text := email_check
NewDoc.Bookmarks("Client").Range.Text := Client
NewDoc.Bookmarks("Location").Range.Text := Location
NewDoc.Bookmarks("Service").Range.Text := Service
NewDoc.Bookmarks("From").Range.Text := ChooseStartDate
NewDoc.Bookmarks("Until").Range.Text := ChooseEndDate
NewDoc.Bookmarks("Hours").Range.Text := ChooseHours
NewDoc.Bookmarks("Supervisor").Range.Text := ChosenSupervisor_Name
NewDoc.Bookmarks("SupervisorEmail").Range.Text := ChosenSupervisor_Email
NewDoc.Bookmarks("SupervisorPhone").Range.Text := ChosenSupervisor_Phone
NewDoc.Bookmarks("Op1").Range.Text := Op1
NewDoc.Bookmarks("Op2").Range.Text := Op2
NewDoc.Bookmarks("Op3").Range.Text := Op3
NewDoc.Bookmarks("Op4").Range.Text := Op4
NewDoc.Bookmarks("Op5").Range.Text := Op5
NewDoc.Bookmarks("Op6").Range.Text := Op6
NewDoc.Bookmarks("Op7").Range.Text := Op7
NewDoc.Bookmarks("Op8").Range.Text := Op8
NewDoc.Bookmarks("Extra").Range.Text := ChosenExtra
NewDoc.Bookmarks("RAMSLocation").Range.Text := Location
NewDoc.Bookmarks("RAMSService").Range.Text := Service

QuoteData := ComObjGet(A_ScriptDir "\Clients\Quotes\" QuoteNumber docx)
;QuoteData.Bookmarks(Method).Range.FormattedText.copy()
QuoteData.Sections(1).Range.FormattedText.copy()
WordApp.ActiveDocument.Bookmarks("Method").select() 
WordApp.Selection.PasteAndFormat(0)
QuoteData.Close()

return
*/


esc::exitapp
f12::reload