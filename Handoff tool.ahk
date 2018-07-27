#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines, -1 ;AHK run faster


Input_L10n := "Assign to Linguist/LM (for approvals), then to dl-pp-triage-dt-g11n-l10n-production."
Input_Labels := "AlphaCRC, LQA, "
Input_AffectVersion := "Unknown – Not Yet Live|Unknown – On Live|N/A"
Input_Project := "eCAT/XPT - Content|eCAT Tool/MRA|PP - BL-Credit|PP - BL-Global Operations|PP - BL-Risk|PP - PL-Consumer|PP - PL-Merchant|PP - PL-Payments & Platform"
Input_TestCase := ""
Input_To := "DL-PayPal-LQATesters@paypal.com"
Input_CC := "DL-PP-LQA-LEADS@paypal.com"
;Input_Locales := "ar_AE|ar_BH|ar_DZ|ar_EG|ar_JO|ar_KW|ar_MA|ar_OM|ar_QA|ar_SA|ar_TN|ar_YE|da_DK|de_AT|de_CH|de_DE|de_LU|en_AR|en_AT|en_AU|en_BE|en_BR|en_C2|en_CA|en_CH|en_CL|en_CR|en_CZ|en_DE|en_DK|en_DO|en_EC|en_ES|en_FI|en_FR|en_GB|en_GF|en_GP|en_GR|en_HK|en_HU|en_ID|en_IE|en_IL|en_IN|en_IS|en_IT|en_JM|en_JP|en_KR|en_LU|en_MC|en_MQ|en_MX|en_MY|en_NL|en_NO|en_NZ|en_PL|en_PT|en_RE|en_RU|en_SE|en_SG|en_TH|en_TR|en_TW|en_US|en_UY|en_VE|es_AR|es_CL|es_CR|es_CZ|es_DK|es_DO|es_EC|es_ES|es_FI|es_GF|es_GP|es_GR|es_HU|es_IE|es_JM|es_LU|es_MC|es_MQ|es_MX|es_NO|es_NZ|es_PT|es_RE|es_RU|es_SE|es_US|es_UY|es_VE|es_XC|fr_BE|fr_CA|fr_CH|fr_CL|fr_CR|fr_CZ|fr_DK|fr_DO|fr_EC|fr_FI|fr_FR|fr_GF|fr_GP|fr_GR|fr_HU|fr_IE|fr_JM|fr_LU|fr_MC|fr_MQ|fr_NO|fr_NZ|fr_PT|fr_RE|fr_RU|fr_SE|fr_US|fr_UY|fr_VE|fr_XC|he_IL|id_ID|it_IT|ja_JP|ko_KR|nl_BE|nl_NL|no_NO|pl_PL|pt_BR|pt_PT|py_en|py_es|ru_RU|sv_SE|th_TH|tr_TR|zh_C2|zh_CL|zh_CN|zh_CR|zh_CZ|zh_DK|zh_DO|zh_EC|zh_FI|zh_GF|zh_GP|zh_GR|zh_HK|zh_HU|zh_IE|zh_JM|zh_LU|zh_MC|zh_MQ|zh_NO|zh_NZ|zh_PT|zh_RE|zh_RU|zh_SE|zh_TW|zh_US|zh_UY|zh_VE|zh_XC"

;Win+Z key to trigger
/*
#Z::
	Gosub CreateGUI
	Return
*/

CreateGUI:
Gui Destroy
;Left column
Gui Add, Text, x10 y10 w95 h25 +0x200 Right Section, Sender
;Gui Add, Text, yp+50 w95 h25 +0x200 Right Section, Sender
Gui Add, Text, yp+25 w95 h25 +0x200 Right, To...
Gui Add, Text, yp+25 w95 h25 +0x200 Right, CC...
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Email subject
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Internal deadline
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Official deadline
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Automation link
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Manual stage
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Build ID
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Project
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Summary
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Affect version
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Component
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Labels
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Test case details
;Gui Add, Text, yp+25 w95 h25 +0x200 Right, Test case URL
;Gui Add, Text, yp+25 w95 h25 +0x200 Right, Locales
Gui Add, Text, yp+25 w95 h25 +0x200 Right, L10n assignee
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Non-L10n assignee
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Demo taker
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Time tracker
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Slack channel
Gui Add, Text, yp+25 w95 h25 +0x200 Right, Slack URL
; Right column
Gui Add, Edit, vInput_Sender ys x115 w450 h21, %Input_Sender%
Gui Add, Edit, vInput_To yp+25 w450 h21, %Input_To%
Gui Add, Edit, vInput_CC yp+25 w450 h21, %Input_CC%
Gui Add, Edit, vInput_Subject yp+25 w450 h21, %Input_Subject%
Gui Add, Edit, vInput_InternalDL yp+25 w450 h21, %Input_InternalDL%
;Gui, Add, DateTime, yp+25 vInput_InternalDL w450 h21, dd/MM/yy HH:mm
Gui Add, Edit, vInput_OfficialDL yp+25 w450 h21, %Input_OfficialDL%
;Gui, Add, DateTime, yp+25 vInput_OfficialDL w450 h21, dd/MM/yy
Gui Add, Edit, vInput_Automation yp+25 w450 h21, %Input_Automation%
Gui Add, Edit, vInput_Manual yp+25 w450 h21, %Input_Manual%
Gui Add, Edit, vInput_BuildID yp+25 w450 h21, %Input_BuildID%
Gui Add, Combobox, vInput_Project yp+25 w450 +Sort, %Input_Project%
Gui Add, Edit, vInput_Summary yp+25 w450 h21, %Input_Summary%
Gui Add, Combobox, vInput_AffectVersion yp+25 w450, %Input_AffectVersion%
Gui Add, Edit, vInput_Component yp+25 w450 h21, %Input_Component%
Gui Add, Edit, vInput_Labels yp+25 w450 h21, %Input_Labels%
Gui Add, Edit, vInput_TestCase yp+25 w450 h21, %Input_TestCase%
;Gui Add, Edit, vInput_TestCaseURL yp+25 w450 h21, %Input_TestCaseURL%
;Gui Add, Button, yp+25 w100 h25 gButtonLocales, Select
;Gui Add, Edit, vInput_Locales yp+25 w450 h61 Multi, %Input_Locales%
;Gui Add, Listbox, vInput_Locales yp+25 w450 h61 +Multi +Sort, %Input_Locales%
Gui Add, Edit, vInput_L10n yp+25 w450 h21, %Input_L10n%
Gui Add, Edit, vInput_NonL10n yp+25 w450 h21, %Input_NonL10n%
Gui Add, Edit, vInput_DemoTaker yp+25 w450 h21, %Input_DemoTaker%
Gui Add, Edit, vInput_TimeTracker yp+25 w450 h21, %Input_TimeTracker%
Gui Add, Edit, vInput_SlackChan yp+25 w450 h21, %Input_SlackChan%
Gui Add, Edit, vInput_SlackURL yp+25 w450 h21, %Input_SlackURL%
;Buttons
Gui Add, Button, xm+5 w100 h25 gButtonCreate, Create
Gui Add, Button, xp+110 w100 h25 gButtonCancel, Cancel
Gui Add, Button, xp+110 w100 h25 gButton_Load, Load template
Gui Add, Button, xp+110 w100 h25 gButton_Save, Save template
Gui Add, Button, xp+110 w100 h25 gButton_Reset, Reset
Gui Show, AutoSize, Handoff tool
Return

/*
;When pressing Select button (locales)
ButtonLocales:
Gui, 2:Add, Text, x10 y10 w200 h25 +0x200, Select locales (Ctrl for multiple)
Gui, 2:Add, Listbox, vInput_LocalesSelected yp+25 w200 r20 +Multi +Sort, %Input_Locales%
Gui, 2:Add, Button, yp+275 w100 h25 gButtonDoneLocales, Done
Gui, 2:Add, Button, xp+110 w100 h25 gButtonCancelLocales, Cancel
Gui, 2:Show, AutoSize, Handoff tool
Return

;When pressing Done in locales
ButtonDoneLocales:
	Gui, 2:Submit
	GuiControlGet, Input_LocalesSelected
	Msgbox, %Input_LocalesSelected%
	return


;When pressing Cancel in Locales menu
ButtonCancelLocales:
	Gui, 2:Destroy
	return
*/

;When pressing Create button
ButtonCreate:
	Gui, Submit, NoHide
	GuiControlGet, Input_Sender
	GuiControlGet, Input_To
	GuiControlGet, Input_CC
	GuiControlGet, Input_Subject
	GuiControlGet, Input_InternalDL
	GuiControlGet, Input_OfficialDL
	GuiControlGet, Input_Automation
	GuiControlGet, Input_Manual
	GuiControlGet, Input_BuildID
	GuiControlGet, Input_Project
	GuiControlGet, Input_Summary
	GuiControlGet, Input_AffectVersion
	GuiControlGet, Input_Component
	GuiControlGet, Input_Labels
	GuiControlGet, Input_TestCase
	GuiControlGet, Input_TestCaseURL
	GuiControlGet, Input_Locales
	GuiControlGet, Input_L10n
	GuiControlGet, Input_NonL10n
	GuiControlGet, Input_DemoTaker
	GuiControlGet, Input_TimeTracker
	GuiControlGet, Input_SlackChan
	GuiControlGet, Input_SlackURL
	;FormatTime, Input_InternalDL, %Input_InternalDL%, dd/MM/yy '–' HH:mm
	;FormatTime, Input_OfficialDL, %Input_OfficialDL%, dd/MM/yy
	;msgbox, %Input_InternalDL%
;	Loop
;	{
;			Input_Labels := StrReplace(Input_Labels, ",", "`n", Count)
;			if Count = 0  ; No more replacements needed.
;					break
;	}
	Gosub CreateMail
	Return

GuiClose:
    ExitApp

;When pressing Cancel button
ButtonCancel:
	ExitApp

;When pressing Reset button
Button_Reset:
	Reload
	return

;When pressing the Load template Button
Button_Load:
	FileSelectFile, Template_Load, ,, Load template
	;msgbox, %Template_Load%
	if Template_Load !=
	{
	;Email
	IniRead, Input_Sender, %Template_Load%, Email, 1
	IniRead, Input_To, %Template_Load%, Email, 2
	IniRead, Input_CC, %Template_Load%, Email, 3
	; Deadline
	IniRead, Input_InternalDL, %Template_Load%, Deadline, 1
	IniRead, Input_OfficialDL, %Template_Load%, Deadline, 2
	;Stage details
	IniRead, Input_Automation, %Template_Load%, Stage, 1
	IniRead, Input_Manual, %Template_Load%, Stage, 2
	IniRead, Input_BuildID, %Template_Load%, Stage, 3
	;Bug logging guidelines
	IniRead, Input_Summary, %Template_Load%, Bug logging guidelines, 1
	IniRead, Input_AffectVersion, %Template_Load%, Bug logging guidelines, 2
	IniRead, Input_Component, %Template_Load%, Bug logging guidelines, 3
	IniRead, Input_Labels, %Template_Load%, Bug logging guidelines, 4
	IniRead, Input_Project, %Template_Load%, Bug logging guidelines, 5
	;Test Cases
	IniRead, Input_TestCase, %Template_Load%, Test cases, 1
	IniRead, Input_TestCaseURL, %Template_Load%, Test cases, 2
	;Locales
	IniRead, Input_Locales, %Template_Load%, Locales, 1
	;Assignee
	IniRead, Input_L10n, %Template_Load%, Assignee, 1
	IniRead, Input_NonL10n, %Template_Load%, Assignee, 2
	;Demo
	IniRead, Input_DemoTaker, %Template_Load%, Demo, 1
	;Time Tracker
	IniRead, Input_TimeTracker, %Template_Load%, Time Tracker, 1
	;Slack
	IniRead, Input_SlackChan, %Template_Load%, Slack, 1
	IniRead, Input_SlackURL, %Template_Load%, Slack, 2
	gosub CreateGUI
	}
	return

;When pressing Save Template Button
Button_Save:
	Gui, Submit, NoHide
	GuiControlGet, Input_Sender
	GuiControlGet, Input_To
	GuiControlGet, Input_CC
	GuiControlGet, Input_Subject
	GuiControlGet, Input_InternalDL
	GuiControlGet, Input_OfficialDL
	GuiControlGet, Input_Automation
	GuiControlGet, Input_Manual
	GuiControlGet, Input_BuildID
	GuiControlGet, Input_Project
	GuiControlGet, Input_Summary
	GuiControlGet, Input_AffectVersion
	GuiControlGet, Input_Component
	GuiControlGet, Input_Labels
	GuiControlGet, Input_TestCase
	GuiControlGet, Input_TestCaseURL
	GuiControlGet, Input_Locales
	GuiControlGet, Input_L10n
	GuiControlGet, Input_NonL10n
	GuiControlGet, Input_DemoTaker
	GuiControlGet, Input_TimeTracker
	GuiControlGet, Input_SlackChan
	GuiControlGet, Input_SlackURL
	FileSelectFile, Template_Save, S 16,, Save template
	; Remove all blank lines from the text in a variable:

	if Template_Save !=
	{
	Input_Project_Default := "eCAT/XPT - Content|eCAT Tool/MRA|PP - BL-Credit|PP - BL-Global Operations|PP - BL-Risk|PP - PL-Consumer|PP - PL-Merchant|PP - PL-Payments & Platform"
	Input_AffectVersion_Default := "Unknown – Not Yet Live|Unknown – On Live|N/A"
	;msgbox, %Template_Save%
	;Email,
	IniWrite, %Input_Sender%, %Template_Save%, Email, 1
	IniWrite, %Input_To%, %Template_Save%, Email, 2
	IniWrite, %Input_CC%, %Template_Save%, Email, 3
	; Deadline
	IniWrite, %Input_InternalDL%, %Template_Save%, Deadline, 1
	IniWrite, %Input_OfficialDL%, %Template_Save%, Deadline, 2
	;Stage details
	IniWrite, %Input_Automation%, %Template_Save%, Stage, 1
	IniWrite, %Input_Manual%, %Template_Save%, Stage, 2
	IniWrite, %Input_BuildID%, %Template_Save%, Stage, 3
	;Bug logging guidelines
	IniWrite, %Input_Summary%, %Template_Save%, Bug logging guidelines, 1
	IfInString, Input_AffectVersion_Default, %Input_AffectVersion%
	{
	StringReplace, Input_AffectVersion2, Input_AffectVersion_Default, %Input_AffectVersion%, %Input_AffectVersion%|
	}
	Else
	{
	Input_AffectVersion2 = %Input_AffectVersion%||%Input_AffectVersion_Default%
	}
	IniWrite, %Input_AffectVersion2%, %Template_Save%, Bug logging guidelines, 2
	IniWrite, %Input_Component%, %Template_Save%, Bug logging guidelines, 3
	IniWrite, %Input_Labels%, %Template_Save%, Bug logging guidelines, 4
	StringReplace, Input_Project2, Input_Project_Default, %Input_Project%, %Input_Project%|
	IniWrite, %Input_Project2%, %Template_Save%, Bug logging guidelines, 5
	;Test Cases
	IniWrite, %Input_TestCase%, %Template_Save%, Test cases, 1
	IniWrite, %Input_TestCaseURL%, %Template_Save%, Test cases, 2
	;Locales
	/*
	msgbox, %Input_LocalesSelected%
	msgbox, %Input_Locales%
	StringSplit, Input_LocalesSelected, Input_LocalesSelected, |,
	Loop, %Input_LocalesSelected%
	{
	Stringreplace, Input_Locales2, Input_Locales, %Input_LocalesSelected%, %Input_LocalesSelected%|
	}
	*/
	IniWrite, %Input_Locales%, %Template_Save%, Locales, 1
	;Assignee
	IniWrite, %Input_L10n%, %Template_Save%, Assignee, 1
	IniWrite, %Input_NonL10n%, %Template_Save%, Assignee, 2
	;Demo
	IniWrite, %Input_DemoTaker%, %Template_Save%, Demo, 1
	;Time Tracker
	IniWrite, %Input_TimeTracker%, %Template_Save%, Time Tracker, 1
	;Time Tracker
	IniWrite, %Input_TimeTracker%, %Template_Save%, Time Tracker, 1
	;Time Tracker
	IniWrite, %Input_SlackChan%, %Template_Save%, Slack, 1
	IniWrite, %Input_SlackURL%, %Template_Save%, Slack, 2
	;msgbox, %Input_Locales2%
	}
	return


CreateMail:
;Table DEADLINE %Email_Deadline%
If Input_InternalDL !=
	{
	Email_InternalDL = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Internal deadline</td>
	<td style="background-color: #ffffff; width: 348.069px;" scope="row">%Input_InternalDL%</td>
	</tr>
	}
Else
	{
	Email_InternalDL = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Internal deadline</td>
	<td style="background-color: #ffffff; width: 348.069px;" scope="row">TBC</td>
	</tr>
	}

If Input_OfficialDL !=
	{
	Email_OfficialDL = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Official deadline</td>
	<td style="background-color: #ffffff; width: 348.069px;" scope="row">%Input_OfficialDL%</td>
	</tr>
	}
Else
	{
	Email_OfficialDL = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Official deadline</td>
	<td style="background-color: #ffffff; width: 348.069px;" scope="row">TBC</td>
	</tr>
	}

	Email_Deadline = %Email_InternalDL%%Email_OfficialDL%



;Table STAGE DETAILS %Email_StageDetails%
If Input_Automation !=
	{
	Email_Automation = <tr>
	<td style="background-color: #ffffff; width: 125.931pxx;" scope="row">Automation</td>
	<td style="background-color: #ffffff; width: 356.757px;" scope="row"><a href="%Input_Automation%" target="_blank" rel="noopener">%Input_Automation%</a></td>
	</tr>
	}
Else
	{
	Email_Automation =
	}

If Input_Manual !=
	{
	Email_Manual = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Manual</td>
	<td style="background-color: #ffffff; width: 356.757px;" scope="row">%Input_Manual%</td>
	</tr>
	}
Else
	{
	Email_Manual =
	}

If Input_BuildID !=
	{
	Email_BuildID = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Build ID</td>
	<td style="background-color: #ffffff; width: 356.757px;" scope="row">%Input_BuildID%</td>
	</tr>
	}
Else
	{
	Email_BuildID =
	}

	{
	Email_stagedetails1 = <table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
	<tbody>
	<tr>
	<th style="width: 484px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="2" scope="colgroup"><span style="color: #ffffff;">STAGE DETAILS</span></th>
	</tr>
	}

	{
	Email_StageDetails2 = 	</tbody>
	</table>
	<br>
	}

	{
	Email_StageDetailsEmpty = <table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
	<tbody>
	<tr>
	<th style="width: 484px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="1" scope="colgroup"><span style="color: #ffffff;">STAGE DETAILS</span></th>
	</tr>
	<td style="background-color: #ffffff; width: 356.757px;" scope="row"><br>&nbsp;</td>
	</tbody>
	</table>
	<br>
	}

If (Input_Automation != "" OR Input_Manual != "" OR Input_BuildID != "")
	{
	Email_StageDetailsFull = %Email_stagedetails1%%Email_Automation%%Email_Manual%%Email_BuildID%%Email_StageDetails2%
	}
Else
	{
	Email_StageDetailsFull = %Email_StageDetailsEmpty%
	}


;Table TEST CASES %Email_TestCase%
if (Input_TestCaseURL = "" AND Input_TestCase != "")
	{
	Email_TestCase = <td style="background-color: #ffffff; width: 125.931px;" scope="row">%Input_TestCase%</td>
	}
Else if (Input_TestCaseURL != "" AND Input_TestCase != "")
	{
	Email_TestCase = <td style="background-color: #ffffff; width: 125.931px;" scope="row"><a href="%Input_TestCaseURL%" target="_blank" rel="noopener">%Input_TestCase%</a></td>
	}
Else if (Input_TestCaseURL != "" AND Input_TestCase = "")
	{
	Email_TestCase = <td style="background-color: #ffffff; width: 125.931px;" scope="row"><a href="%Input_TestCaseURL%" target="_blank" rel="noopener">%Input_TestCaseURL%</a></td>
	}
Else
	{
	Email_TestCase = <td style="background-color: #ffffff; width: 125.931px;" scope="row"><br>&nbsp;</td>
	}


;Table DEMO %Email_Demo%
if Input_DemoTaker !=
	{
	Email_Demo = <table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
	<tbody>
	<tr>
	<th style="width: 756px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="3" scope="colgroup"><span style="color: #ffffff;">DEMO</span></th>
	</tr>
	<tr>
	<td style="background-color: #ffffff; width: 192.698px;" scope="row">Demo taker</td>
	<td style="background-color: #ffffff; width: 563.302px;" colspan="2" scope="row">%Input_DemoTaker%</td>
	</tr>
	<tr>
	<td style="background-color: #ffffff; width: 192.698px;" scope="row"><a href="http://confluence.alphacrc.com/pages/viewpage.action?title=PayPal+demo+videos&amp;spaceKey=PAYP" target="_blank" rel="noopener">Intructions</a></td>
	<td style="background-color: #ffffff; width: 218.302px;" scope="row"><a href="https://paypal.app.box.com/folder/48403977311" target="_blank" rel="noopener">Upload link</a></td>
	<td style="background-color: #ffffff; width: 345px;" scope="row"><a href="http://confluence.alphacrc.com/download/attachments/26283782/Developer Portal.mp4?version=2&modificationDate=1522252791173&api=v2" target="_blank" rel="noopener">Demo example</a></td>
	</tr>
	</tbody>
	</table>
	<br>
	}

;Tble ASSIGNEE %Email_Assignee%
If Input_L10n !=
	{
	Email_Assignee_L10n = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">L10n</td>
	<td style="background-color: #ffffff; width: 345.139px;" scope="row">%Input_L10n%</td>
	</tr>
	}
Else
	{
	Email_Assignee_L10n = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">L10n</td>
	<td style="background-color: #ffffff; width: 345.139px;" scope="row">TBC</td>
	</tr>
	}


If Input_NonL10n !=
	{
	Email_Assignee_NonL10n = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Non-L10n</td>
	<td style="background-color: #ffffff; width: 345.139px;" scope="row">%Input_NonL10n%</td>
	</tr>
	}
Else
	{
	Email_Assignee_NonL10n = <tr>
	<td style="background-color: #ffffff; width: 125.931px;" scope="row">Non-L10n</td>
	<td style="background-color: #ffffff; width: 345.139px;" scope="row">TBC</td>
	</tr>
	}

	Email_Assignee = %Email_Assignee_L10n%%Email_Assignee_NonL10n%

Email_Subject = %Input_Subject%
Email_Sender = %Input_Sender%
Email_To = %Input_To%
Email_CC = %Input_CC%
Email_Body =
(Join
Hello,
<br><br>
Here is the handoff for task &#34%Input_Subject%&#34.
<br>
<br>
-  Issues filed while testing with automation support need to include label <span style="background:yellow;mso-highlight:yellow">
Automation</span>
<br>
-  Issues filed using i18n categories need to include label <span style="background:yellow;mso-highlight:yellow">
LQA_i18n</span>
<br>
- Issues filed outside of testing scope need to include label <span style="background:yellow;mso-highlight:yellow">
adhocbug</span>
<br>
- Slack channel:&#32<a href="%Input_SlackURL%">%Input_SlackChan%</a>
<br>
- <a href="https://engineering.paypalcorp.com/confluence/pages/viewpage.action?pageId=68070263">PayPal bug logging guidelines</a>
<br>
- If you cannot complete task before internal due date, let me know ASAP
<br>
- Remember to update status of task and cases on daily dashboard/automation/testrail (if manual)
<br>
<br>
<b>Task specific notes:</b>
<br>
<br>
<br>
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 474px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="2" scope="colgroup"><span style="color: #ffffff;">DEADLINE</strong></th>
</tr>
%Email_Deadline%
</tbody>
</table>
<br>
%Email_StageDetailsFull%
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 491.806px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="2" scope="colgroup"><span style="color: #ffffff;">BUG LOGGING GUIDELINES</span></th>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">Project</td>
<td style="background-color: #ffffff; width: 348.472px;" scope="row">%Input_Project%</td>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">Summary</td>
<td style="background-color: #ffffff; width: 348.472px;" scope="row">%Input_Summary%</td>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">Affect version</td>
<td style="background-color: #ffffff; width: 348.472px;" scope="row">%Input_AffectVersion%</td>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">Component</td>
<td style="background-color: #ffffff; width: 348.472px;" scope="row">%Input_Component%</td>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">Labels</td>
<td style="background-color: #ffffff; width: 348.472px;" scope="row">%Input_Labels%</td>
</tr>
</tbody>
</table>
<br>
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 491.806px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" scope="colgroup"><span style="color: #ffffff;">TEST CASES</span></th>
</tr>
<tr>
%Email_TestCase%
</tr>
</tbody>
</table>
<br>
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 491.806px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" scope="colgroup"><span style="color: #ffffff;">LOCALES</span></th>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">%Input_Locales%<br>&nbsp;</td>
</tr>
</tbody>
</table>
<br>
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 491.806px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" colspan="2" scope="colgroup"><span style="color: #ffffff;">ASSIGNEE</span></th>
</tr>
%Email_Assignee%
</tbody>
</table>
<br>
%Email_Demo%
<table style="width: 600px; border-color: #000000; background-color: #000000;" cellspacing="2" cellpadding="2">
<tbody>
<tr>
<th style="width: 491.806px; border-color: #000000; background-color: %Cell_Header_BGcolor%; text-align: center; vertical-align: middle;" scope="colgroup"><span style="color: #ffffff;">TIME TRACKER</span></th>
</tr>
<tr>
<td style="background-color: #ffffff; width: 125.931px;" scope="row">%Input_TimeTracker%<br>&nbsp;</td>
</tr>
</tbody>
</table>
<p>&nbsp;</p>
)


;example of creating a MailItem and setting it's format to HTML
olApp := ComObjCreate("Outlook.Application")
olEmail := olApp.CreateItem(0)	; olMailItem := 0
olNameSpace := olApp.GetNamespace("MAPI")
if Input_Sender !=
{
	olEmail.SendUsingAccount := olNameSpace.Accounts.Item(Email_Sender) ; Can also be an index number
}
if Input_To !=
{
	olRecipient := olEmail.Recipients.Add(Email_To)
	olRecipient.Type := 1 ; To: CC: = 2 BCC: = 3
}
if Input_CC !=
{
	olRecipient := olEmail.Recipients.Add(Email_CC)
	olRecipient.Type := 2 ; To: CC: = 2 BCC: = 3
}
;olRecipient := olEmail.Recipients.Add(RecipientBCC1)
;olRecipient.Type := 3 ; To: CC: = 2 BCC: = 3
;olRecipient := olEmail.Recipients.Add(RecipientBCC2)
;olRecipient.Type := 3 ; To: CC: = 2 BCC: = 3
;olEmail.Attachments.Add(A_Desktop "")

olEmail.BodyFormat := 2	; olFormatHTML := 2
if Input_Subject !=
{
olEmail.Subject := Email_Subject " - Start now"
}
olEmail.HTMLBody := Email_Body


olEmail.Display
