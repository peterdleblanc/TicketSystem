~--------------------------------------------------~
~NO EDITING OF THIS DOCUMENT IS ALLOWED.
~PLEASE REPECT YOUR SYSTEM ADMINISTRATOR AND FOLLOW
THE LAYED-OUT RULES WHEN USING THIS PROGRAM.
~--------------------------------------------------~

~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~
Web Default Tickets Readme File
~=~=~=~=~=~=~=~=~=~=~=~=~=~=~=~

Table of Contents
~=~=~=~=~=~=~=~=~
1. Getting Started
2. Folder Structure
3. Version History

Chapter 1: Getting Started
~=~=~=~=~=~=~=~=~=~=~=~=~=
Welcome to the NEW! online default tickets database.
	
To open the database you can either open the "Launch Default Tickets [WEB_VERSION]" shortcut or open the "index.hta" icons in the root folder located in: "t:\all in one\default_tickets".

Once you have the system open you will be prompted to enter your employee # and password.

If you have forgotten your password please see your system administrator to have your login created.

Assuming you are now logged-in you can get access to additional help by clicking on the "Help Index" icon on the top right of the screen.

Chapter 2: Folder Structure
~=~=~=~=~=~=~=~=~=~=~=~=~=~
	T:.
	|   calendar.html
	|   calendar2.js
	|   index.hta
	|   Readme.txt
	|   styles.css
	|   
	+---img
	|       next.gif
	|       next_year.gif
	|       pixel.gif
	|       prev.gif
	|       prev_year.gif
	|       
	\---pub
		+---databases
		|       DefaultTickets_revised.ldb
		|       DefaultTickets_revised.mdb
		|       Stats.ldb
		|       Stats.mdb
		|       
		+---images
		|   +---backgrounds
		|   |       page_bg.jpg
		|   |       
		|   +---buttons
		|   |       img_agentfeedback_report.gif
		|   |       img_completed.gif
		|   |       img_createuser.gif
		|   |       img_default_tickets_totals.gif
		|   |       img_deleteuser.gif
		|   |       img_help.gif
		|   |       img_invalidagentfeedback_report.gif
		|   |       img_promoteuser.gif
		|   |       img_releaseheldticket.gif
		|   |       img_resetuser.gif
		|   |       img_totalticketcomplete_today.gif
		|   |       img_unlockagent_ticket.gif
		|   |       img_unlockall_tickets.gif
		|   |       
		|   +---headers
		|   |       header_bg.jpg
		|   |       
		|   +---icons
		|   |       icon_copy.jpg
		|   |       img_cal.gif
		|   |       next.gif
		|   |       next_year.gif
		|   |       prev.gif
		|   |       prev_year.gif
		|   |       
		|   \---screenshots
		|           interface.png
		|           interface_bak.png
		|           interface_ticketstatusdd.png
		|           interface_ticketstatussection.png
		|           interface_topsection.png
		|           
		+---js
		|       connection.js
		|       controls.js
		|       disabled_js_functions.js
		|       help.js
		|       login.js
		|       main.js
		|       reports.js
		|       usermanagement.js
		|       validation.js
		|       
		+---logs
		|       developers.log
		|       
		+---templates
		|       ActionsLog.html
		|       CommentHistory.html
		|       Help.html
		|       HelpContents.html
		|       Output.html
		|       
		\---tmp
				table_template.html

3. Version History
	11.05.07 - Version 1.4 - Beta test version RC1
	>> Fully working version without any error tracking on one specific error from the Pull function.
	
	11.06.07 - Version 1.5 - RC2 Pre-Release
	>> Fully working version, ready for release.
	
	11.07.07 - Version 1.5.1 - RC2 Pre-Release
	>> All exported fields are not being imported into the TO_BE_COMPLETED and completed tables.
	
	11.09.07 - Version 1.5.5 - Removed the Min, Max and Close buttons.
	Issue related to database being closed incorrectly.

4. Copyright Notice
	~ (c)2007 ~
	Start Date; 10.12.07
	Change Date: 11.05.07
	**********************************************
	Database Backend Programming by: Peter LeBlanc
	Layout Programmed by: Craig Sheppard