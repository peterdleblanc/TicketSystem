/*
	~--------------------------------------------------~
	~NO EDITING OF THIS DOCUMENT IS ALLOWED.
	~PLEASE RESPECT YOUR SYSTEM ADMINISTRATOR AND FOLLOW
	THE LAYED-OUT RULES WHEN USING THIS PROGRAM.
	~--------------------------------------------------~
	
	~ (c)2007 ~
	Start Date; 10.12.07
	Change Date: 11.05.07
	**********************************************
	Database Backend Programming by: Peter LeBlanc
	Layout Programmed by: Craig Sheppard
*/

function ClearControls()
{

	var info = document.frmmain;
	info.txt_Other.value = "";
	info.chk_InvalidBounceback.disabled = true;
	info.chk_Re_Assigned.disabled = true;
	info.chk_CallbackDone.disabled = true;
	info.chk_InvalidTicket.disabled = true;
		
	info.chk_InvalidBounceback.checked = false;
	info.chk_Re_Assigned.checked = false;
	info.chk_CallbackDone.checked = false;
	info.chk_InvalidTicket.checked = false;
	info.chk_InvalidTicket.style.visibility = "visible"
	
	info.cmd_CallbackReason.disabled = true;
	info.cmd_CallbackReason.value = 'n/a';
	info.txt_TTSID.disabled = true;
	info.txt_TTSID.style.visibility = "visible"
	info.cmb_Problem.disabled = true;
	info.cmb_Problem.value = 'n/a';
	info.cmb_Problem.style.visibility = "visible"
	
	info.cmb_TicketStatus.disabled = false;
	info.cmb_TicketStatus.value = "n/a";		//**NOT WORKING YET!
	
	info.cmb_TicketReassignedTo.disabled = true;
	info.cmb_TicketReassignedTo.value = "n/a";
	
	info.cmb_CallbackTo.disabled = true;
	info.cmb_CallbackTo.value = "n/a";
	
	//info.cmb_NextAction.disabled = true;
	//info.cmb_NextAction.value = "n/a";
	
	info.cmb_TicketTTRDuration.disabled = true;
	
	info.txt_ResolutionNotes.value = '';
	
	//info.cmb_Problem.style.visibility = "hidden";
	info.cmd_SubmitFeedback.style.visibility = "hidden";
	
	
	info.cmb_TicketStatus.focus()
	
}


function validate_cmb_TicketStatus()
{

	with(document.frmmain)
	{
		switch(cmb_TicketStatus.value)
		{			
			case 'Completed':				
				//Drop-down boxes				
				cmb_TicketReassignedTo.disabled = true;
				cmb_CallbackTo.disabled = true;
				cmb_CallbackTo.value = "n/a";
				//cmb_NextAction.disabled = true;
				//cmb_NextAction.value = "n/a";
				//Textboxes
				txt_ResolutionNotes.disabled = false;
				cmdsave.disabled = false;
				chk_CallbackDone.disabled = false;
				chk_InvalidTicket.disabled = false;
				
				break
			case 'Callback Needed':
				//Drop-down boxes
				cmb_TicketReassignedTo.disabled = true;
				cmb_CallbackTo.disabled = false;
				//cmb_NextAction.disabled = false;
				//Textboxes
				txt_ResolutionNotes.disabled = false;
				cmdsave.disabled = false;
				chk_CallbackDone.disabled = false;
				chk_InvalidTicket.disabled = false;
				
				break
			case 'Reassigned':
				//Drop-down boxes
				cmb_TicketReassignedTo.disabled = false;
				cmb_CallbackTo.disabled = true;
				cmb_CallbackTo.value = "n/a";
				//cmb_NextAction.disabled = true;
				//Textboxes
				txt_ResolutionNotes.disabled = false;
				cmdsave.disabled = false;
				chk_CallbackDone.disabled = false;
				chk_InvalidTicket.disabled = false;
				
				break
			case 'Never Worked':
				//Drop-down boxes
				cmb_TicketReassignedTo.disabled = true;
				cmb_CallbackTo.disabled = true;
				cmb_CallbackTo.value = "n/a";
				//cmb_NextAction.disabled = true;
				//cmb_NextAction.value = "n/a";
				//Textboxes
				txt_ResolutionNotes.value = "";
				txt_ResolutionNotes.disabled = true;
				cmdsave.disabled = true;
				chk_CallbackDone.disabled = true;
				chk_InvalidTicket.disabled = true;
				
				break
			case 'Too Early':
				//Drop-down boxes
				cmb_TicketReassignedTo.disabled = true;
				cmb_CallbackTo.disabled = true;
				cmb_CallbackTo.value = "n/a";
				//cmb_NextAction.disabled = true;
				//cmb_NextAction.value = "n/a";
				//Textboxes
				txt_ResolutionNotes.value = "";
				txt_ResolutionNotes.disabled = true;
				cmdsave.disabled = false;
				chk_CallbackDone.disabled = true;
				chk_InvalidTicket.disabled = true;
				
				break
			case 'n/a':
				//Drop-down boxes
				cmb_TicketReassignedTo.disabled = true;
				cmb_CallbackTo.disabled = true;
				cmb_CallbackTo.value = "n/a";
				//cmb_NextAction.disabled = true;
				//cmb_NextAction.value = "n/a";
				//Textboxes
				txt_ResolutionNotes.disabled = true;
				cmdsave.disabled = false;
				chk_CallbackDone.disabled = false;
				chk_InvalidTicket.disabled = false;
				
				break
			default:
				
		}
	}
}	



function verify_chk_CallbackDone()
{
	var SwitchValue;
	var info = document.frmmain;
	var TempHolder;
		
	SwitchValue = info.chk_CallbackDone.checked;
	
	
	switch(SwitchValue)
	{
		case true:
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			
			info.cmd_CallbackReason.disabled = false;
			info.cmb_TicketStatus.disabled = false;
				
			TempHolder = info.txt_NumofContacts.value;
			TempHolder++;
			info.txt_NumofContacts.value = TempHolder;
			
			break
		case false:
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			
			info.cmd_CallbackReason.value = "n/a";
			info.cmd_CallbackReason.disabled = true;
			info.cmb_TicketStatus.disabled = false;
			
			info.cmb_TicketTTRDuration.disabled = true;
			
			break
		default:
	}
}


function verify_chk_InvalidTicket()
{
	var info = document.frmmain;
	alert("Please use the Agent Feedback button at the bottom of the form")
	info.chk_InvalidTicket.checked = false;


	/*
	var SwitchValue;
	var info = document.frmmain;
		
	SwitchValue = info.chk_InvalidTicket.checked;
	

	switch(SwitchValue)
	{
		case true:
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			
			info.txt_TTSID.disabled = false;
			
			info.cmb_Problem.disabled = false;
			info.cmb_TicketStatus.disabled = false;
			info.cmb_TicketTTRDuration.disabled = true;
			
			
			break
		case false:
			
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			
			info.txt_TTSID.disabled = true;
			
			info.cmb_Problem.value = "n/a";
			info.cmb_Problem.disabled = true;
			info.cmb_TicketStatus.disabled = false;
			info.cmb_TicketTTRDuration.disabled = true;
			
			break
		default:
	*/
}

function validate_cmb_TicketReassignedTo()
{
	var SwitchValue;
	var info = document.frmmain;
		
	SwitchValue = info.cmb_TicketReassignedTo.value

	switch(SwitchValue)
	{
	
		case 'To Self':
			info.chk_InvalidBounceback.disabled = true;
			info.chk_Re_Assigned.disabled = true;
			info.chk_CallbackDone.disabled = false;
			info.chk_InvalidTicket.disabled = false;
			
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			info.chk_CallbackDone.checked = false;
			info.chk_InvalidTicket.checked = false;
			
			info.cmd_CallbackReason.disabled = true;
			info.txt_TTSID.disabled = true;
			info.cmb_Problem.disabled = true;
			
			info.cmb_TicketStatus.disabled = false;
			info.cmb_TicketReassignedTo.disabled = false;
			//info.cmb_NextAction.disabled = false;
			info.cmb_TicketTTRDuration.disabled = true;
			
			info.txt_ResolutionNotes.disabled = false;
			break
			
		case 'Back to Public Listings':
			info.chk_InvalidBounceback.disabled = true;
			info.chk_Re_Assigned.disabled = true;
			info.chk_CallbackDone.disabled = false;
			info.chk_InvalidTicket.disabled = false;
			
			info.chk_InvalidBounceback.checked = false;
			info.chk_Re_Assigned.checked = false;
			info.chk_CallbackDone.checked = false;
			info.chk_InvalidTicket.checked = false;
			
			info.cmd_CallbackReason.disabled = true;
			info.txt_TTSID.disabled = true;
			info.cmb_Problem.disabled = true;
			
			info.cmb_TicketStatus.disabled = false;
			info.cmb_TicketReassignedTo.disabled = false;
			//info.cmb_NextAction.disabled = false;
			info.cmb_TicketTTRDuration.disabled = true;
			
			info.txt_ResolutionNotes.disabled = false;
			break
	
		default:
	}
}

function validate_cmd_CallbackReason()
{

	if (document.frmmain.cmd_CallbackReason.value == "Other"){
		document.frmmain.txt_Other.disabled = false;
		document.frmmain.txt_Other.focus()
	} else {
		document.frmmain.txt_Other.disabled = true;
	}
}