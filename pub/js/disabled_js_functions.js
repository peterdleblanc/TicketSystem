
//-====================================================================================================-
//	Create New User Functions
//-====================================================================================================-
function ValidateNewUser()
{
	var targetForm = document.frmCreateLogin;
	
	//Validate all form fields and display an errors, if no errors
	//set the value of iValidationFlag to equal 1.
	if (targetForm.txt_EmpNum.value == '')
	{
		showStatus('Employee # is required! Please enter and try again.');
		targetForm.txt_EmpNum.focus();
	}
	else if (targetForm.txt_Name.value == '')
	{
		showStatus('Agent name is required! Please enter and try again.');
		targetForm.txt_Name.focus();
	}
	else if (targetForm.txt_Password.value == '')
	{
		showStatus('Password is required! Please enter and try again.');
		targetForm.txt_Password.focus();
	}
	else if (targetForm.txt_Password.value.length < 5)
	{
		showStatus('Password length must be greater then five! Please enter a new password and try again.');
		targetForm.txt_Password.focus();
	}
	else
	{
		CreateLogin(targetForm.txt_EmpNum.value, targetForm.txt_Name.value, targetForm.txt_Password.value,"Basic");
		
		showStatus('Login Created Sucessfully.');
		CancelCreate();
	}
}

//-====================================================================================================-

function PopulateHelpfileEntries()
{
	alert('need to code');
}

function DisplayModifyHelp()
{
	div_logincheck.style.visibility="hidden";
	
	showStatus('Welcome to the Online Helpfile Editor');
	PopulateHelpfileEntries();
	
	div_modifyhelpfilescreen.style.visibility="visible";
}

function CancelHelpfileScreen()
{
	div_modifyhelpfilescreen.style.visibility="hidden";
	showStatus('');
	div_main.style.visibility="visible";
}

function ValidateNewEntry()
{
	var targetForm = document.frmModifyHelpfile;
	
	//Validate all form fields and display an errors, if no errors
	//set the value of iValidationFlag to equal 1.
	if (targetForm.txt_ControlEnabled.value == '')
	{
		showStatus('A 0 or 1 is required! Please enter and try again.');
		targetForm.txt_ControlEnabled.focus();
	}
	else if (targetForm.txt_Title.value == '')
	{
		showStatus('Title of help topic is required! Please enter and try again.');
		targetForm.txt_Title.focus();
	}
	else if (targetForm.txt_Text.value == '')
	{
		showStatus('Contents of help topic is required! Please enter and try again.');
		targetForm.txt_Text.focus();
	}
	else
	{
		CreateNewHelpfileEntry(targetForm.txt_EmpNum.value, targetForm.txt_Name.value, targetForm.txt_Password.value,"Basic");
		
		showStatus('Entry Created Sucessfully.');
		CancelCreate();
	}
}

function CreateNewHelpfileEntry(Enabled, Title, Contents)
{
	window.grStr3="SELECT * FROM tblTOC_Contents";
	OpenDatabase2()
	OpenTable3()

	recSet3.AddNew
		recSet3.Fields('Enabled').value = Enabled;
		recSet3.Fields('Title').value = Title;
		recSet3.Fields('Text').value = Contents;
		recSet3.Fields('DateAdded').value = Date();
	recSet3.Update

	recSet3.close
	cObj2.close
	recSet3 = null;
	cObj2 = null;
}

function UpdateHelpfileEntry()
{
	alert('need to write this!');
}
//DISABLED FUNCTION
/*
function closeHelp()
{
	div_helpcontents.style.visibility="hidden";
}

function populateTOC(){
	var RecordCount;
	
	window.grStr="SELECT * FROM tblTOC_Contents WHERE ControlEnabled = '1'";
	
	OpenDatabase()
	OpenTable()
	
	RecordCount = 1;
	recSet.MoveFirst;
	
	while(!recSet.eof) {
		document.write('<tr>');
		document.write('	<td>'+RecordCount+'. '+recSet.Fields("title").value+'</td>');
		document.write('</tr>');
		
		RecordCount++;
		recSet.MoveNext
   }
   
   recSet.close
   cObj.close
}

function populateTOC_Contents(){
	var RecordCount;
	
	window.grStr="SELECT * FROM tblTOC_Contents WHERE ControlEnabled = '1'";
	
	OpenDatabase()
	OpenTable()
	
	RecordCount = 1;
	recSet.MoveFirst;
	
	while(!recSet.eof) {
		document.write('<!-- Chapter '+RecordCount+':'+recSet.Fields("title").value+' -->');
		document.write('<tr>');
		document.write('	<td>');
		
		document.write('		<table width="100%" cellspacing="0" cellpadding="1" class="tlb_outer">');
		
		document.write('			<tr class="text_subheader">');
		document.write('				<td>Chapter '+RecordCount+': '+recSet.Fields("title").value+'</td>');
		document.write('			</tr>');
		
		document.write('			<tr class="text_toc-normal">');
		document.write('				<td>'+recSet.Fields("text").value+'</td>');
		document.write('			</tr>');
		
		document.write('		</table>');
		
		document.write('	</td>');
		document.write('</tr>');
		
		RecordCount++;
		recSet.MoveNext
   }
   
   recSet.close
   cObj.close
}


	
	
	<!-- FORMS TO BE DELETED -->
	<!--
		<form name="frmCreateLogin">
			<div id="div_createlogin" name="div_createlogin" style="position:absolute; visibility:hidden; z-index:5;">
				<table width="100%" cellspacing="0" cellpadding="0" style="text-align: center;">
					<tr>
						<td>
							<table width="30%" cellspacing="0" cellpadding="1" class="tlb_outer" style="text-align: center;">
								<tr class="text_header">
									<td>Create Login</td>
								</tr>
								<tr class="text_larger" style="text-align:left; background-color:#F8F8F8;">
									<td>
										Please fill-out the form below to create yourself a login into the system. If you need more access in the system please consult your system administrator.
									</td>
								</tr>
								<tr valign="top">
									<td>
										<table name="tlbmain" id="tlbmain" class="text_normal">
											<tr>
												<td style="text-align:right;"><label for="txt_EmpNum" accesskey="e"><u>E</u>mployee #:</label></td>
												<td><input id="txt_EmpNum" name="txt_EmpNum" maxlength="9" type="text" size="12" class="text_normal"></td>
											</tr>
											<tr>
												<td style="text-align:right;"><label for="txt_Name" accesskey="n"><u>N</u>ame:</label></td>
												<td><input id="txt_Name" name="txt_Name" maxlength="30" type="text" size="20" class="text_normal"></td>
											</tr>
											<tr>
												<td style="text-align:right;"><label for="txt_Password" accesskey="p"><u>P</u>assword:</label></td>
												<td><input id="txt_Password" name="txt_Password" maxlength="20" type="password" size="16" class="text_normal"></td>
											</tr>
											<tr>
												<td>&nbsp;</td>
												<td>
													<label for="cmd_Create" accesskey="c"><input id="cmd_Create" type="button" value="Create" class="text_subheader text_normal text_buttons" onClick="ValidateNewUser();"></label>
													<label for="cmd_Cancel" accesskey="l"><input id="cmd_Cancel" type="button" value="Cancel" class="text_subheader text_normal text_buttons" onClick="CancelCreate();"></label>
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</div>
		</form>
	-->
	<!-- << END -->
	<!-- FORMS TO BE DELETED -->
	
	
*/