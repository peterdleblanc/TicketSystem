/*
	FILE: LOGIN.JS
	DATE: 10.24.2007
	-==============-
	
	TABLE OF CONTENTS:
	1. USER ACCESS LEVEL MANAGEMENT
	2. CHANGING AN USERS ACCESS LEVEL
	
	CHAPTER 1: USER ACCESS LEVEL MANAGEMENT
	---------------------------------------
	The username management system is split four ways:
	>> basic
	>> powerusers
	>> supervisors
	>> admin
	
	The rights and interface will change depending on the allowed user level
	you have been granted. The default is basic which meants you cannot see
	the import, admin or action log buttons.
		
	Let breakdown the users levels to make it easy to know which will fix
	into an individual.
	
	BASIC [default]: Allow no special features.
	POWERUSERS: Powerusers can preform imports.
	SUPERVISORS: Allowed to preform imports and access the Actions Log.
	ADMIN: Full control can access the admin, actions log and import button.
	
	CHAPTER 2: CHANGING AN USERS ACCESS LEVEL
	-----------------------------------------
	The administrator(s) of the system have access to the admin button which
	allows for editing of users table.
	
	PLEASE CONSULT YOU SUPERVISOR OR ADMINISTRATOR TO GET YOUR ACCESS INCREASED
	IF NEEDED!
*/

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

//Globals Variables



function CheckLogin(username,password)
{
	var SystemLog = new Database_Class('SystemLog')
	
   	var EmpNum;
	var Password;
	var LoginPassed;
	
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;
	var TodaysHr;
	var TodayMin;
	var TodaysMinutes;

	var myDate=new Date();
	
	  
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysHr = myDate.getHours();
	TodaysMin = myDate.getMinutes();
	TodaysMinutes = myDate.getMinutes();
	TodaysTime = TodaysHr + ":" + TodaysMin;
	
	if (username != null) {
		EmpNum = username;
		document.frmMainLogin.txt_EmpNum.value = EmpNum;
	} else {
		EmpNum = document.frmMainLogin.txt_EmpNum.value;
	}
	
	EmpCMS = "";
	
	if (username != null) {
		Password = password;
	} else {
		Password = document.frmMainLogin.txt_Password.value;
	}
	
    window.grStr_login="SELECT * FROM UserLogins WHERE EmpNum='"+ EmpNum + "'";
	
    OpenDatabase_login()
    OpenTable_login()
    
	LoginPassed = 'n';
	
	if(recSet_login.RecordCount == 0)
	{
		document.getElementById('lbl_LoginStatus').innerText = "Username not found!";
		showStatus('Username not found!');
	}
	
	if(recSet_login.RecordCount == 1 && recSet_login.Fields('Password').value != Password)
	{
		document.getElementById('lbl_LoginStatus').innerText = "Login Failed! Invalid Password, please recheck the login and try again.";
		showStatus('Login Failed! Invalid Password, please recheck the login and try again.');
			
		document.getElementById('txt_Password').value = "";
		document.frmMainLogin.txt_Password.focus()
	}
	
	if(recSet_login.RecordCount == 1 && recSet_login.Fields('Password').value == Password)
	{
		LoginPassed='y';
	}
	
	if (recSet_login.Fields('Defaults').value == true)
	{
		if(LoginPassed == 'y')
		{
			div_main.style.visibility="Visible";
			document.frmmain.txt_EmployeeNum.value = EmpNum;
			//document.frmmain.txt_EmpCMS.value = EmpCMS;
			document.frmmain.txt_EmployeeName.value = recSet_login.Fields('ShortName').value;
			loginscreen.style.visibility="hidden";
			
			document.getElementById('lbl_currentrequest').innerText = " User:"+recSet_login.Fields('Name').value+"("+EmpNum+")";
			div_processingrequest.style.visibility="hidden";
			
			//document.frmmain.cmd_modifyhelpfile.style.visibility="hidden";
			
			if(recSet_login.Fields('DefaultsUserLevel').value == "")
			{
				document.frmmain.cmd_import.style.visibility="hidden";
				document.frmmain.cmd_actionlog.style.visibility="hidden";
				document.frmmain.cmd_admin.style.visibility="hidden";
				//document.frmmain.cmd_vieworphans.style.visibility="hidden";
				document.frmmain.cmd_RunningTotals.style.visibility="hidden";
			}
			else if(recSet_login.Fields('DefaultsUserLevel').value == "basic")
			{
				document.frmmain.cmd_import.style.visibility="hidden";
				document.frmmain.cmd_actionlog.style.visibility="hidden";
				document.frmmain.cmd_admin.style.visibility="hidden";
				//document.frmmain.cmd_vieworphans.style.visibility="hidden";
				document.frmmain.cmd_RunningTotals.style.visibility="hidden";
			}
			else if(recSet_login.Fields('DefaultsUserLevel').value == "supervisor")
			{
				document.frmmain.cmd_admin.style.visibility="hidden";
			}
			else if(recSet_login.Fields('DefaultsUserLevel').value == "powerusers")
			{
				//document.frmmain.cmd_actionlog.style.visibility="hidden";
				document.frmmain.cmd_admin.style.visibility="hidden";
			}
			
			document.getElementById('lbl_LoginStatus').innerText = "";
			showStatus('Login Complete');
			
			//GlobalSystemTimmer=setTimeout('timedCount()',500);
			timedCount()
			SystemTimeStart = t;
			
			SystemLog.Path = "T:\\All In One\\Reports Viewer\\pub\\databases\\";
			SystemLog.DBName = 'UserInfo.mdb';
			SystemLog.QryString = "SELECT * FROM SystemLog";
			
			
			SystemLog.EmpNum = document.frmmain.txt_EmployeeNum.value;
			SystemLog.SystemName = 'Defaults';
			SystemLog.UserAction = 'Login';
			
			
			
			SystemLog.OpenTable()
				SystemLog.recSet.AddNew
					SystemLog.recSet.Fields('EmpNum').value = SystemLog.EmpNum
					SystemLog.recSet.Fields('SystemName').value = SystemLog.SystemName
					SystemLog.recSet.Fields('Action').value = SystemLog.UserAction
					SystemLog.recSet.Fields('Date').value = TodaysDate
				//	SystemLog.recSet.Fields('Time').value = TodaysTime
				SystemLog.recSet.Update
			SystemLog.CloseTable()
			
			recSet_login.close
			cObj_login.close
			
			recSet_login = null;
			cObj_login = null;
			
			//alert(GlobalSystemTimmer)
			ChangeUserStatus(1,EmpNum)
			
			try
			{
				
				PullRecord()
				
			}
			catch(err)
			{
				
			}
			
			div_processingrequest.style.visibility="hidden";
		}
	}
	else
	{
		//Access Denied!
		document.getElementById('lbl_LoginStatus').innerText = "Access Denied!";
	}
}

function ChangePassword()
{
	var EmpNum;
	var Password;
	var NewPassword1;
	var NewPassword2;
	
	NewPassword1 = prompt("Enter password")
	NewPassword2 = prompt("Confirm Password")
	
	if(NewPassword1 == NewPassword2)
	{
		EmpNum = document.frmMainLogin.txt_EmpNum.value;
	
		window.grStr_login="SELECT * FROM UserLogins WHERE EmpNum='" + EmpNum + "'";
			
		OpenDatabase_login()
		OpenTable_login()		
		recSet_login.WillChange
			recSet_login.Fields('Password').value = NewPassword1;
		recSet_login.Update
		
		recSet_login.close
		cObj_login.close
		
		recSet_login = null;
		cObj_login = null;
		
		alert('Password Changed')
	}
	else
	{
		alert('Password did not match')
	}
	
}


function ValidateUser(button)
{
	if (button == "")
	{
		showStatus('Nothing to Validate! Please consult your system administrator for further help.');
	}
	else
	{
		//Additional Items to remove.
		frmmain.chk_CallbackDone.style.visibility="hidden";
		frmmain.chk_InvalidTicket.style.visibility="hidden";
		frmmain.cmb_Problem.style.visibility="hidden";
		frmmain.txt_TTSID.style.visibility="hidden";
		
		div_main.style.visibility="hidden";
		div_logincheck.style.visibility="visible";
		
		frmLoginCheck.txt_CurrentEmpNum.value = frmmain.txt_EmployeeNum.value;
		frmLoginCheck.hidden_requesttype.value = button;
	}
}

function ValidateLogin()
{
	if (frmLoginCheck.hidden_requesttype.value == "")
	{
		showStatus('Unable to process! Please re-submit you request.');
	}
	else
	{
		var EmpNum;
		var Password;
		var LoginPassed;
		var Request;
	  	  
		window.grStr_login="SELECT * FROM UserLogins Where EmpNum='"+frmLoginCheck.txt_CurrentEmpNum.value +"'";
		EmpNum = document.frmLoginCheck.txt_CurrentEmpNum.value;
		Password = document.frmLoginCheck.txt_Password.value;
		Request = frmLoginCheck.hidden_requesttype.value;
		
		document.frmLoginCheck.txt_Password.value="";
		
		OpenDatabase_login()
		OpenTable_login()
		

		LoginPassed = 'n';
		recSet_login.MoveFirst
		while(recSet_login.EOF == 0)
		{
			if(recSet_login.Fields('EmpNum').value == EmpNum)
			{
				if(recSet_login.Fields('Password').value == Password)
				{
					LoginPassed = 'y';
					
					var SwitchValue;
					SwitchValue = recSet_login.Fields('DefaultsUserLevel').value
					
					document.getElementById('lbl_currentrequest').innerText = " "+Request;
					div_processingrequest.style.visibility="visible";
					
					switch(SwitchValue)
					{
						case 'supervisor':
							if (Request == "Import")
							{
								ImportData();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							if (Request == "ActionsLog") DisplayActions_Log();
							if (Request == "ViewOrphans")
							{
								ViewOrphans();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							
							break
						case 'basic':
							if (Request == "ViewOrphans")
							{
								ViewOrphans();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							
							break
						case 'powerusers':
							if (Request == "Import")
							{
								ImportData();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							if (Request == "ViewOrphans")
							{
								ViewOrphans();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							if (Request == "ActionsLog") DisplayActions_Log();
							
							break
						case 'admin':
							if (Request == "Import")
							{
							
								ImportData();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							if (Request == "Administration")
							{
								document.getElementById('lbl_CurrentLoggedIn').innerText = recSet_login.Fields('Name').value;
								DisplayAdminOptions();
							}
							if (Request == "ActionsLog") DisplayActions_Log();
							if (Request == "ViewOrphans")
							{
								ViewOrphans();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							
							break
						case 'superadmin':
							if (Request == "Import")
							{
							
								ImportData();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							if (Request == "Administration")
							{
								document.getElementById('lbl_CurrentLoggedIn').innerText = recSet_login.Fields('Name').value;
								DisplayAdminOptions();
							}
							if (Request == "ActionsLog") DisplayActions_Log();
							if (Request == "ModifyHelp") DisplayModifyHelp();
							if (Request == "ViewOrphans")
							{
								ViewOrphans();
								
								div_logincheck.style.visibility="hidden";
								div_main.style.visibility="visible";
							}
							
							break
					}
					
					div_processingrequest.style.visibility="hidden";
				}
				else
				{
					showStatus('Validation Failed! Please recheck the login and try again.');
					CancelValidation()
				}
			}
			recSet_login.MoveNext
		}
		recSet_login.close
		cObj_login.close
		
		recSet_login = null;
		cObj_login = null;
	}
}

function CancelValidation()
{
	frmLoginCheck.txt_CurrentEmpNum.value = '';
	
	//Additional Items to remove.
	frmmain.chk_CallbackDone.style.visibility="visible";
	frmmain.chk_InvalidTicket.style.visibility="visible";
	frmmain.cmb_Problem.style.visibility="visible";
	frmmain.txt_TTSID.style.visibility="visible";
	
	div_main.style.visibility="visible";
	div_logincheck.style.visibility="hidden";
}

function SetPageControls()
{
	div_main.style.visibility="hidden";
}

function ChangeUserStatus(status,employeenum)
{
	var StatsDB = new Database_Class('StatsDB')
	
	StatsDB.Path = "T:\\All In One\\Reports Viewer\\pub\\databases\\";
	StatsDB.DBName = 'UserInfo.mdb';
	StatsDB.QryString = "SELECT * FROM UserLogins WHERE EmpNum='"+ employeenum + "'";
	
	StatsDB.OpenTable()
	StatsDB.recSet.WillChange
	
	switch(status)
	{
		case 0:
			StatsDB.recSet.Fields('Status').value = 'Offline';
			break
		case 1:
			StatsDB.recSet.Fields('Status').value = 'Active';
			break
		default:
			StatsDB.recSet.Fields('Status').value = "Offline";
			break
	}
	
	StatsDB.recSet.Update
	StatsDB.CloseTable()
}