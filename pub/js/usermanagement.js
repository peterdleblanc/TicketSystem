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

/*
	Form Controls:
	- txt_UMEmpNum
	- txt_UMPassword
	- txt_UMName
	- cbm_UMUserLevel
*/

function manageUsers(action)
{
	var sEmpNum;
	var sPassword;
	var sName;
	var sUserLevel;
	var sMsg;
	
	showStatus('');
	
	with(document.frmAdmin)
	{
		sEmpNum = txt_UMEmpNum.value;
		sPassword = txt_UMPassword.value;
		sName = txt_UMName.value;
		sUserLevel = cbm_UMUserLevel.value;
	}
	
	//Set SQL Querys
	switch(action)
	{			
		case 'promote':
			window.grStr_login = "SELECT * FROM UserLogins WHERE EmpNum = '" + sEmpNum + "'";
			break
		case 'create':
			window.grStr_login = "SELECT * FROM UserLogins";
			break
		case 'reset':
			window.grStr_login = "SELECT * FROM UserLogins WHERE EmpNum = '" + sEmpNum + "'";
			break
		case 'delete':
			if (sUserLevel != "admin")
			{
				window.grStr_login = "DELETE FROM UserLogins WHERE EmpNum = '" + sEmpNum + "' AND Name = '" + sName + "'";
			}
			else
			{
				showStatus('Unable to complete operation! Cannot delete an admin user using this method');
			}
			break
	}
	
	OpenDatabase_login()	//Open Database
	OpenTable_login()		//Open Table
	
	recSet_login.MoveFirst;
	
	//Process associated actions
	if (action == "delete")
	{
		sMsg = sName + '(' + sEmpNum + ') has been deleted!';
		showStatus(sMsg);
	}
	else
	{
		if (recSet_login.RecordCount > 0)
		{
			while(recSet_login.EOF == 0)
			{
				if (recSet_login.Fields('EmpNum').value == sEmpNum)
				{
					switch(action)
					{
						case 'promote':
							if (sUserLevel != "n/a")
							{
								recSet_login.WillChange
								recSet_login.Fields('DefaultsUserLevel').value = sUserLevel;
								recSet_login.Update
							}
							
							sMsg = sName + ' has been promoted to ' + sUserLevel + ' level.';
							showStatus(sMsg);
							
							break
						case 'reset':
							if (sPassword != "")
							{
								recSet_login.WillChange
								recSet_login.Fields('Password').value = sPassword;
								recSet_login.Update
								
								sMsg = 'Password has been reset on employees #' + sEmpNum + ' account.';
								alert(sMsg);
							}
							
							break
					}
					
					break
				}
				else
				{
					switch(action)
					{
						case 'create':
							recSet_login.AddNew
							
							recSet_login.Fields('EmpNum').value = sEmpNum;
							recSet_login.Fields('Name').value = sName;
							recSet_login.Fields('Password').value = sPassword;
							recSet_login.Fields('DefaultsUserLevel').value = sUserLevel;
							recSet_login.Fields('Defaults').value = 1;
							recSet_login.Fields('TimeAdjuster').value = 1;
							
							recSet_login.Update
							
							sMsg = sName + ' has been created!';
							alert(sMsg);
							break
					}
				}
				
				recSet_login.MoveNext
			}
		}
		else
		{
			alert('User not found! Please check and re-try.');
		}
	}
	
	recSet_login.close
	cObj_login.close
	
	recSet_login = null;
	cObj_login = null;
}