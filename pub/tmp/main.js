/*
	~--------------------------------------------------~
	~NO EDITING OF THIS DOCUMENT IS ALLOWED.
	~PLEASE RESPECT YOUR SYSTEM ADMINISTRATOR AND FOLLOW
	THE LAYED-OUT RULES WHEN USING THIS PROGRAM.
	~--------------------------------------------------~
	
	~ (c)2007 ~
	Start Date; 10.12.07
	Change Date: 11.10.07
	**********************************************
	Database Backend Programming by: Peter LeBlanc
	Layout Programmed by: Craig Sheppard
*/

/* ************** */
/* MISC FUNCTIONS */
/* ************** */

	function generate_BaseNote()
	{
		var Current_Date = get_CurDate();
		var Current_Time = get_CurTime();
		var Current_User_Name = document.frmmain.txt_EmployeeName.value;
		var Current_User_CMS = "INSERTCMS";
		
		document.frmmain.txt_ResolutionNotes.value = Current_User_Name+" "+Current_User_CMS+" STJTASC "+Current_Date+" "+Current_Time;
		document.frmmain.txt_ResolutionNotes.focus()
		CheckAllowedChars()
	}
	
	function get_CurDate()
	{
		var generated_date = new Date();
		var sDay = generated_date.getDate();
		var sMonth = generated_date.getMonth();
		var sMonth = sMonth + 1;
		var sYear = generated_date.getYear();
		var CurrentDate = sMonth + "/" + sDay + "/" + sYear;
		
		return CurrentDate;
	}
	
	function get_CurTime()
	{
		var generated_time = new Date();
		var sHours = generated_time.getHours();
		var sMinute = generated_time.getMinutes();
		sMinute = sMinute+'';
		if (sMinute.length == 1) sMinute = "0"+sMinute;
		var CurrentTime = sHours + ":" + sMinute;
		
		return CurrentTime;
	}

	//Function tested and working: 11.05.07
	function CopyText(value) {
		if (value != ""){
			var source = value;
			
			frmmain.globalhidden.value = "";
			frmmain.globalhidden.value = value;
			
			var tempval=eval("document.frmmain.globalhidden")
			var copytoclip=1
			
			if (document.all&&copytoclip==1){
				therange=tempval.createTextRange()
				therange.execCommand("Copy")
			}
		}
	}
	
	function ForceLogin()
	{
		if (window.event.keyCode == 13) CheckLogin()
	}

	function TestImport()
	{
	
	}
	
	function ImportCSV()
	
	{
		var i;

		window.grStr="SELECT * FROM Report.csv";
		window.grStr2="SELECT * FROM defaultImport";
		OpenDatabaseCSV()
		OpenDatabase()
		OpenTableCSV()
		OpenTable2()
		alert(recSetCSV.RecordCount)
		
		for(i=0;i<recSetCSV.RecordCount;i++)
		{
		recSet2.AddNew
			recSet2.Fields('ASSIGNED_DATE').value=recSetCSV.Fields(0).value  		//ASSIGNED_DATE
			recSet2.Fields('ASSIGNED_GROUP').value=recSetCSV.Fields(1).value			//ASSIGNED_GROUP
			recSet2.Fields('ASSIGNED_TO').value=recSetCSV.Fields(2).value			//ASSIGNED_TO
			recSet2.Fields('AUTO_CORRELATE').value=recSetCSV.Fields(3).value			//AUTO_CORRELATE
			recSet2.Fields('CATEGORY').value=recSetCSV.Fields(4).value			//CATEGORY
			recSet2.Fields('CORRELATION_TYPE').value=recSetCSV.Fields(5).value			//CORRELATION_TYPE
			recSet2.Fields('CORRELATION_VALUE_1').value=recSetCSV.Fields(6).value			//CORRELATION_VALUE_1
			recSet2.Fields('CORRELATION_VALUE_2').value=recSetCSV.Fields(7).value			//CORRELATION_VALUE_2
			recSet2.Fields('CREATE_DATE').value=recSetCSV.Fields(8).value			//CREATE_DATE
			recSet2.Fields('DEPARTMENT').value=recSetCSV.Fields(9).value			//DEPARTMENT
			recSet2.Fields('ETA').value=recSetCSV.Fields(10).value			//ETA
			recSet2.Fields('INDICATOR_CITY').value=recSetCSV.Fields(11).value			//INDICATOR_CITY
			recSet2.Fields('INDICATOR_STATE').value=recSetCSV.Fields(12).value			//INDICATOR_STATE
			recSet2.Fields('MultiRegion Affected Names').value=recSetCSV.Fields(13).value			//MultiRegion Affected Names
			recSet2.Fields('MultiRegion Affected').value=recSetCSV.Fields(14).value			//MultiRegion Affected
			recSet2.Fields('MultiRegion Flag').value=recSetCSV.Fields(15).value			//MultiRegion Flag
			recSet2.Fields('NETWORK_ELEMENT').value=recSetCSV.Fields(16).value			//NETWORK_ELEMENT
			recSet2.Fields('NEXT_ACTION').value=recSetCSV.Fields(17).value			//NEXT_ACTION
			recSet2.Fields('NIU ID').value=recSetCSV.Fields(18).value			//NIU ID
			recSet2.Fields('Node').value=recSetCSV.Fields(19).value			//Node
			recSet2.Fields('OFFICE_NAME').value=recSetCSV.Fields(20).value			//OFFICE_NAME
			recSet2.Fields('PARENT_TICKET').value=recSetCSV.Fields(21).value			//PARENT_TICKET
			recSet2.Fields('PRIMARY_PHONE').value=recSetCSV.Fields(22).value			//PRIMARY_PHONE
			recSet2.Fields('PRIORITY').value=recSetCSV.Fields(23).value			//PRIORITY
			recSet2.Fields('PROBLEM_CODE').value=recSetCSV.Fields(24).value			//PROBLEM_CODE
			recSet2.Fields('PROBLEM_SUMMARY').value=recSetCSV.Fields(25).value			//PROBLEM_SUMMARY
			recSet2.Fields('PRODUCT').value=recSetCSV.Fields(26).value			//PRODUCT
			recSet2.Fields('PURGE_DATE').value=recSetCSV.Fields(27).value			//PURGE_DATE
			recSet2.Fields('QUEUE_ID').value=recSetCSV.Fields(28).value			//QUEUE_ID
			recSet2.Fields('QUEUE_NAME').value=recSetCSV.Fields(29).value			//QUEUE_NAME
			recSet2.Fields('REGION_ID').value=recSetCSV.Fields(30).value			//REGION_ID
			recSet2.Fields('REGION_NAME').value=recSetCSV.Fields(31).value			//REGION_NAME
			recSet2.Fields('SA_ID').value=recSetCSV.Fields(32).value			//SA_ID
			recSet2.Fields('SERVICE_AFFECT').value=recSetCSV.Fields(33).value			//SERVICE_AFFECT
			recSet2.Fields('Severity').value=recSetCSV.Fields(34).value			//Severity
			recSet2.Fields('STATUS').value=recSetCSV.Fields(35).value			//STATUS
			recSet2.Fields('SUBMITTER').value=recSetCSV.Fields(36).value			//SUBMITTER
			recSet2.Fields('SYSTEM_ID').value=recSetCSV.Fields(37).value			//SYSTEM_ID
			recSet2.Fields('SYSTEM_NAME').value=recSetCSV.Fields(38).value			//SYSTEM_NAME
			recSet2.Fields('TECH_ARRIVAL_TIME').value=recSetCSV.Fields(39).value			//TECH_ARRIVAL_TIME
			recSet2.Fields('TECHNICIAN_NAME').value=recSetCSV.Fields(40).value			//TECHNICIAN_NAME
			recSet2.Fields('TECHNICIAN_NUMBER').value=recSetCSV.Fields(41).value			//TECHNICIAN_NUMBER
			recSet2.Fields('TICKET_ID').value=recSetCSV.Fields(42).value			//TICKET_ID
			recSet2.Fields('TTR_SLA').value=recSetCSV.Fields(43).value			//TTR_SLA
			recSet2.Fields('TTR_START').value=recSetCSV.Fields(44).value			//TTR_START
			recSet2.Fields('WF-Scratch1').value=recSetCSV.Fields(45).value			//WF-Scratch1
		recSet2.update
		recSetCSV.MoveNext
		recSet2.MoveNext
		}
	
	recSet2.close
	cObj.close
		
	recSetCSV.close
	cObjCSV.close
	
	alert('Done')
	}


/* **************** */
/* TICKET FUNCTIONS */
/* **************** */
function PullRecord()
{
	showStatus('New Record');
	ClearControls()
	
	var SwitchValue;
	var info = document.frmmain;
	
	var EmpNum;
    var ExitFlag;
    var LastRecFlag;
    var UserAction;
    var TICKETID;
	var Workable;
	var TestValue;
	
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;
	var TodaysMinutes;

	var myDate=new Date();
  
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours();
	TodaysMinutes = myDate.getMinutes();
	
	SwitchValue = info.cmd_WorkedLockedTickets.value
	EmpNum = document.frmmain.txt_EmployeeNum.value;
	
		
		switch(TodaysTime)
		{	
			case 10:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
				break
			case 12:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
				break
			case 13:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 14:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 15:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 16:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 17:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 18:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 19:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 20:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 21:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 22:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 23:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 24:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break	
			case 1:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 2:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' ORDER BY CREATE_DATE ASC";
				break
			default:
				window.grStr2="SELECT * FROM TO_BE_WORKED WHERE INDICATOR_STATE='MA' ORDER BY CREATE_DATE ASC";
				break
		}
	
	
	
	//window.grStr2="SELECT * FROM TO_BE_WORKED ORDER BY CREATE_DATE ASC";
	OpenDatabase()
	OpenTable2()
	
	//rs.Filter="Lastname ='Jonson' OR Lastname='Johnson'" 
	//recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY' OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
	/*
		switch(TodaysTime)
		{	
			case 10:
				recSet2.Filter="INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
				break
			case 12:
				recSet2.Filter="INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
				break
			case 13:
				recSet2.Filter="INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				
				break
			case 14:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT'";
				break
			case 15:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 16:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 17:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 18:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 19:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 20:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 21:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' OR INDICATOR_STATE='CT' OR INDICATOR_STATE='FL' OR INDICATOR_STATE='GA' OR INDICATOR_STATE='PA' OR INDICATOR_STATE='MA' OR INDICATOR_STATE='MI' OR INDICATOR_STATE='NH' OR INDICATOR_STATE='OH' OR INDICATOR_STATE='NY' OR INDICATOR_STATE='VA' OR INDICATOR_STATE='VT' ORDER BY CREATE_DATE ASC";
				break
			case 22:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 23:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 24:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break	
			case 1:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' OR INDICATOR_STATE='UT' OR INDICATOR_STATE='MT' OR INDICATOR_STATE='ID' OR INDICATOR_STATE='CO' OR INDICATOR_STATE='NM' OR INDICATOR_STATE='TX' OR INDICATOR_STATE='MN' OR INDICATOR_STATE='IL' OR INDICATOR_STATE='IN' OR INDICATOR_STATE='WI' OR INDICATOR_STATE='KY OR INDICATOR_STATE='TN' ORDER BY CREATE_DATE ASC";
				break
			case 2:
				recSet2.Filter="INDICATOR_STATE='CA' OR INDICATOR_STATE='OR' OR INDICATOR_STATE='WA' OR INDICATOR_STATE='AZ' OR INDICATOR_STATE='WY' ORDER BY CREATE_DATE ASC";
				break
			default:
				
				break
		}	
	
	*/
		
	ExitFlag = 'n';
	LastRecFlag = 'n';
	if(recSet2.BOF == 0)
	{
		recSet2.MoveFirst
	}
	
			
	switch(SwitchValue)
	{	
		case 'Work Locked Tickets':

			
			while(ExitFlag == 'n')
			{
				if(recSet2.EOF)
				{
					alert('All Tickets Complete')
					showStatus('All Tickets Completed!');
					document.frmmain.cmdSaveAndExit.style.visibility="hidden";
					document.frmmain.cmdsave.style.visibility="hidden";
					ExitFlag = 'y';
					LastRecFlag = 'y';
					info.cmd_WorkedLockedTickets.value = 'Work New Tickets';
					document.getElementById('lbl_Header').innerText = 'Locked Tickets';
				}
			    
				Workable = CheckWorkedDate()
				
				if(Workable == 'y')
				{
					if(recSet2.Fields('BEING_WORK_BY').value == null || recSet2.Fields('BEING_WORK_BY').value == 'z')
					{
						if(recSet2.Fields('TICKET_ID').value != 'na_peter')
						{
							recSet2.WillChangeRecord	
								recSet2.Fields('BEING_WORK_BY').value = EmpNum
								recSet2.Fields('ASSIGNED_TIME').value = Date()
								TICKETID = recSet2.Fields('TICKET_ID').value;
								document.frmmain.txt_RecordNum.value = recSet2.Fields('RecNum').value;
							recSet2.Update
							GetTicketInfo()
							UserAction = 'Assign Ticket';
							WriteToActionLog(UserAction, EmpNum, TICKETID)
							ExitFlag = 'y';
						}
					}
				}
				recSet2.MoveNext
			}
			
			break
		case 'Work New Tickets':

	   
			while(ExitFlag == 'n')
			{
				if(recSet2.EOF)
				{
					//alert('All Tickets Complete')
					info.cmd_WorkedLockedTickets.value = 'Work Locked Tickets';
					document.getElementById('lbl_Header').innerText = 'New Tickets';
					PullRecord()
					ExitFlag = 'y';
					LastRecFlag = 'y';
				}
				
				Workable = CheckWorkedDate()
				
				if(Workable == 'y')
				{
					if(recSet2.Fields('BEING_WORK_BY').value == EmpNum && recSet2.Fields('COMPLETED_BY').value == null)
					{
						if(recSet2.Fields('TICKET_ID').value != 'na_peter')
						{
							recSet2.WillChangeRecord
								recSet2.Fields('BEING_WORK_BY').value = EmpNum
								recSet2.Fields('ASSIGNED_TIME').value = Date()
								TICKETID = recSet2.Fields('TICKET_ID').value;
								document.frmmain.txt_RecordNum.value = recSet2.Fields('RecNum').value;
							recSet2.Update
							GetTicketInfo()
							ExitFlag = 'y';
						}
					}
				}
				recSet2.MoveNext
			}
			
			
			break
		default:
	}
	
	
	recSet2.close
	
	cObj.close
	GetUserStats()

//	validate_cmb_TicketStatus()
}


	function CheckWorkedDate()
	{
		var Workable;
		var TimeFlag;
		var DateFlag;
		
		var myDay;
		var myMonth;
		var myYear;
		var TodaysDate;
		var TodaysTime;
		var TodaysMinutes;
		var TestValue
		
		var WorkedTime;
		var WorkedDate;
		var WaitTime;
		var ReleaseDate;
		
		var myDate=new Date();
	  
		myDay = myDate.getDate();
		myMonth = myDate.getMonth();
		myMonth = myMonth + 1;
		myYear = myDate.getYear();
		TodaysDate = myMonth + "/" + myDay + "/" + myYear;
		TodaysTime = myDate.getHours();
		TodaysMinutes = myDate.getMinutes()
		
		WorkedTime = recSet2.Fields('WORK_TIME').value;
		WorkedDate = recSet2.Fields('WORKED_DATE').value;
		Ind_State = recSet2.Fields('INDICATOR_STATE').value;
		ReleaseDate = recSet2.Fields('RELEASE_DATE').value;
		
		TimeFlag = 'n';
		DateFlag = 'n';
		Workable = 'n';
		
		if(ReleaseDate == null)
		{
			DateFlag = 'y';
		}
		
		if(recSet2.Fields('RELEASE_DATE').value == TodaysDate)
		{
			DateFlag = 'y';
		}
		
		if(DateFlag == 'y')
		{
			if(WorkedDate == TodaysDate)
			{
				WaitTime = TodaysTime - WorkedTime;
				
				if(WaitTime >= 8)
				{
					Workable = 'y';
					return Workable
				}
				
			}
			else
			{
				Workable = 'y';
				return Workable
			}
		}

	}

	function ViewOwnedLocked()
	{
		var EmpNum;
		var UserStats;

		EmpNum = document.frmmain.txt_EmployeeNum.value;
		
		window.grStr2="SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY='" + EmpNum + "' AND TASC_STATUS <> 'Completed'";
		
		OpenDatabase()
		OpenTable2()
		
		UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
		
		if(recSet2.BOF == 0)
		{
			recSet2.MoveFirst
		}
		
		
		UserStats.document.write('<html>\n');

		UserStats.document.write('	<head>\n');
		UserStats.document.write('		<!-- CSS STYLESHEET -->\n');
		UserStats.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
		UserStats.document.write('		<!-- JAVASCRIPT LINKS -->\n');
		UserStats.document.write('		<script type="text/javascript" src="T://TASC Members Folders//TASC Peter LeBlanc//default tickets//main.js"></script>\n');
		UserStats.document.write('		<style type="text/css">\n');
		UserStats.document.write('			body {\n');
		UserStats.document.write('				/*background-color: #D0D0D0;*/\n');
		UserStats.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
		UserStats.document.write('				background-repeat: repeat;\n');
		UserStats.document.write('				border-style: none;\n');
		UserStats.document.write('			}\n');
		UserStats.document.write('		</style>\n');
		UserStats.document.write('	</head>\n');

		UserStats.document.write('	<body>\n');

		UserStats.document.write('		<form name="frmagentsstats">');
		UserStats.document.write('			<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('			<tr style="color:#ffffff;" class="text_header">\n');
		UserStats.document.write('				<td>&nbsp;Locked Ticket for Agent#'+ EmpNum +'</td>\n');
		UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			<tr>\n');
		UserStats.document.write('				<td colspan="2">');
		UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('						<tr class="text_subheader">\n');
		UserStats.document.write('							<td>Ticket#</td>');
		UserStats.document.write('							<td>Status</td>');
		UserStats.document.write('							<td>Next Action</td>');
		UserStats.document.write('							<td>Comments</td>');
		UserStats.document.write('						</tr>\n');	
		//UserStats.document.write('						<tr class="text_normal">\n');


		while(recSet2.EOF == 0)
		{
			UserStats.document.write('						<tr class="text_normal">\n');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('TICKET_ID').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('TASC_STATUS').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('NEXT_ACTION').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('Resolution').value +'</td>');
			UserStats.document.write('						</tr>\n')
			recSet2.MoveNext	
		}
		
		//UserStats.document.write('						</tr>\n');
		UserStats.document.write('					</table>\n');
		UserStats.document.write('				</td>\n');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			</table>\n');
		UserStats.document.write('		</form>\n');

		UserStats.document.write('	</body>\n');

		UserStats.document.write('</html>');
		
		
		recSet2.close
		cObj.close
	}
	
	function ClearImports()
	{
	
		window.grStr2="SELECT * FROM defaultImport";
		
		OpenDatabase()
		OpenTable2()
		
		if(recSet2.RecordCount != 0)
		{
			if(recSet2.BOF == 0)
			{
				recSet2.MoveFirst
			}
				
			while(recSet2.EOF == 0)
			{
				recSet2.Delete
				recSet2.MoveNext
			}
		}
		alert("Import Table cleared")
	}

	function ClearOrphans()
	{
	
		window.grStr2="SELECT * FROM Orphans";
		
		OpenDatabase()
		OpenTable2()
		
		if(recSet2.RecordCount != 0)
		{
			if(recSet2.BOF == 0)
			{
				recSet2.MoveFirst
			}
				
			while(recSet2.EOF == 0)
			{
				recSet2.Delete
				recSet2.MoveNext
			}
		}
		alert("All Orphans Cleared")
	}

	function ViewOrphans()
	{
		var EmpNum;
		var UserStats;

		EmpNum = document.frmmain.txt_EmployeeNum.value;
		
		window.grStr2="SELECT * FROM Orphans";
		
		OpenDatabase()
		OpenTable2()
		
		UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
		
		if(recSet2.BOF == 0)
		{
			recSet2.MoveFirst
		}
		
		
		UserStats.document.write('<html>\n');

		UserStats.document.write('	<head>\n');
		UserStats.document.write('		<!-- CSS STYLESHEET -->\n');
		UserStats.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
		UserStats.document.write('		<!-- JAVASCRIPT LINKS -->\n');
		UserStats.document.write('		<script type="text/javascript" src="pub//js//main.js"></script>\n');
		UserStats.document.write('		<script type="text/javascript" src="pub//js//connection.js"></script>\n');
		UserStats.document.write('		<style type="text/css">\n');
		UserStats.document.write('			body {\n');
		UserStats.document.write('				/*background-color: #D0D0D0;*/\n');
		UserStats.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
		UserStats.document.write('				background-repeat: repeat;\n');
		UserStats.document.write('				border-style: none;\n');
		UserStats.document.write('			}\n');
		UserStats.document.write('		</style>\n');
		UserStats.document.write('	</head>\n');

		UserStats.document.write('	<body>\n');

		UserStats.document.write('		<form name="frmagentsstats">');
		UserStats.document.write('			<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('			<tr style="color:#ffffff;" class="text_header">\n');
		UserStats.document.write('				<td>Current Orphans</td>');
		UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			<tr>\n');
		UserStats.document.write('				<td colspan="2">');
		UserStats.document.write('					<table width="100%" cellpadding="1" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('						<tr class="text_subheader">\n');
		UserStats.document.write('							<td>Ticket ID</td>');
		UserStats.document.write('							<td>TASC Status</td>');
		UserStats.document.write('							<td>COMPLETED_BY</td>');
		UserStats.document.write('							<td>Comments</td>');
		UserStats.document.write('							<td>&nbsp;</td>');
		UserStats.document.write('						</tr>\n');	
		
		while(recSet2.EOF == 0)
		{
			UserStats.document.write('						<tr class="text_normal">\n');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('TICKET_ID').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('TASC_STATUS').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('COMPLETED_BY').value +'</td>');
			UserStats.document.write('							<td>&nbsp;'+ recSet2.Fields('Resolution').value +'</td>');
			UserStats.document.write('							<td><input id="cmd_releaseorphan_' + recSet2.Fields('TICKET_ID').value + '" name="cmd_releaseorphan_' + recSet2.Fields('TICKET_ID').value + '" type="button" value="ReWork" onClick="ReworkOrphan(\'' +recSet2.Fields('TICKET_ID').value + '\')" class="text_subheader text_normal text_buttons"><input id="cmd_closeorphan_' + recSet2.Fields('TICKET_ID').value + '" name="cmd_closeorphan_' + recSet2.Fields('TICKET_ID').value + '" type="button" value="Close" onClick="CloseOrphan(\'' +recSet2.Fields('TICKET_ID').value + '\')" class="text_subheader text_normal text_buttons"></td>');
			UserStats.document.write('						</tr>\n');		
			recSet2.MoveNext	
		}
		
		UserStats.document.write('					</table>\n');
		UserStats.document.write('				</td>\n');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			</table>\n');
		UserStats.document.write('		</form>\n');

		UserStats.document.write('	</body>\n');

		UserStats.document.write('</html>');		
		
		recSet2.close
		cObj.close
	}

	function ReworkOrphan(TicketID)
	{
		
		window.grStr2="SELECT * FROM Orphans";
		window.grStr="SELECT * FROM TO_BE_WORKED";
		OpenDatabase()
		OpenTable2()
		OpenTable()
		recSet.AddNew
			recSet.Fields('ASSIGNED_DATE').value=recSet2.Fields('ASSIGNED_DATE').value  		//ASSIGNED_DATE
		//	recSet.Fields('ASSIGNED_GROUP').value=recSet2.Fields('ASSIGNED_GROUP').value			//ASSIGNED_GROUP
		//	recSet.Fields('ASSIGNED_TO').value=recSet2.Fields('ASSIGNED_TO').value			//ASSIGNED_TO
		//	recSet.Fields('AUTO_CORRELATE').value=recSet2.Fields('AUTO_CORRELATE').value			//AUTO_CORRELATE
		//	recSet.Fields('CATEGORY').value=recSet2.Fields('CATEGORY').value			//CATEGORY
		//	recSet.Fields('CORRELATION_TYPE').value=recSet2.Fields('CORRELATION_TYPE').value			//CORRELATION_TYPE
		//	recSet.Fields('CORRELATION_VALUE_1').value=recSet2.Fields('CORRELATION_VALUE_1').value			//CORRELATION_VALUE_1
		//	recSet.Fields('CORRELATION_VALUE_2').value=recSet2.Fields('CORRELATION_VALUE_2').value			//CORRELATION_VALUE_2
			recSet.Fields('CREATE_DATE').value=recSet2.Fields('CREATE_DATE').value			//CREATE_DATE
			//recSet.Fields('DEPARTMENT').value=recSet2.Fields('DEPARTMENT').value			//DEPARTMENT
		//	recSet.Fields('ETA').value=recSet2.Fields('ETA').value			//ETA
			recSet.Fields('INDICATOR_CITY').value=recSet2.Fields('INDICATOR_CITY').value			//INDICATOR_CITY
		//	recSet.Fields('INDICATOR_STATE').value=recSet2.Fields('INDICATOR_STATE').value			//INDICATOR_STATE
			//recSet.Fields('MultiRegion Affected Names').value=recSet2.Fields('MultiRegion Affected Names').value			//MultiRegion Affected Names
			//recSet.Fields('MultiRegion Affected').value=recSet2.Fields('MultiRegion Affected').value			//MultiRegion Affected
			//recSet.Fields('MultiRegion Flag').value=recSet2.Fields('MultiRegion Flag').value			//MultiRegion Flag
		//	recSet.Fields('NETWORK_ELEMENT').value=recSet2.Fields('NETWORK_ELEMENT').value			//NETWORK_ELEMENT
			recSet.Fields('NEXT_ACTION').value=recSet2.Fields('NEXT_ACTION').value			//NEXT_ACTION
		//	recSet.Fields('NIU ID').value=recSet2.Fields('NIU ID').value			//NIU ID
			recSet.Fields('Node').value=recSet2.Fields('Node').value			//Node
		//	recSet.Fields('OFFICE_NAME').value=recSet2.Fields('OFFICE_NAME').value			//OFFICE_NAME
		//	recSet.Fields('PARENT_TICKET').value=recSet2.Fields('PARENT_TICKET').value			//PARENT_TICKET
			recSet.Fields('PRIMARY_PHONE').value=recSet2.Fields('PRIMARY_PHONE').value			//PRIMARY_PHONE
		//	recSet.Fields('PRIORITY').value=recSet2.Fields('PRIORITY').value			//PRIORITY
		//	recSet.Fields('PROBLEM_CODE').value=recSet2.Fields('PROBLEM_CODE').value			//PROBLEM_CODE
			recSet.Fields('PROBLEM_SUMMARY').value=recSet2.Fields('PROBLEM_SUMMARY').value			//PROBLEM_SUMMARY
			//recSet.Fields('PRODUCT').value=recSet2.Fields('PRODUCT').value			//PRODUCT
		//	recSet.Fields('PURGE_DATE').value=recSet2.Fields('PURGE_DATE').value			//PURGE_DATE
		//	recSet.Fields('QUEUE_ID').value=recSet2.Fields('QUEUE_ID').value			//QUEUE_ID
			//recSet.Fields('QUEUE_NAME').value=recSet2.Fields('QUEUE_NAME').value			//QUEUE_NAME
			//recSet.Fields('REGION_ID').value=recSet2.Fields('REGION_ID').value			//REGION_ID
			recSet.Fields('REGION_NAME').value=recSet2.Fields('REGION_NAME').value			//REGION_NAME
		//	recSet.Fields('SA_ID').value=recSet2.Fields('SA_ID').value			//SA_ID
			//recSet.Fields('SERVICE_AFFECT').value=recSet2.Fields('SERVICE_AFFECT').value			//SERVICE_AFFECT
		//	recSet.Fields('Severity').value=recSet2.Fields('Severity').value			//Severity
			//recSet.Fields('STATUS').value=recSet2.Fields('STATUS').value			//STATUS
			recSet.Fields('SUBMITTER').value=recSet2.Fields('SUBMITTER').value			//SUBMITTER
			//recSet.Fields('SYSTEM_ID').value=recSet2.Fields('SYSTEM_ID').value			//SYSTEM_ID
		//	recSet.Fields('SYSTEM_NAME').value=recSet2.Fields('SYSTEM_NAME').value			//SYSTEM_NAME
		//	recSet.Fields('TECH_ARRIVAL_TIME').value=recSet2.Fields('TECH_ARRIVAL_TIME').value			//TECH_ARRIVAL_TIME
		//	recSet.Fields('TECHNICIAN_NAME').value=recSet2.Fields('TECHNICIAN_NAME').value			//TECHNICIAN_NAME
		//	recSet.Fields('TECHNICIAN_NUMBER').value=recSet2.Fields('TECHNICIAN_NUMBER').value			//TECHNICIAN_NUMBER
			recSet.Fields('TICKET_ID').value=recSet2.Fields('TICKET_ID').value			//TICKET_ID
		//	recSet.Fields('TTR_SLA').value=recSet2.Fields('TTR_SLA').value			//TTR_SLA
		//	recSet.Fields('TTR_START').value=recSet2.Fields('TTR_START').value			//TTR_START
			//recSet.Fields('WF-Scratch1').value=recSet2.Fields('WF-Scratch1').value			//WF-Scratch1
		   
			recSet.Fields('WORKED_DATE').value = recSet2.Fields('WORKED_DATE').value
			recSet.Fields('REASSIGNMENT_QUEUE').value = recSet2.Fields('REASSIGNMENT_QUEUE').value
			//recSet.Fields('INVALID_TICKET').value = recSet2.Fields('INVALID_TICKET').value
			//recSet.Fields('INVALID_PROBLEM').value = recSet2.Fields('INVALID_PROBLEM').value 
			//recSet.Fields('CALLBACK_REASON').value = recSet2.Fields('CALLBACK_REASON').value
			//recSet.Fields('TASC_STATUS').value = recSet2.Fields('TASC_STATUS').value
			//recSet.Fields('REASSIGNED_TO').value = recSet2.Fields('REASSIGNED_TO').value
			
			//recSet.Fields('BEING_WORK_BY').value = recSet2.Fields('BEING_WORK_BY').value 
			//recSet.Fields('ASSIGNED_TIME').value = recSet2.Fields('ASSIGNED_TIME').value
			//recSet.Fields('COMPLETED_TIME').value = recSet2.Fields('COMPLETED_TIME').value
			//recSet.Fields('COMPLETED_BY').value = recSet2.Fields('COMPLETED_BY').value
					
			//recSet.Fields('CALLBACK_DONE').value = recSet2.Fields('CALLBACK_DONE').value
			//recSet.Fields('CALLBACK_REASON').value = recSet2.Fields('CALLBACK_REASON').value
			recSet.Fields('Resolution').value = recSet2.Fields('Resolution').value 
		recSet.Update
		
		recSet2.close
		cObj.close
		
	}
	
	function CloseOrphan(TicketID)
	{
		window.grStr2="SELECT * FROM Orphans WHERE TICKET_ID='" + TicketID + "'";
		
		OpenDatabase()
		OpenTable2()
		recSet2.Delete
		recSet2.Update
		recSet2.close
		cObj.close
		alert('Cleared')
	}

	function CheckAllowedCharsSave() 
	{
		maxchars=255;
		
		if(document.frmmain.txt_ResolutionNotes.value.length > maxchars) 
		{
			alert('Too much data in the text box! Please remove '+
			(document.frmmain.txt_ResolutionNotes.value.length - maxchars)+ ' characters');
			return false;
		}
		else
		{
			return true;
		}
	}

	function CheckAllowedChars()
	{
		iCurChars = document.frmmain.txt_ResolutionNotes.value.length;
		document.getElementById('lbl_AllowedChars').innerText = iCurChars++;
	}

	function addOption(selectObject,optionText,optionValue) {
		var optionObject = new Option(optionText,optionValue)
		var optionRank = selectObject.options.length
		selectObject.options[optionRank]=optionObject
		
		//Sample Usage:
		//addOption(phonenumentry.lstpnephonelog, full_phonenumber_value, full_phonenumber_text);
	}

	function showStatus(sMsg) 
	{
		window.status = sMsg;
		document.getElementById('lbl_formstatus').innerText = sMsg;
	}

	function SetPageControls()
	{
	   div_main.style.visibility="hidden";
	}


	function FillRecord()
	{
		//Section: Import Ticket Value 
		
		var info = document.frmmain;
		
		info.txt_NumofContacts.value = recSet.Fields('NumofContacts').value;
		info.txt_TicketID.value = recSet.Fields('TICKET_ID').value;
		info.txt_IndicatorCity.value = recSet.Fields('INDICATOR_CITY').value;
		info.txt_RegionName.value = recSet.Fields('REGION_NAME').value;
		info.txt_PrimaryPhone.value = recSet.Fields('PRIMARY_PHONE').value;
		info.txt_AssignedDate.value = recSet.Fields('ASSIGNED_DATE').value;
		info.txt_CreateDate.value = recSet.Fields('CREATE_DATE').value;
		info.txt_Node.value = recSet.Fields('Node').value;
		info.txt_WorkedDate.value = recSet.Fields('Worked Date').value;
		info.txt_ProblemSummary.value = recSet.Fields('PROBLEM_SUMMARY').value;
		info.txt_NextAction.value = recSet.Fields('NEXT_ACTION').value;
		info.txt_TTSID.value = recSet.Fields('SUBMITTER').value;  //TTS Submitter
		info.cmb_TicketStatus.value = recSet.Fields('TASC_Status').value;  //User Information
		
		

	}

	function UpdateOldDefaultTable()
	{
		
	   var RecordCount;
	   var TicketID;
	   var UserAction;
	   var EmpNum
	   var DupFlag;
	   var ExitFlag;
	   var i;
	   var ii;
	   var NumOfTickets;

		

		window.grStr="SELECT TOP 20 * FROM tab_default WHERE TASC_Status <> 'Completed' AND TASC_Status <> 'Re-Assigned to Other'" ;      
	   window.grStr2="SELECT * FROM Completed";
	   
	   
	   OpenDatabase()

	   //OpenTable()
	   OpenTable2()

	   i = 0;
	   RecordCount = 0;
	   NumOfTickets = 0;

	   recSet2.MoveFirst;
	   recSet2.MoveNext;
		
		while(recSet2.EOF == 0)
		{
			   TicketID = recSet2.Fields('TICKET_ID').value;
			   
				 
			   window.grStr="SELECT * FROM tab_default WHERE TICKET_ID='" + TicketID + "'";
			   
			   OpenTable()
			   
			   recSet.WillChangeRecord
					recSet.Fields('TICKET_ID').value = recSet2.Fields('TICKET_ID').value
					recSet.Fields('WORKED DATE').value = recSet2.Fields('WORKED_DATE').value
					recSet.Fields('INVALID TICKET').value = recSet2.Fields('INVALID_TICKET').value
					recSet.Fields('INVALID PROBLEM').value = recSet2.Fields('INVALID_PROBLEM').value 
					recSet.Fields('CALLBACK REASON').value = recSet2.Fields('CALLBACK_REASON').value
					recSet.Fields('PROBLEM_SUMMARY').value = recSet2.Fields('PROBLEM_SUMMARY').value
					recSet.Fields('RESOLUTION').value = recSet2.Fields('RESOLUTION').value
					recSet.Fields('TASC_STATUS').value = recSet2.Fields('TASC_STATUS').value
					recSet.Fields('NEXT_ACTION').value = recSet2.Fields('NEXT_ACTION').value 
					recSet.Fields('Node').value = recSet2.Fields('Node').value
					recSet.Fields('Re-assignment Queue').value = recSet2.Fields('REASSIGNED_TO').value
					recSet.Fields('TASC_Status').value = recSet2.Fields('TASC_Status').value	
					recSet.Fields('CALLBACK REASON').value = recSet2.Fields('CALLBACK_REASON').value
					recSet.Fields('Resolution').value = recSet2.Fields('Resolution').value 
			   recSet.Update
				recSet.close
			recSet2.MoveNext;	   
			
		}	   
		   
	   
	   recSet2.close
	   cObj.close

	}

	function UpdateCompletedTable()
	{

		var MissingData;
		var RecNum;
		var UserAction;
		var EmpNum;
		var TICKETID;
		var TicketStatus;
		var Comments;
		var RecCount;
		
		alert('Starting Updates')
		
		window.grStr2="SELECT * FROM TO_BE_WORKED";
		window.grStr="SELECT * FROM Completed";
		RecCount = 0;
		
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = document.frmmain.txt_TicketID.value; 
		Comments = document.frmmain.txt_ResolutionNotes.value;
		TicketStatus = document.frmmain.cmb_TicketStatus.value;
		

		OpenDatabase()
		OpenTable2()
		OpenTable()
		
		if(recSet2.BOF == 0)
		{
			recSet2.MoveFirst
		}
		
		alert('Removing Completed Tickets')
		
		while(recSet2.EOF == 0)
		{
			if(recSet2.Fields('COMPLETED_TIME').value != null)
			{
				RecCount++
				recSet.AddNew;

					recSet.Fields('ASSIGNED_DATE').value=recSet2.Fields('ASSIGNED_DATE').value  		//ASSIGNED_DATE
					recSet.Fields('ASSIGNED_GROUP').value=recSet2.Fields('ASSIGNED_GROUP').value			//ASSIGNED_GROUP
					recSet.Fields('ASSIGNED_TO').value=recSet2.Fields('ASSIGNED_TO').value			//ASSIGNED_TO
					recSet.Fields('AUTO_CORRELATE').value=recSet2.Fields('AUTO_CORRELATE').value			//AUTO_CORRELATE
					recSet.Fields('CATEGORY').value=recSet2.Fields('CATEGORY').value			//CATEGORY
					recSet.Fields('CORRELATION_TYPE').value=recSet2.Fields('CORRELATION_TYPE').value			//CORRELATION_TYPE
					recSet.Fields('CORRELATION_VALUE_1').value=recSet2.Fields('CORRELATION_VALUE_1').value			//CORRELATION_VALUE_1
					recSet.Fields('CORRELATION_VALUE_2').value=recSet2.Fields('CORRELATION_VALUE_2').value			//CORRELATION_VALUE_2
					recSet.Fields('CREATE_DATE').value=recSet2.Fields('CREATE_DATE').value			//CREATE_DATE
					//recSet.Fields('DEPARTMENT').value=recSet2.Fields('DEPARTMENT').value			//DEPARTMENT
					recSet.Fields('ETA').value=recSet2.Fields('ETA').value			//ETA
					recSet.Fields('INDICATOR_CITY').value=recSet2.Fields('INDICATOR_CITY').value			//INDICATOR_CITY
					recSet.Fields('INDICATOR_STATE').value=recSet2.Fields('INDICATOR_STATE').value			//INDICATOR_STATE
					//recSet.Fields('MultiRegion Affected Names').value=recSet2.Fields('MultiRegion Affected Names').value			//MultiRegion Affected Names
					//recSet.Fields('MultiRegion Affected').value=recSet2.Fields('MultiRegion Affected').value			//MultiRegion Affected
					//recSet.Fields('MultiRegion Flag').value=recSet2.Fields('MultiRegion Flag').value			//MultiRegion Flag
					recSet.Fields('NETWORK_ELEMENT').value=recSet2.Fields('NETWORK_ELEMENT').value			//NETWORK_ELEMENT
					recSet.Fields('NEXT_ACTION').value=recSet2.Fields('NEXT_ACTION').value			//NEXT_ACTION
					recSet.Fields('NIU ID').value=recSet2.Fields('NIU ID').value			//NIU ID
					recSet.Fields('Node').value=recSet2.Fields('Node').value			//Node
					recSet.Fields('OFFICE_NAME').value=recSet2.Fields('OFFICE_NAME').value			//OFFICE_NAME
					recSet.Fields('PARENT_TICKET').value=recSet2.Fields('PARENT_TICKET').value			//PARENT_TICKET
					recSet.Fields('PRIMARY_PHONE').value=recSet2.Fields('PRIMARY_PHONE').value			//PRIMARY_PHONE
					recSet.Fields('PRIORITY').value=recSet2.Fields('PRIORITY').value			//PRIORITY
					recSet.Fields('PROBLEM_CODE').value=recSet2.Fields('PROBLEM_CODE').value			//PROBLEM_CODE
					recSet.Fields('PROBLEM_SUMMARY').value=recSet2.Fields('PROBLEM_SUMMARY').value			//PROBLEM_SUMMARY
					//recSet.Fields('PRODUCT').value=recSet2.Fields('PRODUCT').value			//PRODUCT
					recSet.Fields('PURGE_DATE').value=recSet2.Fields('PURGE_DATE').value			//PURGE_DATE
					recSet.Fields('QUEUE_ID').value=recSet2.Fields('QUEUE_ID').value			//QUEUE_ID
					//recSet.Fields('QUEUE_NAME').value=recSet2.Fields('QUEUE_NAME').value			//QUEUE_NAME
					//recSet.Fields('REGION_ID').value=recSet2.Fields('REGION_ID').value			//REGION_ID
					recSet.Fields('REGION_NAME').value=recSet2.Fields('REGION_NAME').value			//REGION_NAME
					recSet.Fields('SA_ID').value=recSet2.Fields('SA_ID').value			//SA_ID
					//recSet.Fields('SERVICE_AFFECT').value=recSet2.Fields('SERVICE_AFFECT').value			//SERVICE_AFFECT
					recSet.Fields('Severity').value=recSet2.Fields('Severity').value			//Severity
					//recSet.Fields('STATUS').value=recSet2.Fields('STATUS').value			//STATUS
					recSet.Fields('SUBMITTER').value=recSet2.Fields('SUBMITTER').value			//SUBMITTER
					//recSet.Fields('SYSTEM_ID').value=recSet2.Fields('SYSTEM_ID').value			//SYSTEM_ID
					recSet.Fields('SYSTEM_NAME').value=recSet2.Fields('SYSTEM_NAME').value			//SYSTEM_NAME
					recSet.Fields('TECH_ARRIVAL_TIME').value=recSet2.Fields('TECH_ARRIVAL_TIME').value			//TECH_ARRIVAL_TIME
					recSet.Fields('TECHNICIAN_NAME').value=recSet2.Fields('TECHNICIAN_NAME').value			//TECHNICIAN_NAME
					recSet.Fields('TECHNICIAN_NUMBER').value=recSet2.Fields('TECHNICIAN_NUMBER').value			//TECHNICIAN_NUMBER
					recSet.Fields('TICKET_ID').value=recSet2.Fields('TICKET_ID').value			//TICKET_ID
					recSet.Fields('TTR_SLA').value=recSet2.Fields('TTR_SLA').value			//TTR_SLA
					recSet.Fields('TTR_START').value=recSet2.Fields('TTR_START').value			//TTR_START
					//recSet.Fields('WF-Scratch1').value=recSet2.Fields('WF-Scratch1').value			//WF-Scratch1
				   
					recSet.Fields('WORKED_DATE').value = recSet2.Fields('WORKED_DATE').value
					recSet.Fields('REASSIGNMENT_QUEUE').value = recSet2.Fields('REASSIGNMENT_QUEUE').value
					recSet.Fields('INVALID_TICKET').value = recSet2.Fields('INVALID_TICKET').value
					recSet.Fields('INVALID_PROBLEM').value = recSet2.Fields('INVALID_PROBLEM').value 
					recSet.Fields('CALLBACK_REASON').value = recSet2.Fields('CALLBACK_REASON').value
					recSet.Fields('TASC_STATUS').value = recSet2.Fields('TASC_STATUS').value
					recSet.Fields('REASSIGNED_TO').value = recSet2.Fields('REASSIGNED_TO').value
					
					recSet.Fields('BEING_WORK_BY').value = recSet2.Fields('BEING_WORK_BY').value 
					recSet.Fields('ASSIGNED_TIME').value = recSet2.Fields('ASSIGNED_TIME').value
					recSet.Fields('COMPLETED_TIME').value = recSet2.Fields('COMPLETED_TIME').value
					recSet.Fields('COMPLETED_BY').value = recSet2.Fields('COMPLETED_BY').value
							
					recSet.Fields('CALLBACK_DONE').value = recSet2.Fields('CALLBACK_DONE').value
					recSet.Fields('CALLBACK_REASON').value = recSet2.Fields('CALLBACK_REASON').value
					recSet.Fields('Resolution').value = recSet2.Fields('Resolution').value 
		   
				
				
				
					recSet2.Delete
					 recSet.Update;	
			}
			recSet2.MoveNext;
		}

		recSet2.close
		cObj.close

	   TICKETID = RecCount;
	   UserAction='UpdateComplete';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)

		alert('Completed Table Updated')
	}

	function GetTicketInfo()
	{
	   //  Populates form fields
	   
	   var info = document.frmmain;
		  
	   info.txt_TicketID.value = recSet2.Fields('TICKET_ID').value;
	   info.txt_IndicatorCity.value = recSet2.Fields('INDICATOR_CITY').value;
	   info.txt_RegionName.value = recSet2.Fields('REGION_NAME').value;
	   info.txt_PrimaryPhone.value = recSet2.Fields('PRIMARY_PHONE').value;
	   info.txt_Assigned_Date.value = recSet2.Fields('ASSIGNED_DATE').value;
	   info.txt_CreateDate.value = recSet2.Fields('CREATE_DATE').value;
		

	   info.txt_WorkedDate.value = recSet2.Fields('WORKED_DATE').value;
	   info.txt_ProblemSummary.value = recSet2.Fields('PROBLEM_SUMMARY').value;
	   info.txt_TTSID.value = recSet2.Fields('SUBMITTER').value;
	   info.txt_NextAction.value = recSet2.Fields('NEXT_ACTION').value;
	   
	   if(recSet2.Fields('NUMOFCONTACTS').value == null)
	   {
			info.txt_NumofContacts.value = 0;
		}
		else
		{
			info.txt_NumofContacts.value = recSet2.Fields('NumofContacts').value;
		}

		if (info.txt_Node.value = recSet2.Fields('Node').value == 'Null')
		{
			info.txt_Node.value = "Not Available";
		} 
		else 
		{
			info.txt_Node.value = recSet2.Fields('NODE').value;
		}
	   
	   //Users Information
	   info.cmb_TicketStatus.value = recSet2.Fields('TASC_Status').value;
	   info.txt_PreResolutionNotes.value = recSet2.Fields('Resolution').value;
	   
	   validate_cmb_TicketStatus()
	}


	function SyncToBeWork()
	{
		recSet2.AddNew;
		   recSet2.Fields('TICKET_ID').value = recSet.Fields('TICKET_ID').value
		   recSet2.Fields('INDICATOR_CITY').value = recSet.Fields('INDICATOR_CITY').value
		   recSet2.Fields('REGION_NAME').value = recSet.Fields('REGION_NAME').value
		   recSet2.Fields('PRIMARY_PHONE').value = recSet.Fields('PRIMARY_PHONE').value    
		   recSet2.Fields('ASSIGNED_DATE').value = recSet.Fields('ASSIGNED_DATE').value
		   recSet2.Fields('CREATE_DATE').value = recSet.Fields('CREATE_DATE').value
		   recSet2.Fields('WORKED_DATE').value = recSet.Fields('Worked Date').value
		   recSet2.Fields('REASSIGNMENT_QUEUE').value = recSet.Fields('Re-assignment Queue').value
		   recSet2.Fields('INVALID_TICKET').value = recSet.Fields('Invalid Ticket').value
		   recSet2.Fields('CALLBACK_REASON').value = recSet.Fields('callback Reason').value
		   recSet2.Fields('SUBMITTER').value = recSet.Fields('SUBMITTER').value
		   recSet2.Fields('PROBLEM_SUMMARY').value = recSet.Fields('PROBLEM_SUMMARY').value
		   recSet2.Fields('RESOLUTION').value = recSet.Fields('RESOLUTION').value
		   recSet2.Fields('TASC_STATUS').value = recSet.Fields('TASC_STATUS').value
		   recSet2.Fields('NEXT_ACTION').value = recSet.Fields('NEXT_ACTION').value 
		recSet2.Update;

	}

	function SaveAndExit()
	{
		showStatus('Saving and Exiting');
		var Comments;
		var TicketStatus;
		var CommentsPass;
		
		Comments = document.frmmain.txt_ResolutionNotes.value;
		TicketStatus = document.frmmain.cmb_TicketStatus.value;
		
		CommentsPass=CheckAllowedCharsSave()
		
		if(CommentsPass == true)
		{
			
			if(Comments != "" && TicketStatus != "" || TicketStatus == 'Never Worked')
			{
				SaveTicketChanges()
				document.frmmain.cmdSaveAndExit.style.visibility="hidden";
				document.frmmain.cmdsave.style.visibility="hidden";
				window.close()
			}
			else
			{
				alert("Must Have Comments and Ticket Status")
			}
		}
	}

	function SaveAndContinue()
	{
		var Comments;
		var TicketStatus;
		var AssignTo;
		var CallbackTo;
		var NextedAction;
		var CommentsPass;
		var info = document.frmmain;
		var UserAction;
		var EmpNum;
		var TICKETID;
		
		Comments = document.frmmain.txt_ResolutionNotes.value;
		TicketStatus = document.frmmain.cmb_TicketStatus.value;
		CallbackTo = document.frmmain.cmb_CallbackTo.value;
		AssignTo = document.frmmain.cmb_TicketReassignedTo.value;
		NextedAction = document.frmmain.cmb_NextAction.value;
		
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = 'Error Catch';
		CommentsPass=CheckAllowedCharsSave()
		
		if(CommentsPass == true)
		{

			switch(TicketStatus)
			{
				case 'Completed':
					if(Comments != '')
					{
						showStatus('Saving And Continue');
						SaveTicketChanges()
						try
						{
							PullRecord()
						}
						catch(err)
						{
							
						}
						
						
					}
					else
					{
						if(Comments == '')
							alert('You Must add a Comment')
					}
					break
				case 'Callback Needed':
					if(CallbackTo != 'n/a' && Comments != '' && NextedAction != 'n/a')
					{
						showStatus('Saving And Continue');
						SaveTicketChanges()
						try
						{
							PullRecord()
						}
						catch(err)
						{
							
						}
					}
					else
					{
						if(CallbackTo == 'n/a')
							alert('You must assign the ticket to someone')
						if(Comments == '')
							alert('You must add a Comment')
						if(NextedAction == 'n/a')
							alert('You must have Nexted Action')
							
						
					}
					break	
				case 'Reassigned':
					if(AssignTo != 'n/a' && AssignTo != 'To Self' && AssignTo != 'Back to Public Listings' && Comments != '')
					{
						showStatus('Saving And Continue');
						SaveTicketChanges()
						try
						{
							PullRecord()
						}
						catch(err)
						{
							
						}
					}
					else
					{
						showStatus(document.frmmain.cmb_CallbackTo.value);
						
						if(CallbackTo == 'n/a')
							alert('You must assign the ticket to someone')
						if(CallbackTo == 'To Self')
							alert('Can NOT Reassign a ticket to self')
						if(CallbackTo == 'Back to Public Listings')
							alert('Can NOT reassign ticket to public')
							
						if(Comments == '')
							alert('You must add a Comment')
						
					}
					break
				case 'Never Worked':
					showStatus('Saving And Continue');
					SaveTicketChanges()
						try
						{
							PullRecord()
						}
						catch(err)
						{
							
						}
					break
				default:
					alert("Must Have Ticket Status")
					break
			}
		}
	}

	function SyncSaveChanges()
	{

		var SwitchValue;
		var info = document.frmmain;
		var UserAction;
		var EmpNum;
		var TICKETID;
		
		var myDay;
		var myMonth;
		var myYear;
		var TodaysDate;
		var TodaysTime;

		var myDate=new Date();
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = document.frmmain.txt_TicketID.value;
	  
		myDay = myDate.getDate();
		myMonth = myDate.getMonth();
		myMonth = myMonth + 1;
		myYear = myDate.getYear();
		TodaysDate = myMonth + "/" + myDay + "/" + myYear;
		TodaysTime = myDate.getHours()
			
		SwitchValue = info.cmb_TicketStatus.value;
		
		recSet2.WillChangeRecord
		
		switch(SwitchValue)
		{
			case 'Completed':
			
				recSet2.Fields('NumofContacts').value = info.txt_NumofContacts.value;
				recSet2.Fields('CALLBACK_DONE').value = info.chk_CallbackDone.checked;
				recSet2.Fields('INVALID_TICKET').value = info.chk_InvalidTicket.checked;
				
				recSet2.Fields('TASC_Status').value = info.cmb_TicketStatus.value;

				recSet2.Fields('REASSIGNED_TO').value = info.cmb_TicketReassignedTo.value;
				recSet2.Fields('NEXT_ACTION').value = info.cmb_NextAction.value;
					
				recSet2.Fields('CALLBACK_REASON').value = info.cmd_CallbackReason.value;
				recSet2.Fields('INVALID_PROBLEM').value = info.cmb_Problem.value;
				recSet2.Fields('Resolution').value = info.txt_ResolutionNotes.value;
				recSet2.Fields('SUBMITTER').value = info.txt_TTSID.value;
			
				recSet2.Fields('WORKED_DATE').value = TodaysDate;
				recSet2.Fields('COMPLETED_BY').value = document.frmmain.txt_EmployeeNum.value;
				recSet2.Fields('COMPLETED_TIME').value = Date();
				
				UserAction = 'Completed';
				WriteToActionLog(UserAction, EmpNum, TICKETID)
				
				if(recSet2.Fields('CALLBACK_DONE').value == true)
				{
					UserAction = 'DidCallBack';
					WriteToActionLog(UserAction, EmpNum, TICKETID)
				}			
				break
				
			case 'Callback Needed':
				
				if(info.cmb_CallbackTo.value == 'Back to Public Listings')
				{
					recSet2.Fields('NumofContacts').value = info.txt_NumofContacts.value;
					recSet2.Fields('BEING_WORK_BY').value = 'z';
					recSet2.Fields('ASSIGNED_TIME').value = 'z';
					recSet2.Fields('NEXT_ACTION').value = info.cmb_NextAction.value;
					recSet2.Fields('TASC_Status').value = info.cmb_TicketStatus.value;	
					recSet2.Fields('CALLBACK_DONE').value = info.chk_CallbackDone.checked;
					recSet2.Fields('CALLBACK_REASON').value = info.cmd_CallbackReason.value;
					recSet2.Fields('Resolution').value = info.txt_ResolutionNotes.value;
					recSet2.Fields('WORKED_DATE').value = TodaysDate;
					recSet2.Fields('WORK_TIME').value = TodaysTime;
					
					if(recSet2.Fields('CALLBACK_DONE').value == false)
					{
						UserAction = 'ToPublic';
						WriteToActionLog(UserAction, EmpNum, TICKETID)
					}
					
					if(recSet2.Fields('CALLBACK_DONE').value == true)
					{
						UserAction = 'ToPublic';
						WriteToActionLog(UserAction, EmpNum, TICKETID)
						UserAction = 'DidCallBack';
						WriteToActionLog(UserAction, EmpNum, TICKETID)
					}
					
					
				}
				
				if(info.cmb_CallbackTo.value == 'To Self')
				{
					recSet2.Fields('NumofContacts').value = info.txt_NumofContacts.value;
					recSet2.Fields('TASC_Status').value = info.cmb_TicketStatus.value;
					recSet2.Fields('NEXT_ACTION').value = info.cmb_NextAction.value;
					recSet2.Fields('Resolution').value = info.txt_ResolutionNotes.value;
					recSet2.Fields('WORKED_DATE').value = TodaysDate;		
					recSet2.Fields('WORK_TIME').value = TodaysTime;
					
					UserAction = 'ToSelf';
					WriteToActionLog(UserAction, EmpNum, TICKETID)
					if(recSet2.Fields('CALLBACK_DONE').value == true)
					{
						UserAction = 'DidCallBack';
						WriteToActionLog(UserAction, EmpNum, TICKETID)
					}
				}
				
				break
				
			case 'Reassigned':
				
				if(info.cmb_TicketReassignedTo.value == 'Denver Support Desk')
				{
					recSet2.Fields('NumofContacts').value = info.txt_NumofContacts.value;
					recSet2.Fields('TASC_Status').value = 'Completed'
					recSet2.Fields('REASSIGNED_TO').value = info.cmb_TicketReassignedTo.value;
					recSet2.Fields('WORKED_DATE').value = TodaysDate;
					recSet2.Fields('COMPLETED_BY').value = document.frmmain.txt_EmployeeNum.value;
					recSet2.Fields('COMPLETED_TIME').value = Date();
					recSet2.Fields('INVALID_PROBLEM').value = info.cmb_Problem.value;
					recSet2.Fields('Resolution').value = info.txt_ResolutionNotes.value;
					recSet2.Fields('SUBMITTER').value = info.txt_TTSID.value;
									
					UserAction = 'Reassigned';
					WriteToActionLog(UserAction, EmpNum, TICKETID)
				}
				
				if(info.cmb_TicketReassignedTo.value == 'National IP')
				{
					recSet2.Fields('NumofContacts').value = info.txt_NumofContacts.value;
					recSet2.Fields('TASC_Status').value = 'Completed'
					recSet2.Fields('REASSIGNED_TO').value = info.cmb_TicketReassignedTo.value;
					recSet2.Fields('WORKED_DATE').value = TodaysDate;
					recSet2.Fields('COMPLETED_BY').value = document.frmmain.txt_EmployeeNum.value;
					recSet2.Fields('COMPLETED_TIME').value = Date();
					recSet2.Fields('INVALID_PROBLEM').value = info.cmb_Problem.value;
					recSet2.Fields('Resolution').value = info.txt_ResolutionNotes.value;
					recSet2.Fields('SUBMITTER').value = info.txt_TTSID.value;
					
					UserAction = 'Reassigned';
					WriteToActionLog(UserAction, EmpNum, TICKETID)
				}
				if(recSet2.Fields('CALLBACK_DONE').value == true)
				{
					UserAction = 'DidCallBack';
					WriteToActionLog(UserAction, EmpNum, TICKETID)
				}
				break
			case 'Never Worked':
				
				recSet2.Fields('BEING_WORK_BY').value = 'z';
				recSet2.Fields('ASSIGNED_TIME').value = 'z';
				recSet2.Fields('TASC_Status').value = 'Working';
				
				UserAction = 'NeverWorked';
				WriteToActionLog(UserAction, EmpNum, TICKETID)
				
				break
			default:
		}

		recSet2.Update
	}

	function SaveTicketChanges()
	{
		var MissingData
		var RecNum;
		var UserAction;
		var EmpNum;
		var TICKETID;
		var TicketStatus;
		var Comments;
		
		RecNum = document.frmmain.txt_RecordNum.value;   

		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = document.frmmain.txt_TicketID.value; 
		Comments = document.frmmain.txt_ResolutionNotes.value;
		TicketStatus = document.frmmain.cmb_TicketStatus.value;

		//window.grStr2="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TICKETID +"'";
		window.grStr2="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TICKETID + "'" + " AND BEING_WORK_BY='" + EmpNum + "'";
		

		OpenDatabase()
		OpenTable2()

	  
		SyncSaveChanges()
	   
		document.frmmain.txt_ResolutionNotes.value = "";
		
		if(document.frmmain.cmb_TicketStatus.value != "Never Worked")
		{
			WriteCommentHistory(EmpNum, TICKETID, Comments)	
		}
		
		recSet2.close
		cObj.close
		
		alert('Record Updated')
		showStatus('Record Updated');
	}

	function SyncOrphans()
	{
		recSet2.AddNew;
		   recSet2.Fields('TICKET_ID').value = recSet4.Fields('TICKET_ID').value
		   recSet2.Fields('INDICATOR_CITY').value = recSet4.Fields('INDICATOR_CITY').value
		   recSet2.Fields('REGION_NAME').value = recSet4.Fields('REGION_NAME').value
		   recSet2.Fields('PRIMARY_PHONE').value = recSet4.Fields('PRIMARY_PHONE').value    
		   recSet2.Fields('ASSIGNED_DATE').value = recSet4.Fields('ASSIGNED_DATE').value
		   recSet2.Fields('CREATE_DATE').value = recSet4.Fields('CREATE_DATE').value
		   recSet2.Fields('WORKED_DATE').value = recSet4.Fields('WORKED_DATE').value
		   recSet2.Fields('REASSIGNMENT_QUEUE').value = recSet4.Fields('REASSIGNMENT_QUEUE').value
		   recSet2.Fields('INVALID_TICKET').value = recSet4.Fields('INVALID_TICKET').value
		   recSet2.Fields('CALLBACK_REASON').value = recSet4.Fields('CALLBACK_REASON').value
		   recSet2.Fields('SUBMITTER').value = recSet4.Fields('SUBMITTER').value
		   recSet2.Fields('PROBLEM_SUMMARY').value = recSet4.Fields('PROBLEM_SUMMARY').value
		   recSet2.Fields('RESOLUTION').value = recSet4.Fields('RESOLUTION').value
		   recSet2.Fields('TASC_STATUS').value = recSet4.Fields('TASC_STATUS').value
		   recSet2.Fields('NEXT_ACTION').value = recSet4.Fields('NEXT_ACTION').value 
		recSet2.Update;
	}

	function NewSyncToBeWork()
	{
		recSet2.AddNew;
	/*
		   recSet2.Fields('TICKET_ID').value = recSet.Fields('TICKET_ID').value
		   recSet2.Fields('INDICATOR_CITY').value = recSet.Fields('INDICATOR_CITY').value
		   recSet2.Fields('REGION_NAME').value = recSet.Fields('REGION_NAME').value
		   recSet2.Fields('PRIMARY_PHONE').value = recSet.Fields('PRIMARY_PHONE').value    
		   recSet2.Fields('ASSIGNED_DATE').value = recSet.Fields('ASSIGNED_DATE').value
		   recSet2.Fields('CREATE_DATE').value = recSet.Fields('CREATE_DATE').value
		   recSet2.Fields('WORKED_DATE').value = recSet.Fields('Worked Date').value
		   recSet2.Fields('REASSIGNMENT_QUEUE').value = recSet.Fields('Re-assignment Queue').value
		   recSet2.Fields('INVALID_TICKET').value = recSet.Fields('Invalid Ticket').value
		   recSet2.Fields('CALLBACK_REASON').value = recSet.Fields('callback Reason').value
		   recSet2.Fields('SUBMITTER').value = recSet.Fields('SUBMITTER').value
		   recSet2.Fields('PROBLEM_SUMMARY').value = recSet.Fields('PROBLEM_SUMMARY').value
		   recSet2.Fields('RESOLUTION').value = recSet.Fields('RESOLUTION').value
		   recSet2.Fields('TASC_STATUS').value = recSet.Fields('TASC STATUS').value
		   recSet2.Fields('NEXT_ACTION').value = recSet.Fields('NEXT_ACTION').value 
		*/
		   	recSet2.Fields('ASSIGNED_DATE').value=recSet.Fields('ASSIGNED_DATE').value  		//ASSIGNED_DATE
			recSet2.Fields('ASSIGNED_GROUP').value=recSet.Fields('ASSIGNED_GROUP').value			//ASSIGNED_GROUP
			recSet2.Fields('ASSIGNED_TO').value=recSet.Fields('ASSIGNED_TO').value			//ASSIGNED_TO
			recSet2.Fields('AUTO_CORRELATE').value=recSet.Fields('AUTO_CORRELATE').value			//AUTO_CORRELATE
			recSet2.Fields('CATEGORY').value=recSet.Fields('CATEGORY').value			//CATEGORY
			recSet2.Fields('CORRELATION_TYPE').value=recSet.Fields('CORRELATION_TYPE').value			//CORRELATION_TYPE
			recSet2.Fields('CORRELATION_VALUE_1').value=recSet.Fields('CORRELATION_VALUE_1').value			//CORRELATION_VALUE_1
			recSet2.Fields('CORRELATION_VALUE_2').value=recSet.Fields('CORRELATION_VALUE_2').value			//CORRELATION_VALUE_2
			recSet2.Fields('CREATE_DATE').value=recSet.Fields('CREATE_DATE').value			//CREATE_DATE
			//recSet2.Fields('DEPARTMENT').value=recSet.Fields('DEPARTMENT').value			//DEPARTMENT
			recSet2.Fields('ETA').value=recSet.Fields('ETA').value			//ETA
			recSet2.Fields('INDICATOR_CITY').value=recSet.Fields('INDICATOR_CITY').value			//INDICATOR_CITY
			recSet2.Fields('INDICATOR_STATE').value=recSet.Fields('INDICATOR_STATE').value			//INDICATOR_STATE
			//recSet2.Fields('MultiRegion Affected Names').value=recSet.Fields('MultiRegion Affected Names').value			//MultiRegion Affected Names
			//recSet2.Fields('MultiRegion Affected').value=recSet.Fields('MultiRegion Affected').value			//MultiRegion Affected
			//recSet2.Fields('MultiRegion Flag').value=recSet.Fields('MultiRegion Flag').value			//MultiRegion Flag
			recSet2.Fields('NETWORK_ELEMENT').value=recSet.Fields(16).value			//NETWORK_ELEMENT
			recSet2.Fields('NEXT_ACTION').value=recSet.Fields('NETWORK_ELEMENT').value			//NEXT_ACTION
			recSet2.Fields('NIU ID').value=recSet.Fields('NIU ID').value			//NIU ID
			recSet2.Fields('Node').value=recSet.Fields('Node').value			//Node
			recSet2.Fields('OFFICE_NAME').value=recSet.Fields('OFFICE_NAME').value			//OFFICE_NAME
			recSet2.Fields('PARENT_TICKET').value=recSet.Fields('PARENT_TICKET').value			//PARENT_TICKET
			recSet2.Fields('PRIMARY_PHONE').value=recSet.Fields('PRIMARY_PHONE').value			//PRIMARY_PHONE
			recSet2.Fields('PRIORITY').value=recSet.Fields('PRIORITY').value			//PRIORITY
			recSet2.Fields('PROBLEM_CODE').value=recSet.Fields('PROBLEM_CODE').value			//PROBLEM_CODE
			recSet2.Fields('PROBLEM_SUMMARY').value=recSet.Fields('PROBLEM_SUMMARY').value			//PROBLEM_SUMMARY
			//recSet2.Fields('PRODUCT').value=recSet.Fields('PRODUCT').value			//PRODUCT
			recSet2.Fields('PURGE_DATE').value=recSet.Fields('PURGE_DATE').value			//PURGE_DATE
			recSet2.Fields('QUEUE_ID').value=recSet.Fields('QUEUE_ID').value			//QUEUE_ID
			//recSet2.Fields('QUEUE_NAME').value=recSet.Fields('QUEUE_NAME').value			//QUEUE_NAME
			//recSet2.Fields('REGION_ID').value=recSet.Fields('REGION_ID').value			//REGION_ID
			recSet2.Fields('REGION_NAME').value=recSet.Fields('REGION_NAME').value			//REGION_NAME
			recSet2.Fields('SA_ID').value=recSet.Fields('SA_ID').value			//SA_ID
			//recSet2.Fields('SERVICE_AFFECT').value=recSet.Fields('SERVICE_AFFECT').value			//SERVICE_AFFECT
			recSet2.Fields('Severity').value=recSet.Fields('Severity').value			//Severity
			//recSet2.Fields('STATUS').value=recSet.Fields('STATUS').value			//STATUS
			recSet2.Fields('SUBMITTER').value=recSet.Fields('SUBMITTER').value			//SUBMITTER
			//recSet2.Fields('SYSTEM_ID').value=recSet.Fields('SYSTEM_ID').value			//SYSTEM_ID
			recSet2.Fields('SYSTEM_NAME').value=recSet.Fields('SYSTEM_NAME').value			//SYSTEM_NAME
			recSet2.Fields('TECH_ARRIVAL_TIME').value=recSet.Fields('TECH_ARRIVAL_TIME').value			//TECH_ARRIVAL_TIME
			recSet2.Fields('TECHNICIAN_NAME').value=recSet.Fields('TECHNICIAN_NAME').value			//TECHNICIAN_NAME
			recSet2.Fields('TECHNICIAN_NUMBER').value=recSet.Fields('TECHNICIAN_NUMBER').value			//TECHNICIAN_NUMBER
			recSet2.Fields('TICKET_ID').value=recSet.Fields('TICKET_ID').value			//TICKET_ID
			recSet2.Fields('TTR_SLA').value=recSet.Fields('TTR_SLA').value			//TTR_SLA
			recSet2.Fields('TTR_START').value=recSet.Fields('TTR_START').value			//TTR_START
			//recSet2.Fields('WF-Scratch1').value=recSet.Fields('WF-Scratch1').value			//WF-Scratch1
		   
		   
		recSet2.Update;
	}

/* **************** */
/* IMPORT FUNCTIONS */
/* **************** */
	function ImportData()
	{
		alert('Importing CSV File')
		ClearImports()
		ImportCSV()
		alert('Updating completed table')
		UpdateCompletedTable()
		
		var i;
		var TicketID;
		var FoundFlag;
		var StillOpenFlag;
		var DenverBounceBack;
		var OrphansCount;
		var WorkingCount;
		var DoubleCount;
		var DenverCount;
		
		
		alert('Starting Import')
		window.grStr="SELECT * FROM defaultImport";
		window.grStr2="SELECT * FROM TO_BE_WORKED";
		
		OpenDatabase()
		OpenTable()
		
		FoundFlag='n';
		StillOpenFlag='n';
		DenverBounceBack='n';
		
		OrphansCount=0;
		WorkingCount=0;
		DoubleCount=0;
		DenverCount=0;
		
		
		if(recSet.BOF == 0)
		{
			recSet.MoveFirst
		}
		
		while(recSet.EOF == 0)
		{
		
			FoundFlag='n';
			StillOpenFlag='n';
			DenverBounceBack='n';
			
			TicketID = recSet.Fields('TICKET_ID').value;
				
			window.grStr2="SELECT * FROM Completed WHERE TICKET_ID='" + TicketID +"'";
			OpenTable2();
	
			if(recSet2.RecordCount==0)
			{
				FoundFlag='n';
				
			}
			else
			{
				FoundFlag='y';
				recSet2.MoveFirst
				for(i=0; i<recSet2.RecordCount; i++)
				{
					if(recSet2.Fields('REASSIGNED_TO').value!='n/a')
					{
						DenverBounceBack='y';
					}
					else
					{
						DenverBounceBack='n';
					}
					recSet2.MoveNext;
				}
				
			}
			recSet2.close
			
			window.grStr2="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TicketID +"'";
			OpenTable2();
			
			if(recSet2.RecordCount==0)
			{
				StillOpenFlag='n';
				
			}	
			else
			{
				StillOpenFlag='y';
				
			}
			recSet2.close
			
					
			if(FoundFlag=='n' && StillOpenFlag=='n' && DenverBounceBack=='n')
			{
				window.grStr2="SELECT * FROM TO_BE_WORKED";
				WorkingCount++;
				OpenTable2();
				NewSyncToBeWork()
				recSet2.close
				recSet.Delete
			}
			
			if(FoundFlag=='y' && StillOpenFlag=='n' && DenverBounceBack=='y')
			{
				window.grStr2="SELECT * FROM TO_BE_WORKED";
				DenverCount++;
				OpenTable2();
				NewSyncToBeWork()
				recSet2.close
				recSet.Delete
			}
			
			
			if(FoundFlag=='y' && StillOpenFlag=='n' && DenverBounceBack=='n')
			{
				window.grStr4="SELECT * FROM Completed WHERE TICKET_ID='" + TicketID + "'";
				window.grStr2="SELECT * FROM Orphans";
				OrphansCount++;
				OpenTable2()
				OpenTable4()
				SyncOrphans()
				recSet4.close
				recSet2.close
				recSet.Delete
			}
			
			if(FoundFlag=='n' && StillOpenFlag=='y' && DenverBounceBack=='n')
			{
				DoubleCount++;
				recSet.Delete
			}
			recSet.MoveNext
			
		}
		
		recSet.close
		cObj.close
		

	   TICKETID = WorkingCount;
	   UserAction='WorkingCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = OrphansCount;
	   UserAction='OrphansCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = DenverCount;
	   UserAction='DenverCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = DoubleCount;
	   UserAction='DoubleCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)   
		
		alert('Import Complete')
	}


	function xxImportData()
	{


	   var RecordCount;
	   var TICKETID;
	   var UserAction;
	   var EmpNum
	   var DupFlag;
	   var ExitFlag;
	   var i;
	   var ii;
	   var NumOfTickets;

		// TOP 20 
		window.grStr="SELECT * FROM tab_default WHERE TASC_Status IS NULL OR TASC_Status <> 'Completed' AND TASC_Status <> 'Re-Assigned to Other'";
	   window.grStr2="SELECT * FROM TO_BE_WORKED";

	alert(grStr)

	   OpenDatabase()

	   OpenTable()
	   OpenTable2()

	   i = 0;
	   RecordCount = 0;
	   NumOfTickets = 0;
	   
	   alert(recSet.RecordCount)

	   recSet.MoveFirst;

	   for(i=0; i<recSet.RecordCount; i++)
	   {
		   TICKETID = recSet.Fields('TICKET_ID').value;
		   
		   DupFlag = 'n';
		   ExitFlag = 'n';
		   recSet2.MoveFirst;

		   while(recSet2.EOF == 0)
		   {
			 
			   if(recSet2.EOF != 0)
				{
				   ExitFlag = 'y';
				}

			   if(TICKETID == recSet2.Fields('TICKET_ID').value)
			   {
				   DupFlag = 'y';
			   }

			   if(recSet2.EOF == 0)
			   {
				   recSet2.MoveNext;
			   }
		   }

		   if(DupFlag == 'n')
		   {
		   NumOfTickets++
			   SyncToBeWork()	
			   
		   }
		   recSet.MoveNext;
	   }
	   recSet.close
	   recSet2.close
	   cObj.close

	   TICKETID = WorkingCount;
	   UserAction='WorkingCOunt';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = OrphansCount;
	   UserAction='OrphansCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = DenverCount;
	   UserAction='DenverCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)
	   
	   TICKETID = DoubleCount;
	   UserAction='DoubleCount';
	   EmpNum=document.frmmain.txt_EmployeeNum.value;
	   WriteToActionLog(UserAction, EmpNum, TICKETID)   
	   
	   alert('Import Complete');
	   showStatus('Import Complete');
	}

/* ************ */
/* ?? FUNCTIONS */
/* ************ */
	function WriteCommentHistory(EmpNum, TICKETID, Comments)
	{
		window.grStr="SELECT * FROM Comment_History";
		
		OpenTable()
		
		recSet.MoveFirst;
		recSet.AddNew;
			recSet.Fields('EmpNum').value = EmpNum;
			recSet.Fields('Date').value = Date();
			recSet.Fields('TicketID').value = TICKETID;
			recSet.Fields('Comments').value = Comments;
		recSet.Update;
		
		recSet.close

	}

	function DisplayActions_Log()
	{
		div_logincheck.style.visibility="hidden";
		div_main.style.visibility="visible";
		
		var TicketID;
		var ActionsLog;
		var rec_Count;
		
		window.grStr3="SELECT TOP 100 * FROM DefaultTicketStats ORDER BY recNum DESC";
		
		ActionsLog = window.open("pub/templates/ActionsLog.html","actionslog",'width=500,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
		
		OpenDatabase2()
		OpenTable3()
		
		recSet3.MoveFirst;
		
		ActionsLog.document.write('<html>\n');
		ActionsLog.document.write('	<head>\n');
		ActionsLog.document.write('		<!-- CSS STYLESHEET -->\n');
		ActionsLog.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
		ActionsLog.document.write('		<!-- JAVASCRIPT LINKS -->\n');
		ActionsLog.document.write('		<script type="text/javascript" src="T://TASC Members Folders//TASC Peter LeBlanc//default tickets//main.js"></script>\n');
		ActionsLog.document.write('		<style type="text/css">\n');
		ActionsLog.document.write('			body {\n');
		ActionsLog.document.write('				/*background-color: #D0D0D0;*/\n');
		ActionsLog.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
		ActionsLog.document.write('				background-repeat: repeat;\n');
		ActionsLog.document.write('				border-style: none;\n');
		ActionsLog.document.write('			}\n');
		ActionsLog.document.write('		</style>\n');
		ActionsLog.document.write('	</head>\n');
		ActionsLog.document.write('	<body>\n');
		ActionsLog.document.write('		<form name="frmactionlog">');
		ActionsLog.document.write('			<table width="100%" cellpadding="0" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		ActionsLog.document.write('			<tr style="color:#ffffff;" class="text_header">\n');
		ActionsLog.document.write('				<td>&nbsp;Actions Log</td>\n');
		ActionsLog.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		ActionsLog.document.write('			</tr>\n');
		ActionsLog.document.write('			<tr>\n');
		ActionsLog.document.write('				<td colspan="2">\n');
		ActionsLog.document.write('					<table width="100%" cellpadding="0" cellspacing="0" border="0" class="text_normal">\n');
		ActionsLog.document.write('			<tr>');
		ActionsLog.document.write('				<td>#</td>');
		ActionsLog.document.write('				<td>Employee #</td>');
		ActionsLog.document.write('				<td>Date</td>');
		ActionsLog.document.write('				<td>Action</td>');
		ActionsLog.document.write('				<td>Record #</td>');
		ActionsLog.document.write('				<td>Ticket ID</td>');
		ActionsLog.document.write('			</tr>');
		
		rec_Count = 0;
		recSet3.MoveNext
		
		while(!recSet3.eof)
		{
			rec_Count++;
			
			if (rec_Count%2)
			{
				ActionsLog.document.write('			<tr class="text_larger" style="background-color:#F8F8F8;">');
			}
			else
			{
				ActionsLog.document.write('			<tr class="text_larger">');
			}
			
			ActionsLog.document.write('				<td style="background-color:#F8F8F8;">'+ rec_Count +'</td>');
			ActionsLog.document.write('				<td>&nbsp;'+ recSet3.Fields("EmpNum").value +'</td>');
			ActionsLog.document.write('				<td>'+ recSet3.Fields("Date").value +'</td>');
			ActionsLog.document.write('				<td>'+ recSet3.Fields("Action").value +'</td>');
			ActionsLog.document.write('				<td style="text-align:center;">'+ recSet3.Fields("RecNum").value +'</td>');
			ActionsLog.document.write('				<td>'+ recSet3.Fields("TICKETID").value +'</td>');
			ActionsLog.document.write('			</tr>');
			
			recSet3.MoveNext
		}

		ActionsLog.document.write('					</table>\n');
		ActionsLog.document.write('				</td>\n');	
		ActionsLog.document.write('			</tr>\n');
		ActionsLog.document.write('			<tr class="text_header text_normal" style="text-align:right">\n');
		ActionsLog.document.write('				<td colspan="2">Found '+rec_Count+' records</td>\n');	
		ActionsLog.document.write('			</tr>\n');
		ActionsLog.document.write('			</table>\n');
		ActionsLog.document.write('		</form>\n');
		ActionsLog.document.write('	</body>\n');
		ActionsLog.document.write('</html>');
		
		recSet3.close
		recSet3 = null;
		cObj2.close
	}

	function DisplayCommentHistory()
	{
		var TicketID;
		var CommentHistory;
		var rec_Count;
		
		window.grStr="SELECT * FROM Comment_History ORDER BY DATE ASC";
		
		CommentHistory = window.open("pub/templates/CommentHistory.html","commenthistory",'width=500,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
		
		OpenDatabase()
		OpenTable()
		
		TicketID = document.frmmain.txt_TicketID.value; 
		
		recSet.MoveFirst;
		
		CommentHistory.document.write('<html>\n');
		CommentHistory.document.write('	<head>\n');
		CommentHistory.document.write('		<!-- CSS STYLESHEET -->\n');
		CommentHistory.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
		CommentHistory.document.write('		<!-- JAVASCRIPT LINKS -->\n');
		CommentHistory.document.write('		<script type="text/javascript" src="T://TASC Members Folders//TASC Peter LeBlanc//default tickets//main.js"></script>\n');
		CommentHistory.document.write('		<style type="text/css">\n');
		CommentHistory.document.write('			body {\n');
		CommentHistory.document.write('				/*background-color: #D0D0D0;*/\n');
		CommentHistory.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
		CommentHistory.document.write('				background-repeat: repeat;\n');
		CommentHistory.document.write('				border-style: none;\n');
		CommentHistory.document.write('			}\n');
		CommentHistory.document.write('		</style>\n');
		CommentHistory.document.write('	</head>\n');
		CommentHistory.document.write('	<body>\n');
		CommentHistory.document.write('		<form name="frmcomments">');
		CommentHistory.document.write('			<table width="100%" cellpadding="0" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		CommentHistory.document.write('			<tr style="color:#ffffff;" class="text_header">\n');
		CommentHistory.document.write('				<td>&nbsp;Comment History -- Ticket#' + TicketID + '</td>\n');
		CommentHistory.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		CommentHistory.document.write('			</tr>\n');
		CommentHistory.document.write('			<tr>\n');
		CommentHistory.document.write('				<td colspan="2">\n');
		CommentHistory.document.write('					<table width="100%" cellpadding="0" cellspacing="0" border="0" class="text_normal">\n');
		
		rec_Count = 0;
		
		while(!recSet.eof)
		{
			if(recSet.Fields('TicketID').value == TicketID)
			{
				rec_Count++;
				
				CommentHistory.document.write('			<tr style="background-color:#F8F8F8;">');
				CommentHistory.document.write("				<td style=\"background-image: url(pub/images/headers/header_bg.jpg); background-repeat: repeat-x;\" width=\"10%\">&nbsp;#<u>" + rec_Count + "</u> -- " + recSet.Fields('Date').value + "&nbsp;(<label for=\"lblcommment_" + rec_Count + "\" accesskey=\"" + rec_Count + "\"><a href=\"#\" onClick=\"CopyTEXT(document.getElementById('txt_Comment_" + rec_Count + "').innerText);\">Copy Text</a></label>)</td><td style=\"text-align:right; background-image: url(pub/images/headers/header_bg.jpg); background-repeat: repeat-x;\" width=\"5%\">Last Worked by: " + recSet.Fields("EmpNum").value + "&nbsp;</td>");
				CommentHistory.document.write("			</tr>");
				CommentHistory.document.write('			<tr class="text_larger">');
				CommentHistory.document.write('				<td colspan="2">&nbsp;<label id="lblcommment_' + rec_Count + '" name="lblcommment_' + rec_Count + '">' + recSet.Fields("Comments").value +'</label></td>');
				CommentHistory.document.write('			</tr><tr><td>&nbsp;</td></tr>\n');
			}
			recSet.MoveNext
		}

		CommentHistory.document.write('					</table>\n');
		CommentHistory.document.write('				</td>\n');	
		CommentHistory.document.write('			</tr>\n');
		CommentHistory.document.write('			<tr class="text_header text_normal" style="text-align:right">\n');
		CommentHistory.document.write('				<td colspan="2">Found '+rec_Count+' records</td>\n');	
		CommentHistory.document.write('			</tr>\n');
		CommentHistory.document.write('			</table>\n');
		CommentHistory.document.write('		</form>\n');
		CommentHistory.document.write('	</body>\n');
		CommentHistory.document.write('</html>');

	}

	function WriteToActionLog(UserAction, EmpNum, TICKETID)
	{
		var myDay;
		var myMonth;
		var myYear;
		var TodaysDate;
		var TodaysTime;

		var myDate=new Date();
		EmpNum = document.frmmain.txt_EmployeeNum.value;
	//	TICKETID = document.frmmain.txt_TicketID.value;
	  
		myDay = myDate.getDate();
		myMonth = myDate.getMonth();
		myMonth = myMonth + 1;
		myYear = myDate.getYear();
		TodaysDate = myMonth + "/" + myDay + "/" + myYear;
		TodaysTime = myDate.getHours()
		
		window.grStr3="SELECT TOP 1 * FROM DefaultTicketStats";
		OpenDatabase2()
		OpenTable3()

		recSet3.AddNew
			recSet3.Fields('EmpNum').value = EmpNum;
			recSet3.Fields('Action').value = UserAction;
			recSet3.Fields('Date').value = TodaysDate;
			recSet3.Fields('Time').value = TodaysTime;
			recSet3.Fields('TICKETID').value = TICKETID;
		recSet3.Update

		recSet3.close
		cObj2.close
	}

	function SwitchNewWorked()
	{
		var info = document.frmmain;
		
		if(info.cmd_WorkedLockedTickets.value == 'Work New Tickets')
		{
			document.getElementById('lbl_Header').innerText = 'New Tickets';
			info.cmd_WorkedLockedTickets.value = 'Work Locked Tickets';
			PullRecord()
		}
		else
		{
			document.getElementById('lbl_Header').innerText = 'Locked Tickets';
			info.cmd_WorkedLockedTickets.value = 'Work New Tickets';
			PullRecord()
		}
	}

	function DisplayAdminOptions()
	{
		div_logincheck.style.visibility="hidden";
		div_adminscreen.style.visibility="visible";
	}

	function CloseAdminOptions()
	{
		div_adminscreen.style.visibility="hidden";
		div_main.style.visibility="visible";
	}

	function AdminUnlockTicket()
	{
		var TICKETID;
		var EmpNum;
		
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = document.frmAdmin.txt_TicketID.value;
		
		window.grStr2="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TICKETID + "'" + " AND BEING_WORK_BY='" + EmpNum + "'";

		OpenDatabase()
		OpenTable2()
			
		recSet2.WillChangeRecord
			recSet2.Fields('WORK_TIME').value = 1;
		recSet2.Update
		
		recSet2.close
		cObj.close
		
		var sMsg = 'Ticket#'+TICKETID+' Released';
		alert(sMsg)
		
		with (document.frmmain) 
		{
			txt_TickettoRelease.value = "";
			txt_TickettoRelease.focus()
		}
	}

	function UnlockTicket()
	{
		var TICKETID;
		var EmpNum;
		var ReleaseTime;
		var ReleaseDate;
		
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		TICKETID = document.frmmain.txt_TickettoRelease.value;
		ReleaseTime = document.frmmain.txt_TimetoRelease.value;
		ReleaseDate = document.frmmain.txt_Date.value 
		
		window.grStr2="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TICKETID + "'";

		OpenDatabase()
		OpenTable2()
		
		if(ReleaseDate != "")
		{
			recSet2.WillChangeRecord
				recSet2.Fields('RELEASE_DATE').value = ReleaseDate;
			recSet2.Update
		}
		
		if(ReleaseTime == "")
		{
			recSet2.WillChangeRecord
				recSet2.Fields('WORK_TIME').value = 1;
			recSet2.Update
		}
		else
		{
			recSet2.WillChangeRecord
				recSet2.Fields('WORK_TIME').value = ReleaseTime - 8;
			recSet2.Update
		}
		
		recSet2.close
		cObj.close
		
		var sMsg = 'Ticket# ' + TICKETID + ' Released';
		alert(sMsg)
		
		with (document.frmmain)
		{
			txt_TickettoRelease.value = "";
			txt_TimetoRelease.value = " ";
			txt_TickettoRelease.focus()
		}
	}

	function UnLockAgentTickets()
	{
		var PassTicketDate;
		var HoldTimeFlag;
		
		window.grStr2="SELECT * FROM TO_BE_WORKED WHERE ASSIGNED_TIME<>null";
		
		
		OpenDatabase()
		OpenTable2()
		
		recSet2.MoveFirst;
		
		while(recSet2.EOF == 0)
		{
			PassTicketDate = recSet2.Fields('ASSIGNED_TIME').value;
			HoldTimeFlag = CheckLockTime(PassTicketDate)		
			
			if(HoldTimeFlag == 'y')
			{
				recSet2.WillChangeRecord
					recSet2.Fields('WORK_TIME').value = 1;
					recSet2.Fields('BEING_WORK_BY').value = 'z';
				recSet2.Update
			}
			recSet2.MoveNext
		}
		
		recSet2.close
		cObj.close
		
		showStatus('All Tickets are unlocked')
		alert('All Tickets are unlocked');
	}

	function CheckLockTime(PassTicketDate)
	{
		
		var TicketDate;
		var TodaysDate;
		var TempMonth;
		
		var ticketMonth;
		var ticketDay;
		var todaysMonth;
		var todaysDay;
		
		var HoldTimeFlag;
		
		TodaysDate = new Date();
		todaysMonth = TodaysDate.getMonth()
		todaysDay = TodaysDate.getDate()
				
		TicketDate = PassTicketDate;
		ticketDay = TicketDate.slice(8,11)
		TempMonth = TicketDate.slice(4,7)
		
		switch(TempMonth)
		{
			case 'Jan':
				ticketMonth = '0';
				break
			case 'Feb':
				ticketMonth = '1';
				break
			case 'Mar':
				ticketMonth = '2';
				break
			case 'Apr':
				ticketMonth = '3';
				break
			case 'May':
				ticketMonth = '4';
				break
			case 'Jun':
				ticketMonth = '5';
				break
			case 'Jul':
				ticketMonth = '6';
				break
			case 'Aug':
				ticketMonth = '7';
				break
			case 'Sep':
				ticketMonth = '8';
				break
			case 'Oct':
				ticketMonth = '9';
				break
			case 'Nov':
				ticketMonth = '10';
				break
			case 'Dec':
				ticketMonth = '11';
				break
		
			default:
		}

		
		HoldTimeFlag = 'n';

		if(ticketMonth == todaysMonth)
		{
			ticketDay++;
			ticketDay++;
			
			if(ticketDay <= todaysDay)
			{
				HoldTimeFlag = 'y';

			}
		}


		return HoldTimeFlag;
	}

/* ********************** */
/* REPORT ISSUE FUNCTIONS */
/* ********************** */
	function Display_ReportIssue_form()
	{
		div_reportissue.style.visibility="visible";
		div_main.style.visibility="hidden";
		
		Reset_IssueForm()
	}
	
	function Reset_IssueForm()
	{
		UpdateIssueProbDD('n/a')
		document.frmReportIssue.cbo_issueprob.focus()
	}
	
	function Cancel_IssueForm()
	{
		div_reportissue.style.visibility="hidden";
		div_main.style.visibility="visible";
	}
	
	function UpdateIssueProbDD(value)
	{		
		with (document.frmReportIssue)
		{
			switch(value)
			{
				case 'n/a':
					txt_IssueOther.value="Choose Other from List to Access";
					txt_issuedetails.value="Please Choose a Reason before typing in this box.";
					cbo_issueprob.value="n/a";
					
					txt_IssueOther.disabled=true;
					txt_issuedetails.disabled=true;
					
					break
				case 'Other':
					txt_IssueOther.value="";
					txt_issuedetails.value="";
					
					txt_IssueOther.disabled=false;
					txt_issuedetails.disabled=false;
					
					txt_IssueOther.focus()
					
					break
				case 'Locked Tickets show a Neg Value':
					txt_IssueOther.disabled=true;
					txt_IssueOther.value="Choose Other from List to Access";
					txt_issuedetails.value="";
					txt_issuedetails.disabled=false;
					txt_issuedetails.focus()
					
					break
				case 'Was unable to save, getting an erorr message':
					txt_IssueOther.disabled=true;
					txt_IssueOther.value="Choose Other from List to Access";
					txt_issuedetails.value="";
					txt_issuedetails.disabled=false;
					txt_issuedetails.focus()
					
					break
				case 'Trouble Accessing my Locked Tickets':
					txt_IssueOther.disabled=true;
					txt_IssueOther.value="Choose Other from List to Access";
					txt_issuedetails.value="";
					txt_issuedetails.disabled=false;
					txt_issuedetails.focus()
					
					break
				case 'I keep getting the same ticket over and over.':
					txt_IssueOther.disabled=true;
					txt_IssueOther.value="Choose Other from List to Access";
					txt_issuedetails.value="";
					txt_issuedetails.disabled=false;
					txt_issuedetails.focus()
					
					break
			}
		}
	}

	function AssignTicketToAgent()
	{
		
	
		var TICKETID;
		var EmpNum;
		var TodaysDate;
		
		TodaysDate = new Date();
		
		EmpNum = document.frmAdmin.txt_EmpNum.value;
		TICKETID = document.frmAdmin.txt_TicketID.value;
		
		if(EmpNum == "" && TICKETID == "")
		{
			if(EmpNum == "")
			{
				alert('You must have a Employee ID')
			}
			if(TICKETID == "")
			{
				alert('You must have a Ticket ID')
			}
		}
		else
		{
			window.grStr="SELECT * FROM TO_BE_WORKED WHERE TICKET_ID='" + TICKETID + "'";
			
			OpenDatabase()
			
			OpenTable()
			
			
			recSet.WillChangeRecord
				recSet.Fields('BEING_WORK_BY').value = EmpNum;
				recSet.Fields('ASSIGNED_TIME').value = Date();
				recSet.Fields('WORK_TIME').value = '1';
			recSet.Update
			recSet.close
			cObj.close
			
			UserAction = 'Assign Ticket';
			WriteToActionLog(UserAction, EmpNum, TICKETID)
			alert('Ticket Assigned')
		}
	}

	
	function submit_Issue()
	{
		//Assign Values
		EmpNum = document.frmmain.txt_EmployeeNum.value;
		Issue = document.frmReportIssue.cbo_issueprob.value;
		if (Issue == "Other") Issue = document.frmReportIssue.txt_IssueOther.value;
		Details = document.frmReportIssue.txt_issuedetails.value;
		
		//Insert into Database
		window.grStr3 = "SELECT * FROM tlb_Reported_Probs";
		
		OpenDatabase2()
		OpenTable3()
		
		recSet3.AddNew
		
		recSet3.Fields('EmpNum').value = EmpNum;
		recSet3.Fields('issue').value = Issue;
		recSet3.Fields('details').value = Details;
		
		recSet3.Update
		
		recSet3.close
		cObj2.close
		
		//Provide Feedback
		sMsg = 'Request has been send!';
		alert(sMsg);
		
		Cancel_IssueForm()
	}