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

// Global Variables

var UserStats;
var myRow;
var myCol;

function HeldTicketCount()
{

	var HeldCount;
	var myTotal;
	
	window.grStr2="SELECT * FROM TO_BE_WORKED";
	
	OpenDatabase()
	OpenTable2()
	
	if(recSet2.BOF == 0)
	{
		recSet2.MoveFirst
	}
	
	HeldCount = 0;
	
	while(recSet2.EOF == 0)
	{
		Workable = CheckWorkedDate()
		if(Workable == 'y')
		{
			HeldCount++;	
		}
		recSet2.MoveNext;
	}
	
	myTotal = recSet2.RecordCount;
	myTotal = myTotal-HeldCount;
	
	//Set recSet2 = Nothing;
	
	recSet2.close
	
	cObj.close

	alert(myTotal);
}




function LockedTicketCount()
{

	var myDate=new Date();
	var SearchValue;
		
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;

	var myDate=new Date();
 
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours()
	
	window.grStr="SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY IS NOT NULL AND BEING_WORK_BY <> 'z'";
	
	OpenDatabase()
	OpenTable()
	
	alert(recSet.RecordCount)
	
	//set recSet = Nothing
	
	recSet.close
	cObj.close

}

function ViewTicketAgents()
{

	var myDate=new Date();
	var SearchValue;
		
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;

	var myDate=new Date();
 
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours()
	
	window.grStr="SELECT DISTINCT(BEING_WORK_BY) FROM TO_BE_WORKED";
	
	OpenDatabase()
	OpenTable()
	
	if(recSet.BOF == 0)
	{
		recSet.MoveFirst
	}
	
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
	alert(recSet.RecordCount)
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
	UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			<tr>\n');
	UserStats.document.write('				<td>');
	UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
	UserStats.document.write('						<tr class="text_subheader">\n');
	UserStats.document.write('							<td>EmpNum</td>');
	UserStats.document.write('						</tr>\n');	
	
	
	while(recSet.EOF == 0)
	{
		UserStats.document.write('						<tr class="text_normal">\n');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('BEING_WORK_BY').value +'</td>');
		UserStats.document.write('						</tr>\n');	
		recSet.MoveNext	
	}
	
	
	UserStats.document.write('					</table>\n');
	UserStats.document.write('				</td>\n');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			</table>\n');
	UserStats.document.write('		</form>\n');

	UserStats.document.write('	</body>\n');

	UserStats.document.write('</html>');
	
	//Set recSet = Nothing
	
	recSet.close
	cObj.close
}

function ViewAgentComments()
{

	var ReportFrom;
	var ReportTo;
	var ReportOn;
		
	ReportOn = document.frmAdmin.txt_EmpNum.value;
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;

	window.grStr="SELECT TOP 300 * FROM Comment_History WHERE EmpNum='" + ReportOn + "' ORDER BY DATE ASC"; //AND DATE BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	
	OpenDatabase()
	OpenTable()

	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
	
	if(recSet.BOF == 0)
	{
		recSet.MoveFirst
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
	UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			<tr>\n');
	UserStats.document.write('				<td>');
	UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
	UserStats.document.write('						<tr class="text_subheader">\n');
	UserStats.document.write('							<td>EmpNum</td>');
	UserStats.document.write('							<td>Date</td>');
	UserStats.document.write('							<td>TicketID</td>');
	UserStats.document.write('							<td>Comments</td>');
	UserStats.document.write('						</tr>\n');	
	
	
	while(recSet.EOF == 0)
	{
		UserStats.document.write('						<tr class="text_normal">\n');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('EmpNum').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('Date').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('TicketID').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('Comments').value +'</td>');
		UserStats.document.write('						</tr>\n');	
		recSet.MoveNext	
	}
	
	
	UserStats.document.write('					</table>\n');
	UserStats.document.write('				</td>\n');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			</table>\n');
	UserStats.document.write('		</form>\n');

	UserStats.document.write('	</body>\n');

	UserStats.document.write('</html>');
	
	//Set recSet = Nothing
	
	recSet.close
	cObj.close

}


function GetFeedbackStats()
{
	var ReportFrom;
	var ReportTo;
		
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	
	//"Date BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	
	window.grStr="SELECT * FROM Completed WHERE INVALID_TICKET=true AND WORKED_DATE BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";

	
	OpenDatabase()
	OpenTable()
	
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
	
	if(recSet.BOF == 0)
	{
		recSet.MoveFirst
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
	UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			<tr>\n');
	UserStats.document.write('				<td>');
	UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
	UserStats.document.write('						<tr class="text_subheader">\n');
	UserStats.document.write('							<td>Ticket ID</td>');
	UserStats.document.write('							<td>Submitter</td>');
	UserStats.document.write('							<td>Resolution</td>');
	UserStats.document.write('							<td>Create Date</td>');
	UserStats.document.write('							<td>Invalid Problem</td>');
	UserStats.document.write('						</tr>\n');	
	UserStats.document.write('						<tr class="text_normal">\n');


	while(recSet.EOF == 0)
	{
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('TICKET_ID').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('SUBMITTER').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('RESOLUTION').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('CREATE_DATE').value +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ recSet.Fields('INVALID_PROBLEM').value +'</td>');
		recSet.MoveNext	
	}
	
	UserStats.document.write('						</tr>\n');
	UserStats.document.write('					</table>\n');
	UserStats.document.write('				</td>\n');
	UserStats.document.write('			</tr>\n');
	UserStats.document.write('			</table>\n');
	UserStats.document.write('		</form>\n');

	UserStats.document.write('	</body>\n');

	UserStats.document.write('</html>');
	
	//Set recSet = Nothing
	
	recSet.close
	cObj.close

}

function GetCompletedTotal()
{

	var myDate=new Date();
	var SearchValue;
		
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;

	var myDate=new Date();
 
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours()
	
	window.grStr="SELECT * FROM Completed WHERE WORKED_DATE='" + TodaysDate + "'";
	
	OpenDatabase()
	OpenTable()
	
	alert(recSet.RecordCount)
	
	//Set recSet = Nothing
	
	recSet.close
	cObj.close
	


}


function GetUserStats()
{
     var EmpNum;
	 var UserAction;
	 
     
     var AssignCount;
     var CompletedCount;
	 var CallbacksPublic;
	 var CallbacksSelf;
	 var Reassigned;
	 var NeverWorked;
	 var LockedCount;
	 var TotalCompleted;
	var ReportFrom;
	var ReportTo;
	var myConnectString;
	var NumOfRec;
	var CallbackDone;

	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;

	var myDate=new Date();
 
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours()

     
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
    EmpNum = document.frmmain.txt_EmployeeNum.value;
     
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE Date='" + TodaysDate +"' AND EmpNum='" + EmpNum + "'";

	window.grStr3=myConnectString;

     OpenDatabase2()
     OpenTable3()

     AssignCount = 0;
     CompletedCount = 0;
	 TotalCompleted = 0;
	 CallbacksPublic = 0;
	 CallbacksSelf = 0;
	 Reassigned = 0;
	 NeverWorked = 0;
	 LockedCount = 0;
	 CallbackDone = 0;
	 
	NumOfRec = recSet3.RecordCount;

	if(recSet3.BOF == 0)
	{
	     recSet3.MoveFirst;
	}

     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;

            switch(UserAction)
            {
            case 'Assign Ticket':
                AssignCount += 1;
       	        break
			
            case 'Completed':
                CompletedCount += 1;
				TotalCompleted += 1;
                break
				
			case 'ToSelf':
				CallbacksSelf += 1;
				break
				
			case 'ToPublic':
				CallbacksPublic += 1;
				TotalCompleted += 1;
				break
				
			case 'Reassigned':
				Reassigned += 1;
				TotalCompleted += 1;
				break
				
			case 'NeverWorked':
				NeverWorked += 1;
				break
			case 'DidCallBack':
				CallbackDone += 1;
				break
            default:
            }
        }

        recSet3.MoveNext;
     }

	
     LockedCount = AssignCount - TotalCompleted - NeverWorked;
	 
	 
     document.frmmain.txt_RecLocked.value = LockedCount;
     document.frmmain.txt_RecCompleted.value = TotalCompleted;

	

     recSet3.close
	 
     cObj2.close

	window.grStr2="SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY IS NULL OR BEING_WORK_BY = 'z'";
	OpenDatabase()
	OpenTable2()
		//document.frmmain.txt_RecRemaining.value = recSet2.RecordCount
		document.getElementById('lbl_RecRemaining').innerText = recSet2.RecordCount;
	recSet2.close
	window.grStr2="SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY='" + EmpNum + "' AND COMPLETED_BY IS NULL";
	OpenTable2()
		document.getElementById('lbl_RecLocked').innerText = recSet2.RecordCount;
	recSet2.close	
	cObj.close
}

function DisplayUserStats()
{
	var EmpNum;
	var UserAction;

	var RequestDate;
	var UserStats;

	var AssignCount;
	var CompletedCount;
	var CallbacksPublic;
	var CallbacksSelf;
	var Reassigned;
	var NeverWorked;
	var LockedCount;
	var TotalCompleted;
	var CallbackDone;

	var ReportFrom;
	var ReportTo;
	var myConnectString;
	
	var myDay;
	var myMonth;
	var myYear;
	var TodaysDate;
	var TodaysTime;

	var myDate=new Date();
 
	myDay = myDate.getDate();
	myMonth = myDate.getMonth();
	myMonth = myMonth + 1;
	myYear = myDate.getYear();
	TodaysDate = myMonth + "/" + myDay + "/" + myYear;
	TodaysTime = myDate.getHours()
	
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	EmpNum = document.frmmain.txt_EmployeeNum.value;
	

	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE EmpNum='" + EmpNum + "'" + " AND Date='" + TodaysDate + "'";
	window.grStr3=myConnectString;
    
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=950,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

	OpenDatabase2()
	OpenTable3()

	AssignCount = 0;
	CompletedCount = 0;
	TotalCompleted =0;
	CallbacksPublic = 0;
	CallbacksSelf = 0;
	Reassigned = 0;
	NeverWorked = 0;
	LockedCount = 0;
	CallbackDone = 0;
	
     recSet3.MoveFirst;
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;

            switch(UserAction)
            {
            case 'Assign Ticket':
                AssignCount += 1;
                break
				
            case 'Completed':
                CompletedCount += 1;
				TotalCompleted += 1;
                break
				
			case 'ToSelf':
				CallbacksSelf += 1;
				break
				
			case 'ToPublic':
				CallbacksPublic += 1;
				TotalCompleted += 1;
				break
				
			case 'Reassigned':
				Reassigned += 1;
				TotalCompleted += 1;
				break
				
			case 'NeverWorked':
				NeverWorked += 1;
				break
			case 'DidCallBack':
				CallbackDone += 1;
				break
            default:
            }
        }
        recSet3.MoveNext;
     }

	if(AssignCount != 0)
	{
		LockedCount = AssignCount - TotalCompleted - NeverWorked;
		
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
		UserStats.document.write('				<td>&nbsp;'+ EmpNum +'</td>\n');
		UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			<tr>\n');
		UserStats.document.write('				<td>');
		UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('						<tr class="text_subheader">\n');
		UserStats.document.write('							<td>Assign Ticket</td>');
		UserStats.document.write('							<td>Resolved</td>');
		UserStats.document.write('							<td>Callback Done</td>');
		UserStats.document.write('							<td>To Self</td>');
		UserStats.document.write('							<td>ToPublic</td>');
		UserStats.document.write('							<td>Reassigned</td>');
		UserStats.document.write('							<td>NeverWorked</td>');
		UserStats.document.write('							<td>Total Completed</td>');
		UserStats.document.write('						</tr>\n');	
		UserStats.document.write('						<tr class="text_normal">\n');
		UserStats.document.write('							<td>&nbsp;'+ LockedCount +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CompletedCount +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbackDone +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbacksSelf +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbacksPublic +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ Reassigned +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ NeverWorked +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ TotalCompleted +'</td>');
		UserStats.document.write('						</tr>\n');
		UserStats.document.write('					</table>\n');
		UserStats.document.write('				</td>\n');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			</table>\n');
		UserStats.document.write('		</form>\n');

		UserStats.document.write('	</body>\n');

		UserStats.document.write('</html>');
	}
	
	//Set recSet3 = Nothing
	
     recSet3.close
     cObj2.close

    
}

function AdminGetUserStats(ReportOn)
{
	var EmpNum;
	var UserAction;

	var RequestDate;
	var UserStats;

	var AssignCount;
	var CompletedCount;
	var CallbacksPublic;
	var CallbacksSelf;
	var Reassigned;
	var NeverWorked;
	var LockedCount;
	var TotalCompleted;
	var CallbackDone;

	var ReportFrom;
	var ReportTo;
	var myConnectString;
	
	ReportOn = document.frmAdmin.txt_EmpNum.value;
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	EmpNum = ReportOn
	

	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE Date BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	window.grStr3=myConnectString;
    
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=950,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

	OpenDatabase2()
	OpenTable3()

	AssignCount = 0;
	CompletedCount = 0;
	TotalCompleted =0;
	CallbacksPublic = 0;
	CallbacksSelf = 0;
	Reassigned = 0;
	NeverWorked = 0;
	LockedCount = 0;
	CallbackDone = 0;
	
     recSet3.MoveFirst;
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;

            switch(UserAction)
            {
            case 'Assign Ticket':
                AssignCount += 1;
                break
				
            case 'Completed':
                CompletedCount += 1;
				TotalCompleted += 1;
                break
				
			case 'ToSelf':
				CallbacksSelf += 1;
				break
				
			case 'ToPublic':
				CallbacksPublic += 1;
				TotalCompleted += 1;
				break
				
			case 'Reassigned':
				Reassigned += 1;
				TotalCompleted += 1;
				break
				
			case 'NeverWorked':
				NeverWorked += 1;
				break
			case 'DidCallBack':
				CallbackDone += 1;
				break
            default:
            }
        }
        recSet3.MoveNext;
     }

	if(AssignCount != 0)
	{
		LockedCount = AssignCount - TotalCompleted - NeverWorked;
		
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
		UserStats.document.write('				<td>&nbsp;'+ EmpNum +'</td>\n');
		UserStats.document.write('				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			<tr>\n');
		UserStats.document.write('				<td>');
		UserStats.document.write('					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('						<tr class="text_subheader">\n');
		UserStats.document.write('							<td>Assign Ticket</td>');
		UserStats.document.write('							<td>Completed</td>');
		UserStats.document.write('							<td>Callback Done</td>');
		UserStats.document.write('							<td>To Self</td>');
		UserStats.document.write('							<td>ToPublic</td>');
		UserStats.document.write('							<td>Reassigned</td>');
		UserStats.document.write('							<td>NeverWorked</td>');
		UserStats.document.write('							<td>Total Completed</td>');
		UserStats.document.write('						</tr>\n');	
		UserStats.document.write('						<tr class="text_normal">\n');
		UserStats.document.write('							<td>&nbsp;'+ LockedCount +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CompletedCount +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbackDone +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbacksSelf +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ CallbacksPublic +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ Reassigned +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ NeverWorked +'</td>');
		UserStats.document.write('							<td>&nbsp;'+ TotalCompleted +'</td>');
		UserStats.document.write('						</tr>\n');
		UserStats.document.write('					</table>\n');
		UserStats.document.write('				</td>\n');
		UserStats.document.write('			</tr>\n');
		UserStats.document.write('			</table>\n');
		UserStats.document.write('		</form>\n');

		UserStats.document.write('	</body>\n');

		UserStats.document.write('</html>');
	}
	
	//Set recSet3 = Nothing
	
     recSet3.close
     cObj2.close

    
}


function DefaultTotal()
{

    var EmpList=new Array()
	var EmpName=new Array()
	
	var UserCount;
	var i;
	
	var ReportName;
	var ReportOn;
	var ReportFrom;
	var ReportTo;
	var myConnectString;

	ReportOn = document.frmAdmin.txt_EmpNum.value;
	ReportFrom = document.frmAdmin.txt_From.value;
		  
    window.grStr3="SELECT * FROM UserLogins";
	
    OpenDatabase2()
    OpenTable3()

	recSet3.MoveFirst;
	UserCount = 0;
	
	
	while(recSet3.EOF == 0)
	{
		EmpList[UserCount] = recSet3.Fields('EmpNum').value;
		EmpName[UserCount] = recSet3.Fields('Name').value;
		UserCount++;
		recSet3.MoveNext
	}
	
	//Set recSet3 = Nothing
	
	recSet3.Close
	cObj2.close
	
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
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

	for(i=0; i<=UserCount; i++)
	{
		ReportOn = EmpList[i];
		ReportName = EmpName[i];
		ReportAllUsers(ReportOn,ReportName)
	}
		
	UserStats.document.write('	</body>\n');
	UserStats.document.write('</html>');
	
	
}



function ReportAllUsers(ReportOn, ReportName)
{
	var EmpNum;
	var UserAction;

	var RequestDate;

	var AssignCount;
	var CompletedCount;
	var CallbacksPublic;
	var CallbacksSelf;
	var Reassigned;
	var NeverWorked;
	var LockedCount;
	var TotalCompleted;
	var CallbackDone;

	var ReportFrom;
	var ReportTo;
	var myConnectString;

	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	
	EmpNum = ReportOn;
		
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE " + "Date BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	window.grStr3=myConnectString;


	//UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

	
	OpenDatabase2()
	OpenTable3()

	AssignCount = 0;
	CompletedCount = 0;
	TotalCompleted = 0;
	CallbacksPublic = 0;
	CallbacksSelf = 0;
	Reassigned = 0;
	NeverWorked = 0;
	LockedCount = 0;
	CallbackDone = 0;
	 
     recSet3.MoveFirst;
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;

            switch(UserAction)
            {
            case 'Assign Ticket':
                AssignCount++;
                break
				
            case 'Completed':
                CompletedCount++;
				TotalCompleted++;
                break
				
			case 'ToSelf':
				CallbacksSelf++;
				break
				
			case 'ToPublic':
				TotalCompleted++;
				CallbacksPublic++;
				break
				
			case 'Reassigned':
				Reassigned++;
				TotalCompleted++;
				break
				
			case 'NeverWorked':
				NeverWorked++;
				break
			case 'DidCallBack':
				CallbackDone++;
				break
            default:
            }
        }
        recSet3.MoveNext;
     }
	
	if(AssignCount != 0)
	{
		LockedCount = AssignCount - TotalCompleted - NeverWorked;

		UserStats.document.write('			<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('				<tr style="color:#ffffff;" class="text_header">\n');
		UserStats.document.write('					<td>&nbsp;'+ EmpNum +' -- ' + ReportName + '</td>\n');
		UserStats.document.write('				</tr>\n');
		UserStats.document.write('				<tr>\n');
		UserStats.document.write('					<td>');
		
		UserStats.document.write('						<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n');
		UserStats.document.write('							<tr class="text_subheader">\n');
		UserStats.document.write('								<td>Assign Ticket</td>');
		UserStats.document.write('								<td>Resolved</td>');
		UserStats.document.write('								<td>Callbacks Done</td>');
		UserStats.document.write('								<td>To Self</td>');
		UserStats.document.write('								<td>ToPublic</td>');
		UserStats.document.write('								<td>Reassigned</td>');
		UserStats.document.write('								<td>NeverWorked</td>');
		UserStats.document.write('								<td>Total Completed</td>');
		UserStats.document.write('							</tr>\n');	
		UserStats.document.write('							<tr class="text_normal">\n');
		UserStats.document.write('								<td>&nbsp;'+ LockedCount +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ CompletedCount +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ CallbackDone +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ CallbacksSelf +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ CallbacksPublic +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ Reassigned +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ NeverWorked +'</td>');
		UserStats.document.write('								<td>&nbsp;'+ TotalCompleted +'</td>');
		UserStats.document.write('							</tr>\n');
		UserStats.document.write('						</table>\n');

		UserStats.document.write('					</td>\n');
		UserStats.document.write('				</tr>\n');
		UserStats.document.write('			</table>\n');
	}
	
	//Set recSet3 = Nothing
	
	recSet3.close
	cObj2.close

}

function DefaultTotalSpreadsheet()
{

    var EmpList=new Array()
	var EmpName=new Array()
	
	var UserCount;
	var i;
	
	var ReportName;
	var ReportOn;
	var ReportFrom;
	var ReportTo;
	var myDate=new Date();
	var myConnectString;

	ReportOn = document.frmAdmin.txt_EmpNum.value;
	ReportFrom = document.frmAdmin.txt_From.value;
		  
    window.grStr3="SELECT * FROM UserLogins";
	
    OpenDatabase2()
    OpenTable3()

	recSet3.MoveFirst;
	UserCount = 0;
	
	
	while(recSet3.EOF == 0)
	{
		EmpList[UserCount] = recSet3.Fields('EmpNum').value;
		EmpName[UserCount] = recSet3.Fields('Name').value;
		UserCount++;
		recSet3.MoveNext
	}
	
	
	recSet3.Close
	recSet3 = null
	cObj2.close
	
	alert("Creating Spreadsheet")

	var myRow;
	var myCol;

	var excel = new ActiveXObject ("Excel.Application");
	excel.visible = true;
	var book = excel.Workbooks.Add ();
	var sheet = excel.Worksheets(1);

	myRow = 1;
	myCol = 1;
	
	
	sheet.Columns(1).columnwidth = 20;
	sheet.Columns(2).columnwidth = 14;
	sheet.Columns(3).columnwidth = 14;
	sheet.Columns(4).columnwidth = 14;
	sheet.Columns(5).columnwidth = 14;
	sheet.Columns(6).columnwidth = 14;
	sheet.Columns(7).columnwidth = 14;
	sheet.Columns(8).columnwidth = 14;
	
	sheet.Cells(myRow,1) = 'Report Generated On: ' + myDate;
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,1).Font.Size = 14;
	sheet.Cells(myRow,1).Font.ColorIndex = 3;
	myRow++;
	myRow++;
	myRow++;
	
	for(i=0; i<=UserCount; i++)
	{
		ReportOn = EmpList[i];
		ReportName = EmpName[i];
		myRow=CreateAgentSpreadSheet(ReportOn, ReportName, sheet, book, excel, myRow)
	}
		

}

function CreateAgentSpreadSheet(ReportOn, ReportName, sheet, book, excel, myRow)
{

	var EmpNum;
	var UserAction;

	var RequestDate;

	var AssignCount;
	var CompletedCount;
	var CallbacksPublic;
	var CallbacksSelf;
	var Reassigned;
	var NeverWorked;
	var LockedCount;
	var TotalCompleted;
	var CallbackDone;

	var ReportFrom;
	var ReportTo;
	var myConnectString;

	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	
	EmpNum = ReportOn;
		
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE " + "Date BETWEEN '" + document.frmAdmin.txt_From.value + "' AND '" + document.frmAdmin.txt_To.value + "' AND EmpNum='" + EmpNum + "'";
	window.grStr3=myConnectString;

	
	//UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

	
	OpenDatabase2()
	OpenTable3()

	AssignCount = 0;
	CompletedCount = 0;
	TotalCompleted = 0;
	CallbacksPublic = 0;
	CallbacksSelf = 0;
	Reassigned = 0;
	NeverWorked = 0;
	LockedCount = 0;
	CallbackDone = 0;
	 
	if(recSet3.BOF == 0)
	{
		recSet3.MoveFirst
	}
	 
     
     while(recSet3.EOF == 0)
     {
		UserAction = recSet3.Fields('Action').value;

		switch(UserAction)
		{
		case 'Assign Ticket':
			AssignCount++;
			break
			
		case 'Completed':
			CompletedCount++;
			TotalCompleted++;
			break
			
		case 'ToSelf':
			CallbacksSelf++;
			break
			
		case 'ToPublic':
			TotalCompleted++;
			CallbacksPublic++;
			break
			
		case 'Reassigned':
			Reassigned++;
			TotalCompleted++;
			break
			
		case 'NeverWorked':
			NeverWorked++;
			break
		case 'DidCallBack':
			CallbackDone++;
			break
		default:
		}
        recSet3.MoveNext;
     }
	 
	
	if(AssignCount != 0)
	{
		sheet.Cells(myRow,1) = EmpNum;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		sheet.Cells(myRow,1) = ReportName;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		sheet.Cells(myRow,1) = 'Assign Ticket';
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,2) = 'Resolved';
		sheet.Cells(myRow,2).Font.Bold = 'TRUE';
		sheet.Cells(myRow,3) = 'Callbacks Done';
		sheet.Cells(myRow,3).Font.Bold = 'TRUE';
		sheet.Cells(myRow,4) = 'To Self';
		sheet.Cells(myRow,4).Font.Bold = 'TRUE';
		sheet.Cells(myRow,5) = 'To Public';
		sheet.Cells(myRow,5).Font.Bold = 'TRUE';
		sheet.Cells(myRow,6) = 'Reassigned';
		sheet.Cells(myRow,6).Font.Bold = 'TRUE';
		sheet.Cells(myRow,7) = 'Never Worked';
		sheet.Cells(myRow,7).Font.Bold = 'TRUE';
		sheet.Cells(myRow,8) = 'Total Completed';
		sheet.Cells(myRow,8).Font.Bold = 'TRUE';
		myRow++
		sheet.Cells(myRow,1) = AssignCount;
		sheet.Cells(myRow,2) = CompletedCount;
		sheet.Cells(myRow,3) = CallbackDone;
		sheet.Cells(myRow,4) = CallbacksSelf;
		sheet.Cells(myRow,5) = CallbacksPublic;
		sheet.Cells(myRow,6) = Reassigned;
		sheet.Cells(myRow,7) = NeverWorked;
		sheet.Cells(myRow,8) = TotalCompleted;
		myRow++ 
		myRow++
	}
	
	
	recSet3.close
	recSet3 = null;
	cObj2.close

	return myRow;

}

function AgentDaybyDay()
{

	var EmpNum;
	var UserAction;

	var RequestDate;

	var AssignCount;
	var CompletedCount;
	var CallbacksPublic;
	var CallbacksSelf;
	var Reassigned;
	var NeverWorked;
	var LockedCount;
	var TotalCompleted;
	var CallbackDone;

	var ReportFrom;
	var ReportTo;
	var myConnectString;

	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	ReportOn = document.frmAdmin.txt_EmpNum.value;
	
	EmpNum = ReportOn;
		
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE " + "Date BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"' AND EmpNum='" + EmpNum + "'";
	window.grStr3=myConnectString;

	
	//UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

	
	OpenDatabase2()
	OpenTable3()

	AssignCount = 0;
	CompletedCount = 0;
	TotalCompleted = 0;
	CallbacksPublic = 0;
	CallbacksSelf = 0;
	Reassigned = 0;
	NeverWorked = 0;
	LockedCount = 0;
	CallbackDone = 0;
	 
	if(recSet3.BOF == 0)
	{
		recSet3.MoveFirst
	}
	 
     
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;

            switch(UserAction)
            {
            case 'Assign Ticket':
                AssignCount++;
                break
				
            case 'Completed':
                CompletedCount++;
				TotalCompleted++;
                break
				
			case 'ToSelf':
				CallbacksSelf++;
				break
				
			case 'ToPublic':
				TotalCompleted++;
				CallbacksPublic++;
				break
				
			case 'Reassigned':
				Reassigned++;
				TotalCompleted++;
				break
				
			case 'NeverWorked':
				NeverWorked++;
				break
			case 'DidCallBack':
				CallbackDone++;
				break
            default:
            }
        }
        recSet3.MoveNext;
     }
	 
	
	if(AssignCount != 0)
	{
		sheet.Cells(myRow,1) = EmpNum;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		sheet.Cells(myRow,1) = ReportName;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		sheet.Cells(myRow,1) = 'Assign Ticket';
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,2) = 'Resolved';
		sheet.Cells(myRow,2).Font.Bold = 'TRUE';
		sheet.Cells(myRow,3) = 'Callbacks Done';
		sheet.Cells(myRow,3).Font.Bold = 'TRUE';
		sheet.Cells(myRow,4) = 'To Self';
		sheet.Cells(myRow,4).Font.Bold = 'TRUE';
		sheet.Cells(myRow,5) = 'To Public';
		sheet.Cells(myRow,5).Font.Bold = 'TRUE';
		sheet.Cells(myRow,6) = 'Reassigned';
		sheet.Cells(myRow,6).Font.Bold = 'TRUE';
		sheet.Cells(myRow,7) = 'Never Worked';
		sheet.Cells(myRow,7).Font.Bold = 'TRUE';
		sheet.Cells(myRow,8) = 'Total Completed';
		sheet.Cells(myRow,8).Font.Bold = 'TRUE';
		myRow++
		sheet.Cells(myRow,1) = AssignCount;
		sheet.Cells(myRow,2) = CompletedCount;
		sheet.Cells(myRow,3) = CallbackDone;
		sheet.Cells(myRow,4) = CallbacksSelf;
		sheet.Cells(myRow,5) = CallbacksPublic;
		sheet.Cells(myRow,6) = Reassigned;
		sheet.Cells(myRow,7) = NeverWorked;
		sheet.Cells(myRow,8) = TotalCompleted;
		myRow++ 
		myRow++
	}
	
	//Set recSet3 = Nothing
	
	recSet3.close
	cObj2.close

	return myRow;

}

function GetFeedbackStatsSpreadsheet()
{
	var ReportFrom;
	var ReportTo;
		
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	
	alert("Creating Spreadsheet")

	var myRow;
	var myCol;

	var excel = new ActiveXObject ("Excel.Application");
	excel.visible = true;
	var book = excel.Workbooks.Add ();
	var sheet = excel.Worksheets(1);

	myRow = 1;
	myCol = 1;	
	
	window.grStr="SELECT * FROM Completed WHERE INVALID_TICKET=true AND WORKED_DATE BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	
	OpenDatabase()
	OpenTable()
	
	sheet.Cells(myRow,1) = 'Invalid Tickets';
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,1).Font.Size = 15;
	myRow++;
	myRow++;
	
	sheet.Columns(1).columnwidth = 10;
	sheet.Columns(2).columnwidth = 10;
	sheet.Columns(3).columnwidth = 25;
	sheet.Columns(4).columnwidth = 50;
	sheet.Columns(5).columnwidth = 15;
	
	if(recSet.BOF == 0)
	{
		recSet.MoveFirst
	}
		
	sheet.Cells(myRow,1) = 'TicketID';
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,2) = 'Submitter';
	sheet.Cells(myRow,2).Font.Bold = 'TRUE';
	sheet.Cells(myRow,3) = 'Invalid Problem';
	sheet.Cells(myRow,3).Font.Bold = 'TRUE';
	sheet.Cells(myRow,4) = 'Comments/Feedback';
	sheet.Cells(myRow,4).Font.Bold = 'TRUE';
	sheet.Cells(myRow,5) = 'Create Date';
	sheet.Cells(myRow,5).Font.Bold = 'TRUE';
	
	
	myRow++;
	while(recSet.EOF == 0)
	{
		sheet.Cells(myRow,1) = recSet.Fields('TICKET_ID').value;
		sheet.Cells(myRow,2) = recSet.Fields('SUBMITTER').value;
		sheet.Cells(myRow,3) = recSet.Fields('INVALID_PROBLEM').value;
		sheet.Cells(myRow,4) = recSet.Fields('RESOLUTION').value;
		sheet.Cells(myRow,5) = recSet.Fields('CREATE_DATE').value;
		
		myRow++;
		recSet.MoveNext	
	}
	
	//Set recSet = Nothing;
	
	recSet.close
	cObj.close

}


function GenerateSystemStats()
{
	/*
	//Setup Variables
	var totals_labels = array(
		'',
		'',
		'',
		'',
		'',
		'',
	);
	var totals_data = array()	
	var sqlquery_totaltickets;
	var sqlquery_totalusers;
	var sqlquery_systemstatus;
	
	//Gather Live data
	myConnectString = "SELECT * FROM DefaultTicketStats";
	
	OpenDatabase2()
	OpenTable3()
	
	recSet3.close
	cObj2.close
	
	//Generate Live Results	
	document.write('<tr>');
	document.write('	<td>Total Tickets in Database:</td>');
	document.write('	<td></td>');
	document.write('</tr>');
	document.write('<tr>');
	document.write('	<td>Total Users:</td>');
	document.write('	<td></td>');
	document.write('</tr>');
	document.write('<tr>');
	document.write('	<td>System Status:</td>');
	document.write('	<td></td>');
	document.write('</tr>');
	*/
}